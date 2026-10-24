/**
 * Google Apps Script API for the GitHub Pages insulin diary.
 *
 * Deploy as a web app that runs as the owner and is accessible to anyone with
 * the URL. The /exec URL is a bearer credential: keep it private.
 */
const SPREADSHEET_ID = '10HzH-zEU-aO1zhxWQqrTx93quhEtbF1nBT7pZuZAT04';
const SHEET_NAME = 'Инъекции';
const PUSH_SHEET_NAME = 'PushDevices';
const HEADERS = ['ID', 'Дата и время (ISO)', 'Доза (ЕД)', 'Сахар (ммоль/л)', 'Зона', 'Приём пищи', 'Примечание', 'Кто сделал запись'];
const PUSH_HEADERS = ['FCM token', 'Обновлено'];
const FCM_SERVICE_ACCOUNT_PROPERTY = 'FCM_SERVICE_ACCOUNT_JSON';
const LAST_PUSH_KEY_PROPERTY = 'LAST_SENT_INJECTION_KEY';
const FCM_ACCESS_TOKEN_CACHE_KEY = 'FCM_ACCESS_TOKEN';

function doGet(e) {
  if (!e || !e.parameter || e.parameter.action !== 'getEntries') {
    return jsonOutput_({ ok: false, error: 'Use action=getEntries.' });
  }
  const entries = getEntries();
  const callback = String(e.parameter.callback || '');
  if (callback) {
    if (!/^[A-Za-z_$][0-9A-Za-z_$]{0,100}$/.test(callback)) {
      return jsonOutput_({ ok: false, error: 'Invalid JSONP callback.' });
    }
    return ContentService.createTextOutput(`${callback}(${JSON.stringify(entries)});`)
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return jsonOutput_(entries);
}

function doPost(e) {
  try {
    const request = JSON.parse(e && e.postData && e.postData.contents || '{}');
    if (request.action === 'saveEntry') return jsonOutput_(saveEntry(request.record));
    if (request.action === 'registerPush') return jsonOutput_(registerPushDevice_(request.token));
    if (request.action === 'unregisterPush') return jsonOutput_(unregisterPushDevice_(request.token));
    return jsonOutput_({ ok: false, error: 'Unknown action.' });
  } catch (error) {
    console.error(error && error.stack || error);
    return jsonOutput_({ ok: false, error: error.message || 'Request failed.' });
  }
}

function setupTrackingSheet() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  } else {
    ensureAuthorColumn_(sheet);
  }
  ensurePushSheet_();
  return `Готово: лист «${SHEET_NAME}» подготовлен.`;
}

/**
 * Install once from the Apps Script editor. The trigger checks every five
 * minutes; Google may run time-driven triggers a little later than scheduled.
 */
function setupPushNotifications() {
  ensurePushSheet_();
  const exists = ScriptApp.getProjectTriggers().some(function (trigger) {
    return trigger.getHandlerFunction() === 'checkDueNotifications';
  });
  if (!exists) {
    ScriptApp.newTrigger('checkDueNotifications').timeBased().everyMinutes(5).create();
  }
  return 'Push-уведомления настроены: проверка таймера каждые 5 минут.';
}

function getEntries() {
  const sheet = getSheet_();
  const authorColumn = ensureAuthorColumn_(sheet);
  const rows = sheet.getLastRow() > 1
    ? sheet.getRange(2, 1, sheet.getLastRow() - 1, Math.max(HEADERS.length, authorColumn)).getValues()
    : [];
  return rows.map(function (row) {
    return {
      id: String(row[0]),
      injectedAt: normalizeDate_(row[1]),
      dose: Number(row[2]),
      glucose: row[3] === '' ? '' : Number(row[3]),
      site: String(row[4]),
      meal: String(row[5]),
      note: String(row[6] || ''),
      author: String(row[authorColumn - 1] || '')
    };
  }).filter(function (record) {
    return record.id && record.injectedAt;
  }).sort(function (a, b) {
    return new Date(b.injectedAt).getTime() - new Date(a.injectedAt).getTime();
  });
}

function saveEntry(record) {
  validateRecord_(record);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheet_();
    const authorColumn = ensureAuthorColumn_(sheet);
    const ids = sheet.getLastRow() > 1
      ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getDisplayValues().flat()
      : [];
    const values = [
      safeCell_(record.id),
      record.injectedAt,
      record.dose,
      record.glucose === '' || record.glucose == null ? '' : record.glucose,
      safeCell_(record.site),
      safeCell_(record.meal),
      safeCell_(record.note || '')
    ];
    const existingIndex = ids.indexOf(record.id);
    const rowNumber = existingIndex === -1 ? sheet.getLastRow() + 1 : existingIndex + 2;
    sheet.getRange(rowNumber, 1, 1, values.length).setValues([values]);
    sheet.getRange(rowNumber, authorColumn).setValue(safeCell_(record.author));
    return { ok: true, id: record.id, updated: existingIndex !== -1 };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Trigger entry point. Sends one push notification when the newest recorded
 * injection has been due for 12 hours. No medication schedule is inferred.
 */
function checkDueNotifications() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return;
  try {
    const devices = getPushSheet_();
    if (devices.getLastRow() < 2) return;

    const latest = getEntries()[0];
    if (!latest) return;
    const dueAt = new Date(latest.injectedAt).getTime() + 12 * 60 * 60 * 1000;
    if (!isFinite(dueAt) || Date.now() < dueAt) return;

    const injectionKey = `${latest.id}:${latest.injectedAt}`;
    const properties = PropertiesService.getScriptProperties();
    if (properties.getProperty(LAST_PUSH_KEY_PROPERTY) === injectionKey) return;

    const tokenRows = devices.getRange(2, 1, devices.getLastRow() - 1, 1).getDisplayValues()
      .map(function (row, index) { return { token: row[0], rowNumber: index + 2 }; })
      .filter(function (device) { return device.token; });
    if (!tokenRows.length) return;
    const staleRows = [];
    tokenRows.forEach(function (device) {
      if (!sendFcmNotification_(device.token, latest)) staleRows.push(device.rowNumber);
    });
    staleRows.reverse().forEach(function (rowNumber) {
      devices.deleteRow(rowNumber);
    });
    properties.setProperty(LAST_PUSH_KEY_PROPERTY, injectionKey);
  } finally {
    lock.releaseLock();
  }
}

function registerPushDevice_(token) {
  validatePushToken_(token);
  const sheet = ensurePushSheet_();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const rows = sheet.getLastRow() > 1
      ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getDisplayValues().flat()
      : [];
    const existingIndex = rows.indexOf(token);
    const rowNumber = existingIndex === -1 ? sheet.getLastRow() + 1 : existingIndex + 2;
    sheet.getRange(rowNumber, 1, 1, 2).setValues([[token, new Date()]]);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function unregisterPushDevice_(token) {
  validatePushToken_(token);
  const sheet = getPushSheet_();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (sheet.getLastRow() < 2) return { ok: true };
    const tokens = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getDisplayValues().flat();
    tokens.map(function (value, index) {
      return value === token ? index + 2 : 0;
    }).filter(Boolean).reverse().forEach(function (rowNumber) {
      sheet.deleteRow(rowNumber);
    });
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function ensurePushSheet_() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = spreadsheet.getSheetByName(PUSH_SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(PUSH_SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(PUSH_HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange('A:A').setNumberFormat('@');
  }
  if (!sheet.isSheetHidden()) sheet.hideSheet();
  return sheet;
}

function getPushSheet_() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(PUSH_SHEET_NAME);
  if (!sheet) throw new Error('Лист push-устройств не найден. Запустите setupPushNotifications().');
  return sheet;
}

function sendFcmNotification_(token, entry) {
  const accountJson = PropertiesService.getScriptProperties().getProperty(FCM_SERVICE_ACCOUNT_PROPERTY);
  if (!accountJson) throw new Error(`Добавьте ${FCM_SERVICE_ACCOUNT_PROPERTY} в Script Properties.`);
  const account = JSON.parse(accountJson);
  if (!account.project_id || !account.client_email || !account.private_key) {
    throw new Error('FCM_SERVICE_ACCOUNT_JSON должен содержать project_id, client_email и private_key.');
  }
  const accessToken = getFcmAccessToken_(account);
  const payload = {
    message: {
      token: token,
      data: {
        title: 'Пора делать укол!',
        body: 'С последней записанной инъекции прошло 12 часов. Следуйте плану ветеринарного врача.',
        url: 'https://con131.github.io/dog-insulin-tracker/',
        tag: `dog-insulin-${entry.id}`
      },
      webpush: { headers: { TTL: '86400' } }
    }
  };
  const response = UrlFetchApp.fetch(
    `https://fcm.googleapis.com/v1/projects/${encodeURIComponent(account.project_id)}/messages:send`,
    {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: `Bearer ${accessToken}` },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    }
  );
  const status = response.getResponseCode();
  if (status >= 200 && status < 300) return true;
  const body = response.getContentText();
  if (status === 404 || /UNREGISTERED|registration-token-not-registered/.test(body)) return false;
  throw new Error(`FCM вернул HTTP ${status}: ${body}`);
}

function getFcmAccessToken_(account) {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(FCM_ACCESS_TOKEN_CACHE_KEY);
  if (cached) return cached;

  const now = Math.floor(Date.now() / 1000);
  const header = base64Url_(JSON.stringify({ alg: 'RS256', typ: 'JWT' }));
  const claims = base64Url_(JSON.stringify({
    iss: account.client_email,
    scope: 'https://www.googleapis.com/auth/firebase.messaging',
    aud: 'https://oauth2.googleapis.com/token',
    iat: now,
    exp: now + 3600
  }));
  const unsigned = `${header}.${claims}`;
  const key = account.private_key.replace(/-----BEGIN PRIVATE KEY-----|-----END PRIVATE KEY-----|\s/g, '');
  const signature = Utilities.computeRsaSignature(
    Utilities.RsaAlgorithm.SHA_256,
    unsigned,
    Utilities.base64Decode(key)
  );
  const assertion = `${unsigned}.${base64Url_(signature)}`;
  const response = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', {
    method: 'post',
    payload: {
      grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
      assertion: assertion
    },
    muteHttpExceptions: true
  });
  const status = response.getResponseCode();
  const result = JSON.parse(response.getContentText() || '{}');
  if (status < 200 || status >= 300 || !result.access_token) {
    throw new Error(`Не удалось получить токен Google OAuth (HTTP ${status}): ${response.getContentText()}`);
  }
  cache.put(FCM_ACCESS_TOKEN_CACHE_KEY, result.access_token, Math.min(Number(result.expires_in) - 60, 3300));
  return result.access_token;
}

function base64Url_(value) {
  const bytes = typeof value === 'string' ? Utilities.newBlob(value).getBytes() : value;
  return Utilities.base64EncodeWebSafe(bytes).replace(/=+$/, '');
}

function validatePushToken_(token) {
  if (typeof token !== 'string' || token.length < 20 || token.length > 4096 ||
      !/^[A-Za-z0-9:_-]+$/.test(token)) {
    throw new Error('Некорректный FCM token.');
  }
}

function getSheet_() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Лист «Инъекции» не найден. Сначала запустите setupTrackingSheet().');
  return sheet;
}

function ensureAuthorColumn_(sheet) {
  const lastColumn = Math.max(sheet.getLastColumn(), HEADERS.length);
  const headers = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0];
  const authorColumn = headers.indexOf('Кто сделал запись') + 1;
  if (authorColumn) return authorColumn;
  const newColumn = Math.max(sheet.getLastColumn() + 1, HEADERS.length);
  sheet.getRange(1, newColumn).setValue('Кто сделал запись');
  return newColumn;
}

function validateRecord_(record) {
  if (!record || typeof record !== 'object') throw new Error('Ожидался объект записи.');
  if (typeof record.id !== 'string' || !record.id || record.id.length > 100) {
    throw new Error('Укажите корректный ID записи.');
  }
  if (typeof record.injectedAt !== 'string' || isNaN(new Date(record.injectedAt).getTime())) {
    throw new Error('Укажите корректные дату и время.');
  }
  if (typeof record.dose !== 'number' || !isFinite(record.dose) || record.dose <= 0) {
    throw new Error('Доза должна быть положительным числом.');
  }
  if (record.glucose !== '' && record.glucose != null &&
      (typeof record.glucose !== 'number' || !isFinite(record.glucose) || record.glucose < 0)) {
    throw new Error('Уровень сахара должен быть неотрицательным числом.');
  }
  if (typeof record.author !== 'string' || !record.author.trim() || record.author.trim().length > 100) {
    throw new Error('Укажите имя того, кто сделал запись (до 100 символов).');
  }
  ['site', 'meal', 'note'].forEach(function (field) {
    if (record[field] != null && String(record[field]).length > (field === 'note' ? 1000 : 100)) {
      throw new Error(`Поле ${field} слишком длинное.`);
    }
  });
}

function normalizeDate_(value) {
  if (value instanceof Date) return value.toISOString();
  const date = new Date(value);
  return isNaN(date.getTime()) ? '' : date.toISOString();
}

function safeCell_(value) {
  const text = String(value == null ? '' : value);
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function jsonOutput_(value) {
  return ContentService.createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}
