/**
 * Backend for the dog insulin tracker.
 *
 * Attach this script to the destination spreadsheet (Extensions > Apps Script),
 * run setupTrackingSheet() once, then deploy it as a web app.
 *
 * ContentService does not expose arbitrary Access-Control-Allow-Origin headers.
 * GET therefore supports JSONP for browsers on GitHub Pages. The browser sends
 * POST as text/plain (a CORS "simple request"), and verifies writes by reading
 * the record back. IDs make retrying an offline queue safe and idempotent.
 */
const SHEET_NAME = 'Инъекции';
const HEADERS = ['ID', 'Дата и время (ISO)', 'Доза (ЕД)', 'Сахар (ммоль/л)', 'Зона', 'Приём пищи', 'Примечание'];

function setupTrackingSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    throw new Error('Скрипт должен быть привязан к Google Таблице.');
  }
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  }
  return `Готово: лист «${SHEET_NAME}» подготовлен.`;
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('Пустое тело запроса.');
    }
    const record = JSON.parse(e.postData.contents);
    validateRecord_(record);

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const sheet = getSheet_();
      const existingIds = sheet.getLastRow() > 1
        ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getDisplayValues().flat()
        : [];
      if (existingIds.indexOf(record.id) === -1) {
        sheet.appendRow([
          safeCell_(record.id),
          record.injectedAt,
          record.dose,
          record.glucose === '' || record.glucose == null ? '' : record.glucose,
          safeCell_(record.site),
          safeCell_(record.meal),
          safeCell_(record.note || '')
        ]);
      }
    } finally {
      lock.releaseLock();
    }
    return jsonOutput_({ ok: true, id: record.id });
  } catch (error) {
    return jsonOutput_({ ok: false, error: String(error.message || error) });
  }
}

function doGet(e) {
  try {
    const sheet = getSheet_();
    const rows = sheet.getLastRow() > 1
      ? sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS.length).getValues()
      : [];
    const records = rows.map(function (row) {
      return {
        id: String(row[0]),
        injectedAt: normalizeDate_(row[1]),
        dose: Number(row[2]),
        glucose: row[3] === '' ? '' : Number(row[3]),
        site: String(row[4]),
        meal: String(row[5]),
        note: String(row[6] || '')
      };
    }).filter(function (record) {
      return record.id && record.injectedAt;
    }).sort(function (a, b) {
      return new Date(b.injectedAt).getTime() - new Date(a.injectedAt).getTime();
    });

    const callback = e && e.parameter && e.parameter.callback;
    if (callback) {
      if (!/^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callback)) {
        throw new Error('Недопустимое имя JSONP callback.');
      }
      return ContentService.createTextOutput(
        callback + '(' + JSON.stringify({ entries: records }) + ');'
      ).setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return jsonOutput_({ entries: records });
  } catch (error) {
    return jsonOutput_({ error: String(error.message || error), entries: [] });
  }
}

function getSheet_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) throw new Error('Не найдена привязанная Google Таблица.');
  const sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Лист не найден. Сначала запустите setupTrackingSheet().');
  return sheet;
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

// Prefix user-provided text beginning with spreadsheet formula characters.
function safeCell_(value) {
  const text = String(value == null ? '' : value);
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function jsonOutput_(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
