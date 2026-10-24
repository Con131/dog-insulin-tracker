/**
 * Private Google Apps Script backend for the insulin diary.
 *
 * In Apps Script, add an HTML file named "index" and paste index.html into it.
 * Deploy the web app as the owner with access set to "Only myself".
 *
 * There is deliberately no public JSON/JSONP endpoint. doGet serves the UI;
 * reads and writes are google.script.run RPC calls protected by the deployment
 * access rule and execute as the signed-in owner.
 */
const SPREADSHEET_ID = '10HzH-zEU-aO1zhxWQqrTx93quhEtbF1nBT7pZuZAT04';
const SHEET_NAME = 'Инъекции';
const HEADERS = ['ID', 'Дата и время (ISO)', 'Доза (ЕД)', 'Сахар (ммоль/л)', 'Зона', 'Приём пищи', 'Примечание'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Инсулин-пёс — личный дневник')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

function setupTrackingSheet() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  }
  return `Готово: лист «${SHEET_NAME}» подготовлен.`;
}

/**
 * Called only through authenticated google.script.run in the private app.
 * The deployment must remain restricted to its owner.
 */
function getEntries() {
  const sheet = getSheet_();
  const rows = sheet.getLastRow() > 1
    ? sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS.length).getValues()
    : [];
  return rows.map(function (row) {
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
}

/**
 * Add one record. IDs make retries safe when a client reconnects after a timeout.
 */
function addEntry(record) {
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
    return { ok: true, id: record.id };
  } finally {
    lock.releaseLock();
  }
}

function getSheet_() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Лист «Инъекции» не найден. Сначала запустите setupTrackingSheet().');
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

function safeCell_(value) {
  const text = String(value == null ? '' : value);
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}
