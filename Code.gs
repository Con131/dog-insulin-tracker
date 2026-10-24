/**
 * Private Google Apps Script backend for the insulin diary.
 *
 * In Apps Script, add an HTML file named "index" and paste index.html into it.
 * Deploy the web app as the owner. Anyone with its URL can access it.
 *
 * There is deliberately no public JSON/JSONP endpoint. doGet serves the UI;
 * reads and writes are google.script.run RPC calls. The deployment executes as
 * the owner; the URL is a bearer credential and must be shared with care.
 */
const SPREADSHEET_ID = '10HzH-zEU-aO1zhxWQqrTx93quhEtbF1nBT7pZuZAT04';
const SHEET_NAME = 'Инъекции';
const HEADERS = ['ID', 'Дата и время (ISO)', 'Доза (ЕД)', 'Сахар (ммоль/л)', 'Зона', 'Приём пищи', 'Примечание', 'Кто сделал запись'];

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
  } else {
    ensureAuthorColumn_(sheet);
  }
  return `Готово: лист «${SHEET_NAME}» подготовлен.`;
}

/**
 * Called only through authenticated google.script.run in the private app.
 * The deployment must remain restricted to its owner.
 */
function getEntries() {
  const sheet = getSheet_();
  const authorColumn = ensureAuthorColumn_(sheet);
  const columnCount = Math.max(HEADERS.length, authorColumn);
  const rows = sheet.getLastRow() > 1
    ? sheet.getRange(2, 1, sheet.getLastRow() - 1, columnCount).getValues()
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

/**
 * Add one record. IDs make retries safe when a client reconnects after a timeout.
 */
function saveEntry(record) {
  validateRecord_(record);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheet_();
    const authorColumn = ensureAuthorColumn_(sheet);
    const existingIds = sheet.getLastRow() > 1
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
    const existingIndex = existingIds.indexOf(record.id);
    const rowNumber = existingIndex === -1 ? sheet.getLastRow() + 1 : existingIndex + 2;
    sheet.getRange(rowNumber, 1, 1, values.length).setValues([values]);
    sheet.getRange(rowNumber, authorColumn).setValue(safeCell_(record.author));
    return { ok: true, id: record.id, updated: existingIndex !== -1 };
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
