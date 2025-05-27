// GAS Web API for reading and writing spreadsheet data
// Spreadsheet ID is shared via script property SPREADSHEET_ID
// This file provides doGet and doPost handlers for external requests.

const SHEET_ID_API = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
const TARGET_SHEET = 'イベント一覧';
const DEFAULT_LIMIT = 10;  // デフォルト取得件数
const MAX_LIMIT = 20;       // 上限取得件数

function getSheetValues(name) {
  const ss = SpreadsheetApp.openById(SHEET_ID_API);
  const sh = ss.getSheetByName(name);
  if (!sh) return [];
  const lastRow = sh.getLastRow();
  if (lastRow === 0) return [];
  return sh.getRange(1, 1, lastRow, 5).getValues();  // 必要な5列のみ取得
}

function appendRows(name, rows) {
  const ss = SpreadsheetApp.openById(SHEET_ID_API);
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (!Array.isArray(rows) || rows.length === 0) return 0;
  const start = sh.getLastRow() + 1;
  sh.getRange(start, 1, rows.length, rows[0].length).setValues(rows);
  return rows.length;
}

function parseDateParam(str) {
  if (!str) return null;
  const d = new Date(str);
  if (isNaN(d)) return null;
  d.setHours(0, 0, 0, 0);
  return d;
}

function searchEvents(params) {
  const start = parseDateParam(params.start);
  const end = parseDateParam(params.end);
  const keyword = params.keyword ? params.keyword.toLowerCase() : '';
  const category = params.category || '';
  let limit = parseInt(params.limit, 10);
  if (isNaN(limit) || limit <= 0) limit = DEFAULT_LIMIT;
  if (limit > MAX_LIMIT) limit = MAX_LIMIT;

  const rows = getSheetValues(TARGET_SHEET);
  const filtered = rows.filter(r => {
    let ok = true;

    if (start || end) {
      const date = parseDateParam(r[2]);
      if (!date) return false;
      if (start && date < start) ok = false;
      if (end && date > end) ok = false;
    }

    if (keyword) {
      const t = (r[0] || '').toString().toLowerCase();
      const s = (r[1] || '').toString().toLowerCase();
      if (t.indexOf(keyword) === -1 && s.indexOf(keyword) === -1) ok = false;
    }

    if (category && r[4] !== category) {
      ok = false;
    }

    return ok;
  });

  const limited = filtered.slice(0, limit);
  return { sheet: TARGET_SHEET, rows: limited };
}

function doGet(e) {
  let sheetName = TARGET_SHEET;
  if (e && e.parameter) {
    if (e.parameter.mode === 'search') {
      const result = searchEvents(e.parameter);
      return ContentService
        .createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }
    if (e.parameter.category) {
      sheetName = e.parameter.category;
    } else if (e.parameter.sheet) {
      sheetName = e.parameter.sheet;
    }
  }
  const limitParam = e && e.parameter ? parseInt(e.parameter.limit, 10) : NaN;
  let limit = isNaN(limitParam) || limitParam <= 0 ? DEFAULT_LIMIT : limitParam;
  if (limit > MAX_LIMIT) limit = MAX_LIMIT;

  const data = getSheetValues(sheetName).slice(0, limit);
  const output = { sheet: sheetName, rows: data };
  return ContentService
    .createTextOutput(JSON.stringify(output))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const body = e && e.postData && e.postData.contents;
  if (!body) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: 'no data' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  let obj;
  try {
    obj = JSON.parse(body);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: 'invalid json' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  const sheetName = obj.sheet || TARGET_SHEET;
  const rows = obj.rows || (obj.row ? [obj.row] : []);
  const count = appendRows(sheetName, rows);
  return ContentService
    .createTextOutput(JSON.stringify({ sheet: sheetName, inserted: count }))
    .setMimeType(ContentService.MimeType.JSON);
}
