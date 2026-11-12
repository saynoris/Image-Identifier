// Student Identification collector. Paste into Extensions > Apps Script in a Google Sheet.
// 1) Change ADMIN_KEY to a long secret only you know.
// 2) Deploy > New deployment > Web app. Execute as: Me. Who has access: Anyone.
// 3) Copy the Web app URL (ends in /exec) into index.html.

const ADMIN_KEY = 'CHANGE-THIS-TO-A-LONG-SECRET';
const SHEET_NAME = 'Entries';

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['Time', 'Photo key', 'Photo file', 'Student ID', 'Full name']);
    sh.getRange('B:E').setNumberFormat('@'); // keep IDs like 00123 and stop text turning into formulas
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action === 'submit') {
    const sid = String(p.sid || '').trim().slice(0, 60);
    const name = String(p.name || '').trim().slice(0, 120);
    const key = String(p.key || '').slice(0, 40);
    const file = String(p.file || '').slice(0, 120);
    if (!sid || !name || !key) return json_({ ok: false, error: 'missing' });
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      getSheet_().appendRow([new Date(), key, file, sid, name]);
    } finally {
      lock.releaseLock();
    }
    return json_({ ok: true });
  }
  if (p.action === 'list') {
    if (p.admin !== ADMIN_KEY) return json_({ ok: false, error: 'forbidden' });
    const values = getSheet_().getDataRange().getValues();
    const rows = values.slice(1).map(function (r) {
      return { t: r[0] instanceof Date ? r[0].getTime() : Date.parse(r[0]), key: r[1], file: r[2], sid: r[3], name: r[4] };
    });
    return json_({ ok: true, rows: rows });
  }
  return json_({ ok: true });
}
