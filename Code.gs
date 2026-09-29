// Paste this whole file into Apps Script (Code.gs), replacing everything.
const PASS = 'manoj@123';   // your passcode
const SHEET_ID = '';        // optional: only fill if this script was NOT opened from the sheet (Extensions > Apps Script)

function doGet() { return out({ ok: true, status: 'Lab backend is running' }); }

function getSheet_() {
  const ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('records');
  if (!sh) {
    sh = ss.insertSheet('records');
    sh.appendRow(['LabNo', 'Date', 'Name', 'AgeSex', 'Data', 'Updated']);
  }
  return sh;
}

function doPost(e) {
  try {
    const q = JSON.parse(e.postData.contents);
    if (q.pass !== PASS) return out({ ok: false, error: 'auth' });
    if (q.action === 'ping') return out({ ok: true });

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const sh = getSheet_();
      const rows = sh.getDataRange().getValues();
      const find = labNo => rows.findIndex((r, i) => i > 0 && String(r[0]) === String(labNo));

      if (q.action === 'save') {
        const row = [q.labNo, q.date, q.name, q.ageSex, JSON.stringify(q.data), new Date()];
        const i = find(q.labNo);
        if (i > -1) sh.getRange(i + 1, 1, 1, 6).setValues([row]); else sh.appendRow(row);
        return out({ ok: true });
      }
      if (q.action === 'list') {
        const list = rows.slice(1).reverse().slice(0, 300)
          .map(r => ({ labNo: r[0], date: r[1], name: r[2], ageSex: r[3] }));
        return out({ ok: true, list });
      }
      if (q.action === 'get') {
        const i = find(q.labNo);
        return out(i > -1 ? { ok: true, data: JSON.parse(rows[i][4]) } : { ok: false, error: 'not found' });
      }
      if (q.action === 'delete') {
        const i = find(q.labNo);
        if (i > -1) sh.deleteRow(i + 1);
        return out({ ok: true });
      }
      return out({ ok: false, error: 'unknown action' });
    } finally { lock.releaseLock(); }
  } catch (err) {
    return out({ ok: false, error: String(err) });
  }
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
