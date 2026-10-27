/**
 * F03 Shift Board — Google Apps Script web app.
 * Data lives in this Google Sheet (tab "ScheduleData"). Staff open the web app link and see the schedule live.
 * Only someone with the PIN below can edit.
 */

// 1) CHANGE THIS PIN before you deploy. Use 6 digits or more. Staff never see this file.
const EDIT_PIN = 'CHANGE-ME';

const SHEET_NAME = 'ScheduleData';
const COLORS_ID = 'settings:colors';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('F03 Shift Board')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

// 2) RUN THIS ONCE from the editor (select "setup" at the top, then Run). It copies your schedule into the sheet.
function setup() {
  const sh = sheet_();
  if (sh.getLastRow() > 1) { Logger.log('Already set up: the sheet has data, nothing was changed.'); return; }
  const ids = Object.keys(SEED.weeks).sort();
  const rows = ids.map(id => ['w:' + id, JSON.stringify(SEED.weeks[id]), new Date().toISOString()]);
  rows.push([COLORS_ID, JSON.stringify(SEED.colors || {}), new Date().toISOString()]);
  sh.getRange(2, 1, rows.length, 3).setValues(rows);
  bump_();
  Logger.log('Done: copied ' + ids.length + ' weeks into the sheet.');
}

/* ---------- called by the page ---------- */

function getAll() {
  const values = sheet_().getDataRange().getValues();
  const weeks = {};
  let colors = {};
  for (let i = 1; i < values.length; i++) {
    const id = String(values[i][0] || '');
    const json = String(values[i][1] || '');
    if (!id || !json) continue;
    try {
      if (id === COLORS_ID) colors = JSON.parse(json);
      else if (id.indexOf('w:') === 0) weeks[id.slice(2)] = JSON.parse(json);
    } catch (e) { /* skip a damaged row */ }
  }
  return JSON.stringify({ weeks: weeks, colors: colors, rev: rev_() });
}

function getRev() { return rev_(); }

function checkPin(pin) { return String(pin || '') === String(EDIT_PIN); }

function saveWeek(pin, id, json) {
  guard_(pin);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(id))) throw new Error('Bad week id');
  JSON.parse(json); // reject anything that isn't valid data
  return withLock_(() => { put_('w:' + id, json); return bump_(); });
}

function deleteWeek(pin, id) {
  guard_(pin);
  return withLock_(() => {
    const sh = sheet_();
    const row = findRow_(sh, 'w:' + id);
    if (row > 0) sh.deleteRow(row);
    return bump_();
  });
}

function saveColors(pin, json) {
  guard_(pin);
  JSON.parse(json);
  return withLock_(() => { put_(COLORS_ID, json); return bump_(); });
}

/* ---------- helpers ---------- */

function guard_(pin) { if (!checkPin(pin)) throw new Error('PIN incorrect'); }

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, 3).setValues([['id', 'data', 'updated']]);
    sh.getRange('A:A').setNumberFormat('@');
  }
  return sh;
}

function findRow_(sh, id) {
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === id) return i + 2;
  return -1;
}

function put_(id, json) {
  const sh = sheet_();
  const row = [id, json, new Date().toISOString()];
  const r = findRow_(sh, id);
  if (r > 0) sh.getRange(r, 1, 1, 3).setValues([row]);
  else sh.getRange(sh.getLastRow() + 1, 1, 1, 3).setValues([row]);
}

function rev_() { return Number(PropertiesService.getScriptProperties().getProperty('rev') || '0'); }

function bump_() {
  const r = rev_() + 1;
  PropertiesService.getScriptProperties().setProperty('rev', String(r));
  return r;
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try { return fn(); } finally { lock.releaseLock(); }
}

/** API for the installable app (Dul-Shift Board on GitHub Pages). */
function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents);
    const fns = { getAll: getAll, getRev: getRev, checkPin: checkPin, saveWeek: saveWeek, deleteWeek: deleteWeek, saveColors: saveColors };
    const f = fns[req.fn];
    if (!f) throw new Error('Unknown action');
    const r = f.apply(null, req.args || []);
    out = { ok: r === undefined ? null : r };
  } catch (err) { out = { err: err.message }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
