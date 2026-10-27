/**
 * Sports Corner — Store Dashboard (Google Apps Script web app).
 * Data lives in this Google Sheet (tab "DashboardData"). Every change is also written to the "History" tab,
 * so any save or delete can be undone from the app (Data entry → Undo & history).
 * Everyone with the link sees the dashboard live. Only someone with the PIN below can enter or change data.
 */

// 1) CHANGE THIS PIN before you deploy. Use 6 digits or more. Staff never see this file.
const EDIT_PIN = 'CHANGE-ME';

const SHEET_NAME = 'DashboardData';
const HISTORY_NAME = 'History';
const ID_RULE = /^((d|p):\d{4}-\d{2}-\d{2}|(m|h|s):\d{4}-\d{2})$/;

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Sports Corner · Store Dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

// 2) RUN THIS ONCE from the editor (choose "setup" at the top, then Run). It copies your data (October daily, last year, YTD months, staff list) into the sheet.
function setup() {
  const sh = sheet_();
  if (sh.getLastRow() > 1) { Logger.log('Already set up: the sheet has data, nothing was changed.'); return; }
  const ids = Object.keys(SEED).sort();
  const now = new Date().toISOString();
  const rows = ids.map(id => [id, JSON.stringify(SEED[id]), now]);
  sh.getRange(2, 1, rows.length, 3).setValues(rows);
  history_();
  bump_();
  Logger.log('Done: copied ' + ids.length + ' records into the sheet.');
}

/* ---------- called by the page ---------- */

function getAll() {
  const values = sheet_().getDataRange().getValues();
  const docs = {};
  for (let i = 1; i < values.length; i++) {
    const id = String(values[i][0] || '');
    const json = String(values[i][1] || '');
    if (!id || !json) continue;
    try { docs[id] = JSON.parse(json); } catch (e) { /* skip a damaged row */ }
  }
  return JSON.stringify({ docs: docs, rev: rev_() });
}

function getRev() { return rev_(); }

function checkPin(pin) { return String(pin || '') === String(EDIT_PIN); }

/** changes: JSON array of [id, data] pairs; data null deletes that record. label: short description for the history list. */
function saveDocs(pin, changesJson, label) {
  guard_(pin);
  const changes = JSON.parse(changesJson);
  if (!Array.isArray(changes) || changes.length > 800) throw new Error('Bad request');
  changes.forEach(c => { if (!ID_RULE.test(String(c[0]))) throw new Error('Bad record id ' + c[0]); });
  return withLock_(() => JSON.stringify(apply_(changes, String(label || 'Change').slice(0, 120))));
}

/** The most recent saves, newest first, for the "Undo & history" list. */
function getHistory(pin) {
  guard_(pin);
  const h = history_();
  const last = h.getLastRow();
  if (last < 2) return JSON.stringify([]);
  const start = Math.max(2, last - 1500);
  const rows = h.getRange(start, 1, last - start + 1, 4).getValues();
  const batches = [];
  const byId = {};
  rows.forEach(r => {
    const b = String(r[0]);
    if (!byId[b]) { byId[b] = { batch: b, time: String(r[1]), label: String(r[2]), count: 0, ids: [] }; batches.push(byId[b]); }
    byId[b].count++;
    if (byId[b].ids.length < 4) byId[b].ids.push(String(r[3]));
  });
  return JSON.stringify(batches.slice(-40).reverse());
}

/** Puts every record touched by one save back the way it was before that save (deleted records come back). */
function undoBatch(pin, batch) {
  guard_(pin);
  return withLock_(() => {
    const h = history_();
    const last = h.getLastRow();
    if (last < 2) throw new Error('Nothing to undo');
    const rows = h.getRange(2, 1, last - 1, 6).getValues().filter(r => String(r[0]) === String(batch));
    if (!rows.length) throw new Error('That change is no longer in the history');
    const changes = rows.slice().reverse().map(r => [String(r[3]), r[4] === '' ? null : JSON.parse(String(r[4]))]);
    return JSON.stringify(apply_(changes, 'Undo: ' + String(rows[0][2])));
  });
}

/** Copies the whole data tab to a new tab named "Backup <date time>". */
function backupNow(pin) {
  guard_(pin);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = 'Backup ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH-mm');
  sheet_().copyTo(ss).setName(name);
  return name;
}

/* ---------- helpers ---------- */

function apply_(changes, label) {
  const sh = sheet_();
  const last = sh.getLastRow();
  const vals = last > 1 ? sh.getRange(2, 1, last - 1, 2).getValues() : [];
  const ids = vals.map(r => String(r[0]));
  const now = new Date().toISOString();
  const batch = Utilities.getUuid().slice(0, 8);
  const hist = [];
  const toDelete = [];
  const seen = {};
  changes.forEach(c => {
    const id = String(c[0]);
    if (seen[id]) return; seen[id] = true;
    const i = ids.indexOf(id);
    const before = i >= 0 ? String(vals[i][1]) : '';
    const after = c[1] === null ? '' : JSON.stringify(c[1]);
    if (before === after) return;
    hist.push([batch, now, label, id, before, after]);
    if (c[1] === null) { if (i >= 0) toDelete.push(i + 2); return; }
    const row = [id, after, now];
    if (i >= 0) sh.getRange(i + 2, 1, 1, 3).setValues([row]);
    else { sh.getRange(sh.getLastRow() + 1, 1, 1, 3).setValues([row]); ids.push(id); vals.push([id, after]); }
  });
  toDelete.sort((a, b) => b - a).forEach(r => sh.deleteRow(r));
  if (hist.length) {
    const h = history_();
    h.getRange(h.getLastRow() + 1, 1, hist.length, 6).setValues(hist);
    if (h.getLastRow() > 20000) h.deleteRows(2, 5000);
  }
  return { rev: bump_(), batch: hist.length ? batch : null, changed: hist.length };
}

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

function history_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let h = ss.getSheetByName(HISTORY_NAME);
  if (!h) {
    h = ss.insertSheet(HISTORY_NAME);
    h.getRange(1, 1, 1, 6).setValues([['change', 'time', 'what', 'record', 'before', 'after']]);
    h.getRange('A:D').setNumberFormat('@');
  }
  return h;
}

function rev_() { return Number(PropertiesService.getScriptProperties().getProperty('rev') || '0'); }

function bump_() {
  const r = rev_() + 1;
  PropertiesService.getScriptProperties().setProperty('rev', String(r));
  return r;
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

/** API for the installable app (Dul-KPI Dashboard on GitHub Pages). */
function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents);
    const fns = { getAll: getAll, getRev: getRev, checkPin: checkPin, saveDocs: saveDocs, getHistory: getHistory, undoBatch: undoBatch, backupNow: backupNow };
    const f = fns[req.fn];
    if (!f) throw new Error('Unknown action');
    const r = f.apply(null, req.args || []);
    out = { ok: r === undefined ? null : r };
  } catch (err) { out = { err: err.message }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
