/**
 * Sports Corner – Price Check (Was / Now) for PDT and phones.
 * Staff open the web app link, scan a barcode or type a search name.
 * The manager uploads the price list Excel (xlsb / xlsx) from the ⚙ screen with the PIN below.
 * A new upload replaces the old list for every device automatically.
 */
const ADMIN_PIN = 'CHANGE-ME';       // <-- your own PIN (set it in the Apps Script editor, not here)
const FOLDER_NAME = 'Price Check files';
const CHUNK = 900000;                // characters per download/upload piece

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Price Check')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function folder_() {
  const it = DriveApp.getFoldersByName(FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME);
}
function meta_() {
  const s = PropertiesService.getScriptProperties().getProperty('meta');
  return s ? JSON.parse(s) : null;
}
function pin_(pin) {
  if (String(pin) !== String(ADMIN_PIN)) throw new Error('Wrong PIN');
}

/** Small call every device makes to see if the price list changed. */
function getMeta() { return meta_(); }

function checkPin(pin) { pin_(pin); return true; }

/** Download one piece of the current price list. */
function getChunk(ver, i) {
  const m = meta_();
  if (!m || m.ver !== ver) throw new Error('The price list was changed. Loading the new one…');
  const s = DriveApp.getFileById(m.fileId).getBlob().getDataAsString('UTF-8');
  return s.slice(i * CHUNK, (i + 1) * CHUNK);
}

/** Upload one piece of a new price list (saved as a temporary file). */
function uploadChunk(pin, id, i, data) {
  pin_(pin);
  if (!/^[a-z0-9]{6,20}$/.test(id)) throw new Error('Bad upload id');
  folder_().createFile('upload_' + id + '_' + i + '.part', data, MimeType.PLAIN_TEXT);
  return true;
}

/** Join the pieces, make it the live price list and remove the old list. */
function finishUpload(pin, id, n, info) {
  pin_(pin);
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const f = folder_(), parts = [];
    for (let i = 0; i < n; i++) {
      const it = f.getFilesByName('upload_' + id + '_' + i + '.part');
      if (!it.hasNext()) throw new Error('Upload piece ' + (i + 1) + ' is missing. Please upload again.');
      const file = it.next(); parts.push(file.getBlob().getDataAsString('UTF-8')); file.setTrashed(true);
    }
    const text = parts.join('');
    JSON.parse(text);                                   // make sure it is complete
    const ver = id;
    const file = f.createFile('pricelist_' + ver + '.json', text, MimeType.PLAIN_TEXT);
    const old = meta_();
    if (old && old.fileId) { try { DriveApp.getFileById(old.fileId).setTrashed(true); } catch (e) {} }
    const m = Object.assign({}, info || {}, {
      ver: ver, fileId: file.getId(), size: text.length, chunks: Math.ceil(text.length / CHUNK),
      uploaded: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')
    });
    PropertiesService.getScriptProperties().setProperty('meta', JSON.stringify(m));
    return m;
  } finally { lock.releaseLock(); }
}

/** Delete the price list (devices show "no price list" until a new one is uploaded). */
function deleteList(pin) {
  pin_(pin);
  const old = meta_();
  if (old && old.fileId) { try { DriveApp.getFileById(old.fileId).setTrashed(true); } catch (e) {} }
  PropertiesService.getScriptProperties().deleteProperty('meta');
  return true;
}

/** API for the installable app (GitHub Pages). Same functions as above, called with JSON. */
function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents);
    const fns = { getMeta: getMeta, checkPin: checkPin, getChunk: getChunk, uploadChunk: uploadChunk, finishUpload: finishUpload, deleteList: deleteList };
    const f = fns[req.fn];
    if (!f) throw new Error('Unknown action');
    const r = f.apply(null, req.args || []);
    out = { ok: r === undefined ? null : r };
  } catch (err) { out = { err: err.message }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
