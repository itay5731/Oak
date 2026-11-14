/**
 * Oak: missing invoices tracker (Google Apps Script web app)
 * Data lives in a Google Sheet inside the Drive folder you pick on first run.
 * Invoices are saved to that folder, one subfolder per month, renamed automatically.
 */

/** Optional: hard-code who can open Oak, e.g. ['acme.com', 'partner.org', 'accountant@gmail.com'].
 *  Leave empty to set it in the app (first-run screen and Settings). */
const ALLOWED = [];

/** Optional: a public https link to a PNG for the browser tab icon. */
const ICON_URL = 'https://raw.githubusercontent.com/itay5731/Oak/main/assets/icon.png';

const PROPS = PropertiesService.getScriptProperties();
const COLS = ['id', 'key', 'month', 'date', 'merchant', 'amount', 'currency', 'charged',
  'chargedCurrency', 'details', 'status', 'owner', 'notes', 'files', 'extra', 'updatedBy', 'updatedAt', 'createdAt', 'invoiceNo'];
const JSON_COLS = { notes: [], files: [], extra: {} };


function doGet() {
  const out = HtmlService.createHtmlOutputFromFile('Index');
  if (ICON_URL) { try { out.setFaviconUrl(ICON_URL); } catch (e) { /* tab icon is optional */ } }
  return out
    .setTitle('Oak')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ---------- state ---------- */

function getState() {
  if (!allowed_()) return { denied: true, user: me_(), url: ScriptApp.getService().getUrl() };
  const c = config_();
  const state = { user: me_(), configured: !!(c.folderId && c.sheetId), domain: me_().split('@')[1] || '' };
  if (!state.configured) return state;
  let folder;
  try { folder = DriveApp.getFolderById(c.folderId); folder.getName(); SpreadsheetApp.openById(c.sheetId); }
  catch (e) { throw new Error('You don\'t have access to Oak\'s Drive folder yet. Ask the person who set up Oak to share it with you as an editor.'); }
  state.folder = { name: folder.getName(), url: folder.getUrl() };
  state.team = team_(folder);
  state.sheetUrl = SpreadsheetApp.openById(c.sheetId).getUrl();
  state.allowed = allowList_();
  state.allowedLocked = ALLOWED.length > 0;
  state.summary = { email: PROPS.getProperty('summaryEmail') || '', on: PROPS.getProperty('summaryOn') === '1' };
  state.syncDue = Date.now() - Number(PROPS.getProperty('lastSync') || 0) > 15 * 60 * 1000;
  const sh = sheet_();
  if (sh.getRange(1, COLS.length).getValue() !== COLS[COLS.length - 1]) sh.getRange(1, 1, 1, COLS.length).setValues([COLS]).setFontWeight('bold');
  state.rows = readRows_(sh).map(clean_);
  return state;
}

function setup(link, allow) {
  auth_();
  return withLock_(function () {
    const id = parseFolderId_(link);
    if (!id) throw new Error('That doesn\'t look like a Drive folder link. Copy it from the address bar of the folder.');
    let folder;
    try { folder = DriveApp.getFolderById(id); folder.getName(); }
    catch (e) { throw new Error('Can\'t open that folder. Check the link, and make sure the app owner can edit it.'); }

    let sheetId = config_().sheetId;
    if (sheetId) {
      DriveApp.getFileById(sheetId).moveTo(folder); // keep the data sheet with the invoices
    } else {
      const ss = SpreadsheetApp.create('Invoice tracker data');
      const sh = ss.getSheets()[0];
      sh.setName('Invoices');
      sh.getRange(1, 1, sh.getMaxRows(), COLS.length).setNumberFormat('@');
      sh.getRange(1, 1, 1, COLS.length).setValues([COLS]).setFontWeight('bold');
      sh.setFrozenRows(1);
      const log = ss.insertSheet('Activity');
      log.appendRow(['time', 'user', 'action', 'row', 'detail']);
      log.setFrozenRows(1);
      DriveApp.getFileById(ss.getId()).moveTo(folder);
      sheetId = ss.getId();
    }
    PROPS.setProperties({ folderId: id, sheetId: sheetId });
    if (allow != null || !allowList_().length) saveAccessRaw_(allow);
    log_('setup', '', folder.getName());
    return getState();
  });
}

/* ---------- import ---------- */

function importRows(incoming) {
  auth_();
  return withLock_(function () {
    const sh = sheet_();
    const all = readRows_(sh);
    const byKey = {};
    all.forEach(function (r) { byKey[r.key] = r; });
    const now = new Date().toISOString();
    const user = me_();
    let added = 0, updated = 0, kept = 0;
    const fresh = [], changed = [];

    incoming.forEach(function (inc) {
      const ex = byKey[inc.key];
      if (!ex) {
        const r = {
          id: Utilities.getUuid().replace(/-/g, '').slice(0, 10), key: inc.key, month: inc.month,
          date: inc.date || '', merchant: inc.merchant, amount: inc.amount, currency: inc.currency || '',
          charged: inc.charged, chargedCurrency: inc.chargedCurrency || '', details: inc.details || '',
          status: inc.status === 'Done' ? 'Done' : 'Missing', owner: inc.owner || '',
          notes: inc.note ? [{ by: 'Excel', at: now, text: inc.note }] : [],
          files: [], extra: inc.extra || {}, updatedBy: user, updatedAt: now, createdAt: now
        };
        fresh.push(r); byKey[inc.key] = r; added++;
      } else {
        let ch = false;
        if (inc.status === 'Done' && ex.status !== 'Done') { ex.status = 'Done'; ch = true; }
        if (inc.note && !ex.notes.some(function (n) { return n.text === inc.note; })) {
          ex.notes.push({ by: 'Excel', at: now, text: inc.note }); ch = true;
        }
        if (inc.owner && !ex.owner) { ex.owner = inc.owner; ch = true; }
        if (inc.extra) {
          Object.keys(inc.extra).forEach(function (k) {
            if (ex.extra[k] !== inc.extra[k]) { ex.extra[k] = inc.extra[k]; ch = true; }
          });
        }
        if (ch) { ex.updatedBy = user; ex.updatedAt = now; changed.push(ex); updated++; } else kept++;
      }
    });

    if (fresh.length) {
      const rng = sh.getRange(sh.getLastRow() + 1, 1, fresh.length, COLS.length);
      rng.setNumberFormat('@');
      rng.setValues(fresh.map(toArr_));
    }
    changed.forEach(function (r) { writeRow_(sh, r); });
    log_('import', '', added + ' added, ' + updated + ' updated');
    return { added: added, updated: updated, kept: kept, rows: readRows_(sh).map(clean_) };
  });
}

/** A charge typed in by hand. It gets the same key an Excel row would, so a later import merges instead of duplicating. */
function addCharge(c) {
  auth_();
  const merchant = String(c.merchant || '').replace(/\s+/g, ' ').trim();
  const amount = Number(c.amount);
  if (!merchant) throw new Error('Add the vendor.');
  if (!isFinite(amount) || amount <= 0) throw new Error('Add the amount.');
  const date = /^\d{4}-\d{2}-\d{2}$/.test(c.date || '') ? c.date : '';
  if (!date) throw new Error('Add the date of the charge.');
  const currency = String(c.currency || '').trim();
  const base = [date.slice(0, 7), date, merchant.toLowerCase(), amount.toFixed(2), currency].join('|');
  const n = readRows_().filter(function (r) { return String(r.key).indexOf(base + '#') === 0; }).length + 1;
  const key = base + '#' + n;
  const res = importRows([{
    key: key, month: date.slice(0, 7), date: date, merchant: merchant, amount: amount, currency: currency,
    charged: amount, chargedCurrency: currency, details: String(c.details || '').trim(),
    status: 'Missing', owner: String(c.owner || '').trim().toLowerCase(), note: String(c.note || '').trim()
  }]);
  log_('add charge', '', merchant);
  return res.rows.filter(function (r) { return r.key === key; })[0];
}

/* ---------- invoices in Drive ---------- */

function uploadInvoice(rowId, file) {
  auth_();
  const r0 = findRow_(sheet_(), rowId);
  const blob = Utilities.newBlob(Utilities.base64Decode(file.data), file.type || 'application/octet-stream', file.name);
  blob.setName(invoiceName_(r0, file.name, r0.files.length));
  const f = monthFolder_(r0.month).createFile(blob);
  f.setDescription('Uploaded by ' + me_() + '. Original name: ' + file.name);
  const invoiceNo = findInvoiceNo_(ocrText_(f.getId()));
  return withLock_(function () {
    const sh = sheet_();
    const r = findRow_(sh, rowId);
    r.files.push({ id: f.getId(), name: f.getName(), original: file.name, url: f.getUrl(), by: me_(), at: new Date().toISOString(), invoiceNo: invoiceNo });
    r.status = 'Done';
    touch_(r); writeRow_(sh, r);
    log_('upload', r.id, f.getName());
    return clean_(r);
  });
}

function removeFile(rowId, fileId) {
  auth_();
  return withLock_(function () {
    const sh = sheet_();
    const r = findRow_(sh, rowId);
    try { DriveApp.getFileById(fileId).setTrashed(true); } catch (e) { /* already gone */ }
    r.files = r.files.filter(function (f) { return f.id !== fileId; });
    if (!r.files.length) r.status = 'Missing';
    touch_(r); writeRow_(sh, r);
    log_('delete file', r.id, fileId);
    return clean_(r);
  });
}

function deleteCharge(rowId) {
  auth_();
  return withLock_(function () {
    const sh = sheet_();
    const r = findRow_(sh, rowId);
    r.files.forEach(function (f) { try { DriveApp.getFileById(f.id).setTrashed(true); } catch (e) {} });
    sh.deleteRow(r._row);
    log_('delete charge', r.id, r.merchant + ' ' + r.amount);
    return { ok: true };
  });
}

/** Re-checks every attached file: renames to the house format, moves to the right month folder,
 *  and flags files removed from Drive. Finance removes handled invoices, so the charge stays Done. */
function syncDrive() {
  auth_();
  return withLock_(function () {
    const sh = sheet_();
    const rows = readRows_(sh);
    let renamed = 0, moved = 0, missing = 0;
    PROPS.setProperty('lastSync', String(Date.now()));
    const folders = {};
    rows.forEach(function (r) {
      if (!r.files.length) return;
      let ch = false;
      const target = folders[r.month] || (folders[r.month] = monthFolder_(r.month));
      r.files.forEach(function (f, i) {
        let file;
        try { file = DriveApp.getFileById(f.id); if (file.isTrashed()) throw new Error('trashed'); }
        catch (e) { if (!f.gone) { f.gone = true; missing++; ch = true; } return; }
        if (f.gone) { delete f.gone; ch = true; }
        const want = invoiceName_(r, f.original || file.getName(), i);
        if (file.getName() !== want) { file.setName(want); f.name = want; renamed++; ch = true; }
        const parents = file.getParents();
        if (!parents.hasNext() || parents.next().getId() !== target.getId()) { file.moveTo(target); moved++; }
        f.url = file.getUrl();
      });
      if (ch) writeRow_(sh, r);
    });
    PROPS.setProperty('lastSync', String(Date.now()));
    if (renamed || moved || missing) log_('sync drive', '', renamed + ' renamed, ' + moved + ' moved, ' + missing + ' removed from Drive');
    return { renamed: renamed, moved: moved, missing: missing, rows: readRows_(sh).map(clean_) };
  });
}

/** Files in the Drive folder (and its subfolders) that aren't attached to any charge yet. */
function listLooseFiles() {
  auth_();
  const cfg = config_();
  const attached = {};
  readRows_().forEach(function (r) { r.files.forEach(function (f) { attached[f.id] = true; }); });
  const out = [];
  const walk = function (folder, depth) {
    const it = folder.getFiles();
    while (it.hasNext() && out.length < 500) {
      const f = it.next(), id = f.getId();
      if (attached[id] || id === cfg.sheetId || f.isTrashed()) continue;
      out.push({ id: id, name: f.getName(), type: f.getMimeType(), url: f.getUrl(), size: f.getSize() });
    }
    if (depth < 2) { const fi = folder.getFolders(); while (fi.hasNext()) walk(fi.next(), depth + 1); }
  };
  walk(DriveApp.getFolderById(cfg.folderId), 0);
  return out;
}

/** Attach a file that's already in Drive: rename it, move it to its month folder, mark the charge done. */
function attachExisting(rowId, fileId, invoiceNo) {
  auth_();
  if (invoiceNo == null) invoiceNo = findInvoiceNo_(ocrText_(fileId));
  return withLock_(function () {
    const sh = sheet_();
    const r = findRow_(sh, rowId);
    const f = DriveApp.getFileById(fileId);
    const original = f.getName();
    f.setName(invoiceName_(r, original, r.files.length));
    f.moveTo(monthFolder_(r.month));
    r.files.push({ id: fileId, name: f.getName(), original: original, url: f.getUrl(), by: me_(), at: new Date().toISOString(), invoiceNo: String(invoiceNo || '') });
    r.status = 'Done';
    touch_(r); writeRow_(sh, r);
    log_('attach existing', r.id, original);
    return clean_(r);
  });
}

/** The uploaded file isn't right (e.g. a receipt, not an invoice): reopen the charge with a reason. */
function requestAgain(rowId, reason) {
  auth_();
  reason = String(reason || '').trim();
  if (!reason) throw new Error('Add a short reason, so the uploader knows what to send.');
  return withLock_(function () {
    const sh = sheet_();
    const r = findRow_(sh, rowId);
    const now = new Date().toISOString();
    r.files.forEach(function (f) { if (!f.rejected) { f.rejected = true; f.rejectedBy = me_(); f.rejectedAt = now; f.rejectReason = reason.slice(0, 300); } });
    r.notes.push({ by: me_(), at: now, text: reason.slice(0, 2000), kind: 'request' });
    r.status = 'Missing';
    touch_(r); writeRow_(sh, r);
    log_('request again', r.id, reason.slice(0, 80));
    return clean_(r);
  });
}

/* ---------- row edits ---------- */

/** Bulk edits from the selection bar. field: 'status' | 'owner' */
function updateMany(rowIds, field, value) {
  auth_();
  if (['status', 'owner'].indexOf(field) < 0) throw new Error('Unknown field.');
  return withLock_(function () {
    const sh = sheet_();
    const ids = {};
    rowIds.forEach(function (id) { ids[id] = true; });
    const out = [];
    readRows_(sh).forEach(function (r) {
      if (!ids[r.id]) return;
      r[field] = field === 'status' ? status_(value) : String(value || '').trim();
      touch_(r); writeRow_(sh, r); out.push(clean_(r));
    });
    log_(field, rowIds.length + ' rows', value);
    return out;
  });
}

/** Done, Lost (the invoice can't be found) or Missing. */
function status_(s) {
  return s === 'Done' || s === 'Lost' ? s : 'Missing';
}

function setStatus(rowId, status) {
  auth_();
  return withLock_(function () {
    const sh = sheet_();
    const r = findRow_(sh, rowId);
    r.status = status_(status);
    touch_(r); writeRow_(sh, r);
    log_('status', r.id, r.status);
    return clean_(r);
  });
}

function addNote(rowId, text) {
  auth_();
  text = String(text || '').trim();
  if (!text) throw new Error('Write something before adding the note.');
  return withLock_(function () {
    const sh = sheet_();
    const r = findRow_(sh, rowId);
    r.notes.push({ by: me_(), at: new Date().toISOString(), text: text.slice(0, 2000) });
    touch_(r); writeRow_(sh, r);
    log_('note', r.id, text.slice(0, 80));
    return clean_(r);
  });
}

/* ---------- who can open Oak ---------- */

function saveAccess(text) {
  auth_();
  return saveAccessRaw_(text);
}

function saveAccessRaw_(text) {
  if (ALLOWED.length) throw new Error('Access is set in the code (ALLOWED in Code.gs).');
  let list = parseAllow_(text);
  if (!list.length) list = [me_().split('@')[1]];
  if (!matches_(me_(), list)) throw new Error('Include your own domain or email, or you\'d lock yourself out.');
  PROPS.setProperty('allowed', list.join(','));
  log_('access', '', list.join(', '));
  return list;
}

/* ---------- monthly CFO summary ---------- */

function saveSummary(email, on) {
  auth_();
  email = String(email || '').trim();
  if (on && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw new Error('Enter a valid email address.');
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'sendMonthlySummary') ScriptApp.deleteTrigger(t);
  });
  if (on) ScriptApp.newTrigger('sendMonthlySummary').timeBased().onMonthDay(1).atHour(9).create();
  PROPS.setProperties({ summaryEmail: email, summaryOn: on ? '1' : '0' });
  log_('summary settings', '', on ? 'on, to ' + email : 'off');
  return { email: email, on: !!on };
}

/** Runs on the 1st of every month at 09:00, or from Settings with "Send now". */
function sendMonthlySummary() {
  auth_();
  const to = PROPS.getProperty('summaryEmail');
  if (!to) throw new Error('Add the CFO\'s email in Settings first.');
  const rows = readRows_().filter(function (r) { return r.status !== 'Done' && r.status !== 'Lost'; });
  const url = ScriptApp.getService().getUrl();
  const esc = function (v) { return String(v == null ? '' : v).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); };
  const money = function (n, cur) { return n == null || isNaN(n) ? '' : (cur || '') + ' ' + Number(n).toFixed(2); };
  const byMonth = {};
  rows.forEach(function (r) { (byMonth[r.month] = byMonth[r.month] || []).push(r); });
  const months = Object.keys(byMonth).sort().reverse();
  const allRows = readRows_();

  let html = '<div style="font-family:Arial,sans-serif;color:#17201C;max-width:640px">';
  html += '<h2 style="font-weight:500;margin:0 0 6px">Missing invoices: ' + rows.length + '</h2>';
  html += rows.length ? '<p style="color:#4B5550;margin:0 0 18px">Charges that still have no invoice, by month.</p>'
                      : '<p style="color:#4B5550">Every charge has an invoice.</p>';
  months.forEach(function (m) {
    const list = byMonth[m];
    const totals = {};
    list.forEach(function (r) {
      const useCharged = r.charged != null && !isNaN(r.charged);
      const cur = useCharged ? (r.chargedCurrency || r.currency) : r.currency;
      totals[cur] = (totals[cur] || 0) + (useCharged ? r.charged : (r.amount || 0));
    });
    const all = {};
    allRows.filter(function (r) { return r.month === m; }).forEach(function (r) {
      const useCharged = r.charged != null && !isNaN(r.charged);
      const cur = useCharged ? (r.chargedCurrency || r.currency) : r.currency;
      all[cur] = (all[cur] || 0) + (useCharged ? r.charged : (r.amount || 0));
    });
    const fmt = function (t) { return Object.keys(t).map(function (c) { return esc(money(t[c], c)); }).join(' + '); };
    html += '<h3 style="font-weight:600;font-size:15px;margin:22px 0 4px">' + esc(m) + ': ' + list.length + ' missing</h3>';
    html += '<p style="color:#4B5550;margin:0 0 8px">Missing ' + fmt(totals) + ' out of ' + fmt(all) + ' charged</p>';
    html += '<table style="border-collapse:collapse;width:100%;font-size:13px">';
    list.slice(0, 40).forEach(function (r) {
      html += '<tr style="border-top:1px solid #E0E3DF"><td style="padding:6px 8px 6px 0;color:#4B5550;white-space:nowrap">' + esc(r.date) + '</td>' +
        '<td style="padding:6px 8px" dir="auto">' + esc(r.merchant) + '</td>' +
        '<td style="padding:6px 8px;text-align:right;white-space:nowrap">' + esc(money(r.amount, r.currency)) + '</td>' +
        '<td style="padding:6px 0 6px 8px;color:#4B5550">' + esc(r.owner || 'Nobody') + '</td></tr>';
    });
    if (list.length > 40) html += '<tr><td colspan="4" style="padding:6px 0;color:#848D88">and ' + (list.length - 40) + ' more</td></tr>';
    html += '</table>';
  });
  if (url) html += '<p style="margin:26px 0 0"><a href="' + url + '" style="background:#17201C;color:#fff;padding:10px 16px;border-radius:7px;text-decoration:none">Open Oak</a></p>';
  html += '</div>';

  MailApp.sendEmail({ to: to, subject: 'Oak: ' + rows.length + ' missing invoices', htmlBody: html, name: 'Oak' });
  log_('summary sent', '', to);
  return { sent: true, count: rows.length };
}

/* ---------- helpers ---------- */

/** Everyone the Drive folder is shared with, for the Responsible list. */
function team_(folder) {
  try {
    const out = {};
    const add = function (u) { const e = u && u.getEmail && u.getEmail(); if (e) out[e.toLowerCase()] = true; };
    add(folder.getOwner());
    folder.getEditors().forEach(add);
    return Object.keys(out).sort();
  } catch (e) { return []; }
}

function me_() { return Session.getActiveUser().getEmail() || ''; }

function allowList_() {
  if (ALLOWED.length) return ALLOWED.map(function (x) { return String(x).trim().toLowerCase(); });
  return (PROPS.getProperty('allowed') || '').split(',').filter(String);
}

function parseAllow_(text) {
  return String(text || '').split(/[\s,;]+/).map(function (x) { return x.trim().toLowerCase().replace(/^@/, ''); })
    .filter(function (x) { return /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(x) || /^[a-z0-9-]+(\.[a-z0-9-]+)+$/.test(x); });
}

function matches_(email, list) {
  email = String(email || '').toLowerCase();
  const domain = email.split('@')[1];
  return !!email && list.some(function (x) { return x === email || x === domain; });
}

/** Before first setup the list is empty, so only a fresh deployment's first visitor (the person deploying) can set it. */
function allowed_() {
  const list = allowList_();
  if (!list.length) return !config_().folderId && !!me_();
  return matches_(me_(), list);
}

function auth_() {
  if (!allowed_()) throw new Error('Your account doesn\'t have access to Oak.');
}

function config_() {
  return { folderId: PROPS.getProperty('folderId'), sheetId: PROPS.getProperty('sheetId') };
}

function sheet_() {
  const c = config_();
  if (!c.sheetId) throw new Error('Not set up yet. Paste a Drive folder link first.');
  return SpreadsheetApp.openById(c.sheetId).getSheetByName('Invoices');
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function readRows_(sh) {
  sh = sh || sheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const vals = sh.getRange(2, 1, last - 1, COLS.length).getDisplayValues();
  return vals.map(function (v, i) {
    const o = { _row: i + 2 };
    COLS.forEach(function (c, j) { o[c] = v[j]; });
    Object.keys(JSON_COLS).forEach(function (c) {
      try { o[c] = o[c] ? JSON.parse(o[c]) : JSON.parse(JSON.stringify(JSON_COLS[c])); }
      catch (e) { o[c] = JSON.parse(JSON.stringify(JSON_COLS[c])); }
    });
    o.amount = o.amount === '' ? null : Number(o.amount);
    o.charged = o.charged === '' ? null : Number(o.charged);
    return o;
  }).filter(function (o) { return o.id; });
}

function toArr_(r) {
  return COLS.map(function (c) {
    if (c in JSON_COLS) return JSON.stringify(r[c] || JSON_COLS[c]);
    return r[c] == null ? '' : String(r[c]);
  });
}

function writeRow_(sh, r) {
  r.invoiceNo = r.files.filter(function (f) { return f.invoiceNo && !f.rejected; }).map(function (f) { return f.invoiceNo; }).join(', ');
  const rng = sh.getRange(r._row, 1, 1, COLS.length);
  rng.setNumberFormat('@');
  rng.setValues([toArr_(r)]);
}

function findRow_(sh, id) {
  const r = readRows_(sh).filter(function (x) { return x.id === id; })[0];
  if (!r) throw new Error('That charge no longer exists. Refresh the page.');
  return r;
}

function clean_(r) { const o = {}; Object.keys(r).forEach(function (k) { if (k !== '_row') o[k] = r[k]; }); return o; }

function touch_(r) { r.updatedBy = me_(); r.updatedAt = new Date().toISOString(); }

function monthFolder_(month) {
  const root = DriveApp.getFolderById(config_().folderId);
  const name = month || 'No month';
  const it = root.getFoldersByName(name);
  return it.hasNext() ? it.next() : root.createFolder(name);
}

/** 2026-10-21 LINEAR APP 19.00 USD.pdf (second file gets " (2)") */
function invoiceName_(r, original, index) {
  const ext = (String(original).match(/\.[A-Za-z0-9]{1,5}$/) || [''])[0].toLowerCase();
  const cur = { '₪': 'ILS', '$': 'USD', '€': 'EUR' }[r.currency] || r.currency || '';
  const amt = r.amount == null ? '' : Number(r.amount).toFixed(2);
  const base = [r.date, String(r.merchant).replace(/\s+/g, ' ').trim(), amt, cur].filter(String).join(' ');
  return (base + (index ? ' (' + (index + 1) + ')' : '') + ext).replace(/[\\/:*?"<>|#%]/g, '-');
}

function parseFolderId_(link) {
  link = String(link || '').trim();
  const m = link.match(/folders\/([\w-]{10,})/) || link.match(/[?&]id=([\w-]{10,})/) || link.match(/^([\w-]{20,})$/);
  return m ? m[1] : null;
}

function log_(action, row, detail) {
  try {
    const ss = SpreadsheetApp.openById(config_().sheetId);
    ss.getSheetByName('Activity').appendRow([new Date().toISOString(), me_(), action, row, detail]);
  } catch (e) { /* logging is best effort */ }
}
