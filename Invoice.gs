/* Reading invoices: Drive's OCR turns a PDF or photo into text, then we look for the invoice number. */

const INVOICE_LABEL = new RegExp([
  '\\b(?:tax\\s+)?(?:invoice|receipt|bill|document)\\s*(?:no\\.?|number|num\\.?|nr\\.?|id|#|:)',
  'מספר\\s*(?:חשבונית|מסמך|קבלה)',
  'מס[\'׳]?\\s*(?:חשבונית|מסמך|קבלה)',
  'חשבונית(?:\\s*מס)?(?:\\s*/\\s*קבלה)?(?:\\s*(?:מספר|מס[\'׳]?))?',
  'קבלה(?:\\s*(?:מספר|מס[\'׳]?))?'
].join('|'), 'i');

/** Saves a dropped file in the Drive folder, not attached yet, so it can be read. */
function stageFile(file) {
  auth_();
  const blob = Utilities.newBlob(Utilities.base64Decode(file.data), file.type || 'application/octet-stream', file.name);
  const f = DriveApp.getFolderById(config_().folderId).createFile(blob);
  f.setDescription('Uploaded by ' + me_() + '. Original name: ' + file.name);
  return { id: f.getId(), name: f.getName(), type: f.getMimeType(), url: f.getUrl() };
}

/** The text of a PDF or image in Drive, and the invoice number found in it. */
function readInvoice(fileId) {
  auth_();
  const text = ocrText_(fileId);
  return { text: text.slice(0, 20000), invoiceNo: findInvoiceNo_(text) };
}

function ocrText_(fileId) {
  const f = DriveApp.getFileById(fileId);
  if (!/pdf|image/.test(f.getMimeType()) || f.getSize() > 10 * 1024 * 1024) return '';
  let doc = null;
  try {
    doc = Drive.Files.copy({ name: 'Oak is reading this file', mimeType: MimeType.GOOGLE_DOCS }, fileId);
    return DocumentApp.openById(doc.id).getBody().getText();
  } catch (e) {
    console.error('Could not read ' + f.getName() + ': ' + e.message);
    return '';
  } finally {
    if (doc) try { Drive.Files.remove(doc.id); } catch (e) { /* left in Drive, harmless */ }
  }
}

/** Reads the invoice number of files that don't have one yet. Pass back the ids already tried, until left is 0. */
function fillInvoiceNos(tried) {
  auth_();
  const skip = {}, found = {}, todo = [], start = Date.now();
  (tried || []).forEach(function (id) { skip[id] = true; });
  readRows_().forEach(function (r) { r.files.forEach(function (f) { if (!f.invoiceNo && !f.gone && !skip[f.id]) todo.push(f.id); }); });
  let i = 0;
  for (; i < todo.length && Date.now() - start < 30000; i++) {
    try { found[todo[i]] = findInvoiceNo_(ocrText_(todo[i])); } catch (e) { found[todo[i]] = ''; }
  }
  return withLock_(function () {
    const sh = sheet_(), changed = [];
    readRows_(sh).forEach(function (r) {
      let ch = false;
      r.files.forEach(function (f) { if (found[f.id]) { f.invoiceNo = found[f.id]; ch = true; } });
      if (ch) { writeRow_(sh, r); changed.push(clean_(r)); }
    });
    return { left: todo.length - i, tried: todo.slice(0, i), found: Object.keys(found).filter(function (k) { return found[k]; }).length, rows: changed };
  });
}

/** The number next to an "Invoice no." style label, on the same line (either side, for right-to-left text) or the next. */
function findInvoiceNo_(text) {
  const lines = String(text || '').split(/\r?\n/);
  for (let i = 0; i < lines.length; i++) {
    const m = INVOICE_LABEL.exec(lines[i]);
    if (!m) continue;
    const hit = idIn_(lines[i].slice(m.index + m[0].length)) || idIn_(lines[i].slice(0, m.index)) || idIn_(lines[i + 1] || '');
    if (hit) return hit;
  }
  return '';
}

function idIn_(s) {
  const tokens = s.match(/[A-Za-z]{0,5}[-#]?\d[\w\-\/.]*/g) || [];
  for (let i = 0; i < tokens.length; i++) {
    const t = tokens[i].replace(/[.\-\/]+$/, '').replace(/^#/, '');
    if ((t.match(/\d/g) || []).length < 3) continue;
    if (/^\d{1,4}[.\/-]\d{1,2}[.\/-]\d{1,4}$/.test(t)) continue;
    if (/^\d[\d,]*\.\d{2}$/.test(t)) continue;
    return t;
  }
  return '';
}
