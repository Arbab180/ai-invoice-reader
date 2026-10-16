/**
 * Split the invoice scans into one PDF per invoice, sorted into a folder per restaurant.
 * Everything stays in Google Drive.
 *
 * SETUP (once)
 *   1. Go to script.google.com > New project. Paste this whole file over the sample code.
 *   2. Left sidebar: Services (+) > Drive API > Add.
 *   3. Save.
 *
 * STEP 1: choose `step1_splitAndRead` in the function dropdown and click Run. Approve permissions.
 *   It splits every scan into single pages and reads the text of each page.
 *   It carries on by itself in the background until all pages are done (runs of ~4 minutes, about an hour in total for ~350 pages).
 *   When it finishes, the "page-text" sheet in Scans/_invoice work has one row per page.
 *   Tell Claude it's done. Claude then adds a "page-mapping" sheet to the same folder.
 *
 * STEP 2: run `step2_buildFolders`.
 *   It creates "Invoices sorted by restaurant (PDFs)" in Scans, with a folder per restaurant
 *   and a "Needs check" folder, and saves each invoice there as its own PDF.
 *
 * OPTIONAL: run `cleanup` afterwards to move the _invoice work folder to the bin.
 */

const SOURCE_FOLDER_ID = 'YOUR_TEST_SCANS_FOLDER_ID'; // folder of test scans
const PARENT_FOLDER_ID = 'YOUR_SCANS_FOLDER_ID'; // Scans
const WORK_NAME = '_invoice work';
const OUTPUT_NAME = 'Invoices sorted by restaurant (PDFs)';
const TEXT_SHEET = 'page-text';
const MAPPING_SHEET = 'page-mapping';
const TIME_LIMIT_MS = 4 * 60 * 1000; // Google stops scripts at 6 minutes; leave room for the last page

// ---------- Step 1 ----------

async function step1_splitAndRead() {
  const start = Date.now();
  removeTriggers_('step1_splitAndRead');
  const work = childFolder_(DriveApp.getFolderById(PARENT_FOLDER_ID), WORK_NAME);
  const pagesRoot = childFolder_(work, 'pages');
  const sheet = textSheet_(work);
  const done = new Set(sheet.getDataRange().getValues().slice(1).map(r => r[0] + '|' + r[1]));
  await loadPdfLib_();

  const scans = listPdfs_(DriveApp.getFolderById(SOURCE_FOLDER_ID));
  for (const scan of scans) {
    const name = scan.getName();
    const pageFolder = childFolder_(pagesRoot, name.replace(/\.pdf$/i, ''));
    let src = null;
    let count = Number(PropertiesService.getScriptProperties().getProperty('pages:' + name)) || 0;

    for (let p = 1; count === 0 || p <= count; p++) {
      if (count && done.has(name + '|' + p)) continue;
      if (Date.now() - start > TIME_LIMIT_MS) return continueLater_('step1_splitAndRead');

      if (!src) {
        src = await PDFLib.PDFDocument.load(toUint8_(scan.getBlob()), { ignoreEncryption: true });
        count = src.getPageCount();
        PropertiesService.getScriptProperties().setProperty('pages:' + name, String(count));
        if (done.has(name + '|' + p)) continue;
      }

      const pageName = 'p' + String(p).padStart(3, '0') + '.pdf';
      let pageFile = firstFile_(pageFolder, pageName);
      if (!pageFile) {
        const one = await PDFLib.PDFDocument.create();
        const [copied] = await one.copyPages(src, [p - 1]);
        one.addPage(copied);
        const bytes = await one.save();
        pageFile = pageFolder.createFile(pdfBlob_(bytes, pageName));
      }
      sheet.appendRow([name, p, pageFile.getId(), ocr_(pageFile).slice(0, 45000)]);
      done.add(name + '|' + p);
    }
  }
  removeTriggers_('step1_splitAndRead');
  Logger.log('Step 1 finished. Tell Claude the page text is ready.');
}

function ocr_(file) {
  const doc = Drive.Files.create(
    { name: 'ocr-tmp', mimeType: MimeType.GOOGLE_DOCS },
    file.getBlob(),
    { ocrLanguage: 'el' }
  );
  const text = DocumentApp.openById(doc.id).getBody().getText();
  Drive.Files.remove(doc.id);
  return text;
}

// ---------- Step 2 ----------

async function step2_buildFolders() {
  const start = Date.now();
  removeTriggers_('step2_buildFolders');
  const parent = DriveApp.getFolderById(PARENT_FOLDER_ID);
  const work = childFolder_(parent, WORK_NAME);
  const mapping = firstFile_(work, MAPPING_SHEET);
  if (!mapping) throw new Error('No "' + MAPPING_SHEET + '" sheet in ' + WORK_NAME + ' yet. Ask Claude to add it.');
  const pageIds = {};
  textSheet_(work).getDataRange().getValues().slice(1).forEach(r => { pageIds[r[0] + '|' + r[1]] = r[2]; });

  const rows = SpreadsheetApp.open(mapping).getSheets()[0].getDataRange().getValues().slice(1);
  const output = childFolder_(parent, OUTPUT_NAME);
  await loadPdfLib_();

  for (const [scan, fromPage, toPage, folderName, fileName] of rows) {
    if (!scan) continue;
    const folder = childFolder_(output, String(folderName));
    const outName = safeName_(fileName) + '.pdf';
    if (firstFile_(folder, outName)) continue;
    if (Date.now() - start > TIME_LIMIT_MS) return continueLater_('step2_buildFolders');

    const merged = await PDFLib.PDFDocument.create();
    for (let p = Number(fromPage); p <= Number(toPage); p++) {
      const id = pageIds[scan + '|' + p];
      if (!id) throw new Error('Missing page ' + p + ' of ' + scan + '. Run step 1 again.');
      const one = await PDFLib.PDFDocument.load(toUint8_(DriveApp.getFileById(id).getBlob()));
      const [copied] = await merged.copyPages(one, [0]);
      merged.addPage(copied);
    }
    folder.createFile(pdfBlob_(await merged.save(), outName));
  }
  removeTriggers_('step2_buildFolders');
  Logger.log('Step 2 finished. Invoices are in "' + OUTPUT_NAME + '".');
}

function cleanup() {
  const it = DriveApp.getFolderById(PARENT_FOLDER_ID).getFoldersByName(WORK_NAME);
  while (it.hasNext()) it.next().setTrashed(true);
}

// ---------- Helpers ----------

// Apps Script gives bytes as signed numbers (-128..127); pdf-lib needs a Uint8Array, and back again for saving.
function toUint8_(blob) {
  return new Uint8Array(blob.getBytes());
}

function pdfBlob_(bytes, name) {
  return Utilities.newBlob(Array.from(new Int8Array(bytes.buffer, bytes.byteOffset, bytes.byteLength)), MimeType.PDF, name);
}

async function loadPdfLib_() {
  if (typeof PDFLib !== 'undefined') return;
  const src = UrlFetchApp.fetch('https://cdnjs.cloudflare.com/ajax/libs/pdf-lib/1.17.1/pdf-lib.min.js').getContentText();
  // Apps Script has no setTimeout; pdf-lib only uses it to yield, so run the callback straight away.
  eval(src.replace(/setTimeout\(.*?,.*?(\d*?)\)/g, 'Utilities.sleep($1);return t();'));
}

function continueLater_(fn) {
  ScriptApp.newTrigger(fn).timeBased().after(60 * 1000).create();
  Logger.log('Paused to stay under the time limit. It will continue by itself in a minute.');
}

function removeTriggers_(fn) {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === fn).forEach(t => ScriptApp.deleteTrigger(t));
}

function textSheet_(work) {
  const file = firstFile_(work, TEXT_SHEET);
  if (file) return SpreadsheetApp.open(file).getSheets()[0];
  const ss = SpreadsheetApp.create(TEXT_SHEET);
  DriveApp.getFileById(ss.getId()).moveTo(work);
  const sheet = ss.getSheets()[0];
  sheet.appendRow(['scan', 'page', 'page_file_id', 'text']);
  return sheet;
}

function listPdfs_(folder) {
  const out = [];
  const it = folder.getFilesByType(MimeType.PDF);
  while (it.hasNext()) out.push(it.next());
  return out.sort((a, b) => a.getName().localeCompare(b.getName()));
}

function childFolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function firstFile_(folder, name) {
  const it = folder.getFilesByName(name);
  return it.hasNext() ? it.next() : null;
}

function safeName_(s) {
  return String(s).replace(/[\\/:*?"<>|]+/g, '-').trim().slice(0, 150);
}
