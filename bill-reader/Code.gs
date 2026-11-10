/**
 * Bill Reader: reads bill photos sent through a Google Form and logs them to a Google Sheet.
 *
 * Flow (runs automatically on every form submission):
 *   1. A rep submits the form with one or more bill photos (or PDFs).
 *   2. Each file is sent to the Gemini API, which returns the bill number, date,
 *      customer, hospital, line items and total as JSON.
 *   3. The file is renamed to the bill number and moved to the bills folder
 *      (one subfolder per month).
 *   4. One row is added to the "Bills" sheet and one row per item to the "Bill items" sheet.
 *   5. Rows that need a human look (duplicate bill number, unreadable field,
 *      items not adding up to the total, read error) get Status = CHECK or ERROR.
 *
 * Paste this whole file into Extensions > Apps Script of the form's response spreadsheet,
 * then run setup() once. See SETUP.md for the step-by-step guide.
 *
 * The Gemini API key is stored in Script properties (GEMINI_API_KEY), never in this file.
 */

// ---------------------------------------------------------------------------
// Settings: change these to match your form.
// ---------------------------------------------------------------------------
const CONFIG = {
  // Form question titles. They must match the form exactly (same spelling and spaces).
  Q_REP: 'ชื่อผู้แทน',
  Q_PHOTO: 'รูปบิล',
  Q_NOTE: 'หมายเหตุ', // optional question; leave the form without it if you don't need it

  // Gemini model. Flash models are fast and fit the free tier.
  GEMINI_MODEL: 'gemini-2.5-flash',

  // Drive folder for the renamed bills. Leave blank and setup() creates one for you.
  FOLDER_ID: '',
  MONTH_SUBFOLDERS: true, // put each bill in a "YYYY-MM" subfolder by bill date

  SHEET_BILLS: 'Bills',
  SHEET_ITEMS: 'Bill items',

  // Flag the bill when the item amounts and the total differ by more than this.
  TOTAL_TOLERANCE: 1,

  // Files above this size are not sent to Gemini (inline request limit is about 20 MB).
  MAX_FILE_MB: 15,
};

const BILL_HEADERS = [
  'Submitted at', 'Rep name', 'Bill number', 'Bill date', 'Customer', 'Hospital',
  'Items', 'Total amount', 'File', 'Status', 'Check notes', 'Rep note', 'File ID',
];
const ITEM_HEADERS = [
  'Bill number', 'Bill date', 'Customer', 'Hospital', 'Item', 'Qty', 'Unit price',
  'Amount', 'Rep name', 'File ID',
];
// 1-based column numbers in the Bills sheet.
const COL = {
  SUBMITTED: 1, REP: 2, BILL_NO: 3, DATE: 4, CUSTOMER: 5, HOSPITAL: 6, ITEMS: 7,
  TOTAL: 8, FILE: 9, STATUS: 10, NOTES: 11, REP_NOTE: 12, FILE_ID: 13,
};
const ITEM_FILE_ID_COL = 10;

const STATUS_COLORS = { OK: '#d9f2e3', CHECK: '#fff2cc', ERROR: '#f8d7da' };

// ---------------------------------------------------------------------------
// One-time setup and menu
// ---------------------------------------------------------------------------

/** Run once from the Apps Script editor (or the Bill reader menu). Safe to run again. */
function setup() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();

  // 1. Sheets with headers.
  const bills = ensureSheet_(ss, CONFIG.SHEET_BILLS, BILL_HEADERS);
  ensureSheet_(ss, CONFIG.SHEET_ITEMS, ITEM_HEADERS);
  // Keep bill numbers as text so leading zeros survive.
  bills.getRange(2, COL.BILL_NO, bills.getMaxRows() - 1, 1).setNumberFormat('@');
  ss.getSheetByName(CONFIG.SHEET_ITEMS)
    .getRange(2, 1, ss.getSheetByName(CONFIG.SHEET_ITEMS).getMaxRows() - 1, 1).setNumberFormat('@');

  // 2. Gemini API key.
  if (!props.getProperty('GEMINI_API_KEY')) {
    const res = ui.prompt('Gemini API key',
      'Paste the API key from Google AI Studio (aistudio.google.com > Get API key):',
      ui.ButtonSet.OK_CANCEL);
    if (res.getSelectedButton() !== ui.Button.OK || !res.getResponseText().trim()) {
      ui.alert('Setup stopped: no API key was entered. Run setup again when you have one.');
      return;
    }
    props.setProperty('GEMINI_API_KEY', res.getResponseText().trim());
  }

  // 3. Bills folder.
  getBillsFolder_();

  // 4. Form-submit trigger (replace any old one so it never runs twice).
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'onBillSubmit')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('onBillSubmit').forSpreadsheet(ss).onFormSubmit().create();

  // 5. Quick test of the key.
  let keyMsg;
  try {
    callGemini_([{ text: 'Reply with the word OK.' }], null);
    keyMsg = 'Gemini API key works.';
  } catch (err) {
    keyMsg = 'Warning: the Gemini test call failed: ' + err.message +
      '\nCheck the key with "Bill reader > Change Gemini API key".';
  }

  ui.alert('Bill reader is set up.\n\n' + keyMsg +
    '\nBills folder: ' + getBillsFolder_().getUrl() +
    '\n\nNew form submissions will now be processed automatically.');
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Bill reader')
    .addItem('Re-read selected bills', 'rereadSelected')
    .addItem('Process all ERROR rows again', 'retryErrors')
    .addSeparator()
    .addItem('Change Gemini API key', 'changeApiKey')
    .addItem('Run setup', 'setup')
    .addToUi();
}

function changeApiKey() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('Gemini API key', 'Paste the new API key:', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() === ui.Button.OK && res.getResponseText().trim()) {
    PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', res.getResponseText().trim());
    ui.alert('Saved.');
  }
}

// ---------------------------------------------------------------------------
// Automatic processing
// ---------------------------------------------------------------------------

/** Installable trigger: runs on every form submission. */
function onBillSubmit(e) {
  const answers = e.namedValues || {};
  const rep = first_(answers[CONFIG.Q_REP]);
  const note = first_(answers[CONFIG.Q_NOTE]);
  const submittedAt = first_(answers['Timestamp']) || first_(answers['ประทับเวลา']) || new Date();
  const fileIds = extractFileIds_((answers[CONFIG.Q_PHOTO] || []).join(','));

  if (!fileIds.length) {
    appendRecord_({
      submittedAt, rep, note, fileId: '',
      result: { status: 'ERROR', checks: ['No file found in question "' + CONFIG.Q_PHOTO + '". Check CONFIG.Q_PHOTO matches the form.'] },
    });
    return;
  }
  fileIds.forEach(fileId => {
    const result = readFile_(fileId);
    appendRecord_({ submittedAt, rep, note, fileId, result });
  });
}

/** Menu: read the selected rows of the Bills sheet again (e.g. after fixing a photo). */
function rereadSelected() {
  const sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getName() !== CONFIG.SHEET_BILLS) {
    SpreadsheetApp.getUi().alert('Select rows in the "' + CONFIG.SHEET_BILLS + '" sheet first.');
    return;
  }
  const rows = [];
  sheet.getActiveRangeList().getRanges().forEach(r => {
    for (let i = 0; i < r.getNumRows(); i++) rows.push(r.getRow() + i);
  });
  rereadRows_(sheet, rows.filter(r => r > 1));
}

/** Menu: try every row with Status = ERROR again (e.g. after a Gemini outage). */
function retryErrors() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.SHEET_BILLS);
  const last = sheet.getLastRow();
  if (last < 2) return;
  const statuses = sheet.getRange(2, COL.STATUS, last - 1, 1).getValues();
  const rows = [];
  statuses.forEach((s, i) => { if (s[0] === 'ERROR') rows.push(i + 2); });
  rereadRows_(sheet, rows);
}

function rereadRows_(sheet, rows) {
  let done = 0;
  rows.forEach(row => {
    const fileId = String(sheet.getRange(row, COL.FILE_ID).getValue() || '');
    if (!fileId) return;
    const result = readFile_(fileId);
    updateRecord_(sheet, row, fileId, result);
    done++;
  });
  SpreadsheetApp.getActive().toast(done + ' bill(s) read again.', 'Bill reader');
}

// ---------------------------------------------------------------------------
// Reading a bill
// ---------------------------------------------------------------------------

/** Reads one Drive file with Gemini and runs the checks. Never throws. */
function readFile_(fileId) {
  try {
    const file = DriveApp.getFileById(fileId);
    if (file.getSize() > CONFIG.MAX_FILE_MB * 1024 * 1024) {
      return { status: 'ERROR', checks: ['File is larger than ' + CONFIG.MAX_FILE_MB + ' MB'] };
    }
    const data = readBill_(toSupportedBlob_(file.getBlob()));
    const items = (data.items || []).filter(it => it && it.description);
    data.items = items;

    const checks = [];
    if (data.is_bill === false) checks.push('Does not look like a bill');
    if (!data.bill_number) checks.push('Bill number not found');
    if (!data.bill_date) checks.push('Date not found');
    if (!data.customer && !data.hospital) checks.push('Customer not found');
    if (data.total_amount === null || data.total_amount === undefined) checks.push('Total not found');
    if (data.bill_date && !/^\d{4}-\d{2}-\d{2}$/.test(data.bill_date)) checks.push('Date format unclear');
    const amounts = items.map(it => it.amount).filter(a => typeof a === 'number');
    if (amounts.length && amounts.length === items.length && typeof data.total_amount === 'number') {
      const sum = amounts.reduce((a, b) => a + b, 0);
      // Totals often include VAT, so only flag when the total is smaller than the items
      // or differs from both the plain sum and the sum plus 7% VAT.
      const diffPlain = Math.abs(sum - data.total_amount);
      const diffVat = Math.abs(sum * 1.07 - data.total_amount);
      if (diffPlain > CONFIG.TOTAL_TOLERANCE && diffVat > CONFIG.TOTAL_TOLERANCE) {
        checks.push('Items add up to ' + sum.toFixed(2) + ', total says ' + data.total_amount);
      }
    }
    (data.unclear_fields || []).forEach(f => checks.push('Unclear: ' + f));

    return { data, checks, status: checks.length ? 'CHECK' : 'OK' };
  } catch (err) {
    return { status: 'ERROR', checks: ['Could not read: ' + err.message] };
  }
}

const BILL_PROMPT = [
  'You read photos or scans of sales bills, invoices, tax invoices and delivery notes,',
  'mostly in Thai, sometimes in English. Extract the fields in the JSON schema.',
  'Rules:',
  '- bill_number: the document number (เลขที่ / เลขที่ใบกำกับภาษี / Invoice No.). Copy it exactly, including letters, dashes and leading zeros.',
  '- bill_date: the document date as YYYY-MM-DD in the Gregorian calendar.',
  '  Thai dates often use the Buddhist year (พ.ศ., e.g. 2569): subtract 543. Two-digit Thai years like 69 mean 2569.',
  '  Thai month names: ม.ค.=01 ก.พ.=02 มี.ค.=03 เม.ย.=04 พ.ค.=05 มิ.ย.=06 ก.ค.=07 ส.ค.=08 ก.ย.=09 ต.ค.=10 พ.ย.=11 ธ.ค.=12.',
  '- customer: the billed-to customer name (ลูกค้า / นาม / Sold to / Bill to).',
  '- hospital: the hospital or clinic name if one appears anywhere (it may be the same as the customer); otherwise null.',
  '- items: every line item with description, quantity, unit_price and amount (line total) as plain numbers.',
  '- total_amount: the final amount to pay (จำนวนเงินรวมทั้งสิ้น / Grand total), including VAT.',
  '- Numbers: no commas or currency signs. Convert Thai digits (๐-๙) to 0-9.',
  '- Never guess. If a field is missing or unreadable, use null and add its name to unclear_fields.',
  '  Also add a field to unclear_fields when you can read it but are not sure (blurry, handwritten, cut off).',
  '- is_bill: false if the image is not a bill/invoice/receipt at all.',
].join('\n');

const BILL_SCHEMA = {
  type: 'OBJECT',
  properties: {
    is_bill: { type: 'BOOLEAN' },
    bill_number: { type: 'STRING', nullable: true },
    bill_date: { type: 'STRING', nullable: true, description: 'YYYY-MM-DD, Gregorian calendar' },
    customer: { type: 'STRING', nullable: true },
    hospital: { type: 'STRING', nullable: true },
    items: {
      type: 'ARRAY',
      items: {
        type: 'OBJECT',
        properties: {
          description: { type: 'STRING' },
          quantity: { type: 'NUMBER', nullable: true },
          unit_price: { type: 'NUMBER', nullable: true },
          amount: { type: 'NUMBER', nullable: true },
        },
        required: ['description', 'quantity', 'unit_price', 'amount'],
      },
    },
    total_amount: { type: 'NUMBER', nullable: true },
    unclear_fields: { type: 'ARRAY', items: { type: 'STRING' } },
  },
  required: ['is_bill', 'bill_number', 'bill_date', 'customer', 'hospital', 'items', 'total_amount', 'unclear_fields'],
};

function readBill_(blob) {
  const parts = [
    { text: BILL_PROMPT },
    { inlineData: { mimeType: blob.getContentType(), data: Utilities.base64Encode(blob.getBytes()) } },
  ];
  return JSON.parse(callGemini_(parts, BILL_SCHEMA));
}

/** Calls Gemini and returns the text of the first answer. Retries on rate limits and outages. */
function callGemini_(parts, schema) {
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!key) throw new Error('GEMINI_API_KEY is not set. Run setup().');
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' +
    encodeURIComponent(CONFIG.GEMINI_MODEL) + ':generateContent';
  const generationConfig = { temperature: 0 };
  if (schema) {
    generationConfig.responseMimeType = 'application/json';
    generationConfig.responseSchema = schema;
  }
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-goog-api-key': key },
    payload: JSON.stringify({ contents: [{ role: 'user', parts }], generationConfig }),
    muteHttpExceptions: true,
  };

  let lastError = '';
  for (let attempt = 0; attempt < 4; attempt++) {
    const res = UrlFetchApp.fetch(url, options);
    const code = res.getResponseCode();
    const body = res.getContentText();
    if (code === 200) {
      const json = JSON.parse(body);
      const cand = (json.candidates || [])[0];
      const text = cand && cand.content && (cand.content.parts || []).map(p => p.text || '').join('');
      if (!text) throw new Error('Gemini returned no answer (' + ((cand && cand.finishReason) || 'blocked') + ')');
      return text;
    }
    lastError = 'Gemini HTTP ' + code + ': ' + body.slice(0, 300);
    if (code !== 429 && code < 500) break; // bad key, bad request: retrying will not help
    Utilities.sleep(5000 * Math.pow(2, attempt)); // 5s, 10s, 20s, 40s
  }
  throw new Error(lastError);
}

/** Gemini accepts JPEG, PNG, WEBP, HEIC/HEIF and PDF; convert anything else to JPEG. */
function toSupportedBlob_(blob) {
  const ok = ['image/jpeg', 'image/png', 'image/webp', 'image/heic', 'image/heif', 'application/pdf'];
  if (ok.indexOf(blob.getContentType()) >= 0) return blob;
  return blob.getAs('image/jpeg');
}

// ---------------------------------------------------------------------------
// Writing to the sheet and Drive
// ---------------------------------------------------------------------------

function appendRecord_(rec) {
  const lock = LockService.getScriptLock();
  lock.waitLock(120000); // one writer at a time so duplicate checks are reliable
  try {
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName(CONFIG.SHEET_BILLS);
    const row = sheet.getLastRow() + 1;
    const result = withDuplicateCheck_(sheet, rec.result, row);
    const fileUrl = rec.fileId ? fileAndMove_(rec.fileId, result) : '';
    writeBillRow_(sheet, row, rec.submittedAt, rec.rep, rec.note, rec.fileId, fileUrl, result);
    writeItemRows_(ss, rec.rep, rec.fileId, result);
  } finally {
    lock.releaseLock();
  }
}

function updateRecord_(sheet, row, fileId, result) {
  const lock = LockService.getScriptLock();
  lock.waitLock(120000);
  try {
    const ss = sheet.getParent();
    const current = sheet.getRange(row, 1, 1, BILL_HEADERS.length).getValues()[0];
    result = withDuplicateCheck_(sheet, result, row);
    const fileUrl = fileAndMove_(fileId, result);
    writeBillRow_(sheet, row, current[COL.SUBMITTED - 1], current[COL.REP - 1],
      current[COL.REP_NOTE - 1], fileId, fileUrl, result);
    deleteItemRows_(ss, fileId);
    writeItemRows_(ss, current[COL.REP - 1], fileId, result);
  } finally {
    lock.releaseLock();
  }
}

/** Adds a CHECK note when the bill number is already in another row. */
function withDuplicateCheck_(sheet, result, ownRow) {
  const billNo = result.data && result.data.bill_number;
  if (!billNo) return result;
  const last = sheet.getLastRow();
  if (last < 2) return result;
  const key = normalizeBillNo_(billNo);
  const values = sheet.getRange(2, COL.BILL_NO, last - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    const row = i + 2;
    if (row !== ownRow && normalizeBillNo_(values[i][0]) === key) {
      result.checks.push('Duplicate bill number (also in row ' + row + ')');
      result.duplicate = true;
      if (result.status === 'OK') result.status = 'CHECK';
      break;
    }
  }
  return result;
}

/** Renames the file to the bill number and moves it into the bills folder. Returns its URL. */
function fileAndMove_(fileId, result) {
  const file = DriveApp.getFileById(fileId);
  const data = result.data || {};
  const ext = (file.getName().match(/\.[A-Za-z0-9]{2,5}$/) || [''])[0];
  const stamp = Utilities.formatDate(new Date(), tz_(), 'yyyyMMdd-HHmmss');
  let name;
  if (data.bill_number) {
    name = 'BILL_' + safeName_(data.bill_number);
    if (data.customer || data.hospital) name += '_' + safeName_(data.hospital || data.customer).slice(0, 40);
    if (result.duplicate) name += '_DUPLICATE_' + stamp;
  } else {
    name = 'UNREAD_' + stamp;
  }
  file.setName(name + ext);

  let folder = getBillsFolder_();
  if (CONFIG.MONTH_SUBFOLDERS) {
    const month = (data.bill_date && /^\d{4}-\d{2}/.test(data.bill_date))
      ? data.bill_date.slice(0, 7)
      : Utilities.formatDate(new Date(), tz_(), 'yyyy-MM');
    folder = getOrCreateSubfolder_(folder, month);
  }
  file.moveTo(folder);
  return file.getUrl();
}

function writeBillRow_(sheet, row, submittedAt, rep, note, fileId, fileUrl, result) {
  const d = result.data || {};
  const itemsText = (d.items || []).map(it => {
    let s = it.description;
    if (it.quantity !== null && it.quantity !== undefined) s += ' x' + it.quantity;
    if (it.amount !== null && it.amount !== undefined) s += ' = ' + it.amount;
    return s;
  }).join('\n');

  const values = [
    submittedAt, rep, d.bill_number || '', d.bill_date || '', d.customer || '', d.hospital || '',
    itemsText, numOrBlank_(d.total_amount), '', result.status, (result.checks || []).join('; '),
    note || '', fileId || '',
  ].map(safeCell_);

  sheet.getRange(row, COL.BILL_NO).setNumberFormat('@');
  sheet.getRange(row, 1, 1, values.length).setValues([values]);
  if (fileUrl) sheet.getRange(row, COL.FILE).setFormula('=HYPERLINK("' + fileUrl + '","Open")');
  sheet.getRange(row, COL.TOTAL).setNumberFormat('#,##0.00');
  sheet.getRange(row, COL.STATUS).setBackground(STATUS_COLORS[result.status] || null);
}

function writeItemRows_(ss, rep, fileId, result) {
  const d = result.data;
  if (!d || !d.items || !d.items.length) return;
  const sheet = ss.getSheetByName(CONFIG.SHEET_ITEMS);
  const rows = d.items.map(it => [
    d.bill_number || '', d.bill_date || '', d.customer || '', d.hospital || '', it.description,
    numOrBlank_(it.quantity), numOrBlank_(it.unit_price), numOrBlank_(it.amount), rep || '', fileId,
  ].map(safeCell_));
  const start = sheet.getLastRow() + 1;
  sheet.getRange(start, 1, rows.length, 1).setNumberFormat('@');
  sheet.getRange(start, 1, rows.length, ITEM_HEADERS.length).setValues(rows);
}

function deleteItemRows_(ss, fileId) {
  const sheet = ss.getSheetByName(CONFIG.SHEET_ITEMS);
  const last = sheet.getLastRow();
  if (last < 2) return;
  const ids = sheet.getRange(2, ITEM_FILE_ID_COL, last - 1, 1).getValues();
  for (let i = ids.length - 1; i >= 0; i--) {
    if (String(ids[i][0]) === fileId) sheet.deleteRow(i + 2);
  }
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function ensureSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sheet.setFrozenRows(1);
  return sheet;
}

function getBillsFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = CONFIG.FOLDER_ID || props.getProperty('BILLS_FOLDER_ID');
  if (id) return DriveApp.getFolderById(id);
  const folder = DriveApp.createFolder('Bills (auto-renamed)');
  props.setProperty('BILLS_FOLDER_ID', folder.getId());
  return folder;
}

function getOrCreateSubfolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

/** Pulls Drive file IDs out of the form's upload answer ("https://drive.google.com/open?id=..."). */
function extractFileIds_(text) {
  const ids = [];
  const re = /(?:id=|\/d\/)([A-Za-z0-9_-]{20,})/g;
  let m;
  while ((m = re.exec(text)) !== null) if (ids.indexOf(m[1]) < 0) ids.push(m[1]);
  return ids;
}

function normalizeBillNo_(v) {
  return String(v || '').toUpperCase().replace(/\s+/g, '');
}

function safeName_(s) {
  return String(s).replace(/[\\/:*?"<>|#%\r\n\t]+/g, '-').replace(/\s+/g, ' ').trim();
}

/** Stops text read from a bill from being treated as a spreadsheet formula. */
function safeCell_(v) {
  return (typeof v === 'string' && /^[=+\-@]/.test(v)) ? "'" + v : v;
}

function numOrBlank_(v) {
  return typeof v === 'number' && isFinite(v) ? v : '';
}

function first_(arr) {
  return arr && arr.length ? String(arr[0]).trim() : '';
}

function tz_() {
  return SpreadsheetApp.getActive().getSpreadsheetTimeZone() || Session.getScriptTimeZone();
}
