/**
 * Marsa Event — booth booking backend (Google Apps Script, bound to a Google Sheet)
 *
 * The website sends each booking here; it is written as a new row in the "Bookings" sheet.
 * Status column (K): قيد المراجعة = pending (orange on the map)
 *                    مؤكد         = confirmed (grey / booked on the map)
 *                    ملغي          = cancelled (booth becomes available again)
 *
 * Setup: see SETUP.md in the repository.
 */

const SHEET_NAME = 'Bookings';
const HEADERS = ['Timestamp', 'Booth', 'Tier', 'Price (SAR)', 'Brand', 'Contact name',
                 'WhatsApp', 'Instagram', 'Product type', 'Brand profile link', 'Status', 'Notes'];
const STATUS_PENDING = 'قيد المراجعة';
const STATUS_CONFIRMED = 'مؤكد';
const STATUS_CANCELLED = 'ملغي';

// Booth catalogue — must match the website (41 booths)
const FIRST_TIER = ['Y1', 'Y25', 'Z1', 'Z7'];
function boothInfo(id) {
  if (/^VIP[1-4]$/.test(id)) return { tier: 'VIP', price: 8500 };
  if (FIRST_TIER.indexOf(id) !== -1) return { tier: 'First tier', price: 7500 };
  if (/^Y([1-9]|1\d|2[0-5])$/.test(id) || /^Z([1-9]|1[0-2])$/.test(id)) return { tier: 'Standard', price: 5000 };
  return null;
}

/** Run once from the Apps Script editor: creates the sheet, headers and the Status dropdown. */
function setup() {
  const sh = getSheet_();
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList([STATUS_PENDING, STATUS_CONFIRMED, STATUS_CANCELLED], true)
    .setAllowInvalid(false).build();
  sh.getRange(2, 11, sh.getMaxRows() - 1, 1).setDataValidation(rule);
  sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#450693').setFontColor('#ffffff');
  sh.setColumnWidths(1, HEADERS.length, 150);
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
  }
  return sh;
}

/** booth -> 'pending' | 'booked' (cancelled rows are ignored) */
function statuses_() {
  const values = getSheet_().getDataRange().getValues();
  const out = {};
  for (let i = 1; i < values.length; i++) {
    const booth = String(values[i][1]).trim();
    const status = String(values[i][10]).trim();
    if (!booth || status === STATUS_CANCELLED) continue;
    const s = status === STATUS_CONFIRMED ? 'booked' : 'pending';
    if (out[booth] !== 'booked') out[booth] = s;
  }
  return out;
}

function doGet() {
  return json_({ ok: true, booths: statuses_() });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const d = JSON.parse(e.postData.contents || '{}');
    const booth = String(d.booth || '').trim().toUpperCase();
    const info = boothInfo(booth);
    if (!info) return json_({ ok: false, error: 'invalid_booth' });
    const brand = clean_(d.brand, 80), name = clean_(d.name, 80), phone = clean_(d.phone, 30), product = clean_(d.product, 80);
    if (!brand || !name || !phone || !product) return json_({ ok: false, error: 'missing_fields' });
    if (statuses_()[booth]) return json_({ ok: false, error: 'taken', booths: statuses_() });
    getSheet_().appendRow([new Date(), booth, info.tier, info.price, brand, name, phone,
                           clean_(d.instagram, 80), product, clean_(d.profile, 300), STATUS_PENDING, '']);
    return json_({ ok: true, booth: booth, tier: info.tier, price: info.price });
  } catch (err) {
    return json_({ ok: false, error: 'server_error' });
  } finally {
    lock.releaseLock();
  }
}

// Trim, cap length, and stop spreadsheet formula injection (=, +, -, @)
function clean_(v, max) {
  let s = String(v == null ? '' : v).replace(/[\r\n\t]+/g, ' ').trim().slice(0, max);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
