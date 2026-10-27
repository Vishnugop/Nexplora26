/**
 * NEXPLORA '26 school registration endpoint.
 * 1. Create/open a Google Sheet.
 * 2. Extensions > Apps Script; replace the editor contents with this file.
 * 3. Deploy > New deployment > Web app.
 *    Execute as: Me
 *    Who has access: Anyone
 * 4. Copy the deployed /exec URL into index.html's scriptUrl field.
 */

const SHEET_NAME = 'Registrations';

function doGet() {
  return ContentService
    .createTextOutput('NEXPLORA 26 registration endpoint is running.')
    .setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const p = (e && e.parameter) ? e.parameter : {};

    // Basic spam trap: the front-end honeypot should be blank.
    if (p.website) {
      return json_({ ok: false, message: 'Invalid submission.' });
    }

    const schoolName = clean_(p.schoolName);
    const schoolAddress = clean_(p.schoolAddress);
    const coordinatorName = clean_(p.coordinatorName);
    const designation = clean_(p.designation);
    const email = clean_(p.email);
    const mobile = clean_(p.mobile);

    if (!schoolName || !schoolAddress || !coordinatorName ||
        !designation || !email || !mobile) {
      return json_({ ok: false, message: 'Please complete all required fields.' });
    }
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      return json_({ ok: false, message: 'Invalid email address.' });
    }

    const sheet = getSheet_();
    sheet.appendRow([
      new Date(),
      safeCell_(schoolName),
      safeCell_(schoolAddress),
      safeCell_(coordinatorName),
      safeCell_(designation),
      safeCell_(email),
      safeCell_(mobile)
    ]);

    return json_({ ok: true, message: 'Registration saved.' });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, message: 'Unable to save registration.' });
  } finally {
    lock.releaseLock();
  }
}

function getSheet_() {
  const props = PropertiesService.getScriptProperties();
  const spreadsheetId = props.getProperty('SPREADSHEET_ID');

  if (!spreadsheetId) {
    throw new Error('Set the SPREADSHEET_ID script property to your Google Sheet ID.');
  }

  const ss = SpreadsheetApp.openById(spreadsheetId);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'Timestamp',
      'School Name',
      'School Address',
      'Coordinator Name',
      'Coordinator Designation',
      'Email Address',
      'Mobile Number'
    ]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 7).setFontWeight('bold');
  }
  return sheet;
}

function clean_(value) {
  return String(value || '').trim().replace(/\s+/g, ' ').slice(0, 1000);
}

// Prevent values beginning with formula characters from being interpreted by Sheets.
function safeCell_(value) {
  const s = String(value || '');
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
