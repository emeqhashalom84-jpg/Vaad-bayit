// ============================================================
// Vaad Bayit — announcement email notifications (Google Apps Script)
//
// Bind this to the announcements RESPONSE SHEET:
// https://docs.google.com/spreadsheets/d/1DjYtYbEWhEnb9e5X7roZ-yFdb0R8NpAGAAE8VS7F4cM
// (fed by form https://docs.google.com/forms/d/1SVlzQsBNJNEGTZHMKKxJR01Pl3fAbeDFeSh8pj2Mqlk/edit)
//
// SETUP:
// 1. Open the sheet above -> Extensions -> Apps Script -> paste this file's content
//    (replace any default Code.gs content)
// 2. Project Settings -> Script Properties -> Add:
//      TELEGRAM_TOKEN = <bot token> (separate Apps Script project from the calls one, own copy)
//      TEST_EMAIL_1   = <first test address>
//      TEST_EMAIL_2   = <second test address>
// 3. Triggers (clock icon) -> Add Trigger:
//      Function: onFormSubmit | Event source: From spreadsheet | Event type: On form submit
//    (first save will ask you to authorize the script — approve it)
// 4. Test: submit the announcements form, confirm Telegram pings both admins regardless of
//    Active; confirm the two test addresses only receive an email when Active = כן.
// 5. When ready to go live: set TEST_MODE = false below (pulls all 14 tenant emails from the
//    contacts sheet instead of the test addresses) and re-save.
//
// RESPONSE SHEET COLUMNS (0-indexed) — email collection is ON for this form, so
// "Email Address" is an extra auto-collected column right after Timestamp:
//   0 Timestamp | 1 Email Address (auto) | 2 Date | 3 Title | 4 Response Content
//   5 Category | 6 Priority | 7 Active (כן/לא — only כן triggers an email)
//   8 תוקף עד (expiry date — new question, appended as the last column by Forms)
//
// EXTRA SETUP for auto-expiry:
//   Triggers -> Add Trigger -> Function: checkExpiredAnnouncements_
//   Event source: Time-driven | Type: Day timer | Time: any convenient window (e.g. 00:00-01:00)
// ============================================================

const CONTACTS_SHEET_ID = '1AttLipED7i-6iv7ZH6cjx8j32AnoSNQ2Kh9n-TTiJ8Q';
const ADMIN_EMAIL        = 'emeqhashalom84@gmail.com';
const ACTIVE_COL = 8;  // column H, 1-indexed
const EXPIRY_COL = 9;  // column I, 1-indexed — new תוקף עד question

// Testing phase: Oren only. Add Michael's Telegram chat id once he's set up (same as calls script).
const ADMIN_TELEGRAM_IDS = ['996999913'];

function prop_(key) {
  return PropertiesService.getScriptProperties().getProperty(key);
}

function sendTelegram_(chatId, text) {
  const token = prop_('TELEGRAM_TOKEN');
  if (!token) { Logger.log('Missing TELEGRAM_TOKEN script property'); return; }
  const url = 'https://api.telegram.org/bot' + token + '/sendMessage';
  UrlFetchApp.fetch(url, {
    method: 'post',
    payload: { chat_id: chatId, text: text },
    muteHttpExceptions: true
  });
}

function notifyAdminsTelegram_(text) {
  ADMIN_TELEGRAM_IDS.forEach(function (id) { sendTelegram_(id, text); });
}

// Link that opens the sheet directly at the Active cell for this row, so either
// admin can toggle it from their phone without hunting for the row.
function rowEditLink_(sheet, row) {
  const ss = sheet.getParent();
  return 'https://docs.google.com/spreadsheets/d/' + ss.getId() +
    '/edit#gid=' + sheet.getSheetId() + '&range=H' + row;
}

// TESTING PHASE — switch to false once all 14 tenant emails are confirmed in the contacts sheet.
// Test addresses come from Script Properties (TEST_EMAIL_1 / TEST_EMAIL_2), not hardcoded here.
const TEST_MODE = true;

function testEmails_() {
  return [prop_('TEST_EMAIL_1'), prop_('TEST_EMAIL_2')].filter(function (e) { return e; });
}

// Contacts sheet columns (0-indexed): 1=name, 2=building, 3=apartment, 4=email, 5=phone
function allTenantEmails_() {
  const ss = SpreadsheetApp.openById(CONTACTS_SHEET_ID);
  const sheet = ss.getSheets()[0];
  const rows = sheet.getDataRange().getValues();
  const emails = [];
  for (var i = 1; i < rows.length; i++) {
    var email = rows[i][4];
    if (email && String(email).trim()) emails.push(String(email).trim());
  }
  return emails;
}

function recipientEmails_() {
  return TEST_MODE ? testEmails_() : allTenantEmails_();
}

function onFormSubmit(e) {
  const sheet    = e.range.getSheet();
  const rowNum   = e.range.getRow();
  const row      = e.values;
  const title    = row[3];
  const content  = row[4];
  const category = row[5];
  const priority = row[6];
  const active   = row[7];

  const editLink = rowEditLink_(sheet, rowNum);
  notifyAdminsTelegram_('📢 הודעה חדשה הוגשה\n' +
    'כותרת: ' + title + '\nקטגוריה: ' + category + '\nעדיפות: ' + priority +
    '\nפעיל: ' + active + '\n\nלעריכה: ' + editLink);

  if (active !== 'כן') return; // draft/inactive announcement — no tenant email

  const recipients = recipientEmails_();
  if (!recipients.length) return;

  const subject = '📢 הודעה מהוועד: ' + title;
  const body = title + '\n\n' + content +
    '\n\nקטגוריה: ' + category + '\nעדיפות: ' + priority;

  MailApp.sendEmail({
    to: ADMIN_EMAIL,
    bcc: recipients.join(','),
    subject: subject,
    body: body
  });
}

// Run daily (time-driven trigger) — flips any expired כן row to לא פעיל.
function checkExpiredAnnouncements_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const rows = sheet.getDataRange().getValues();
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  for (var i = 1; i < rows.length; i++) {
    var active = rows[i][7];
    var expiry = rows[i][8];
    if (active !== 'כן' || !expiry) continue;

    var expiryDate = expiry instanceof Date ? new Date(expiry) : new Date(String(expiry));
    if (isNaN(expiryDate.getTime())) continue;
    expiryDate.setHours(0, 0, 0, 0);

    if (expiryDate < today) {
      sheet.getRange(i + 1, ACTIVE_COL).setValue('לא פעיל');
    }
  }
}
