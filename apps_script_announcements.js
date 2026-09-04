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
//      GITHUB_PAT     = <the token from D:\Claude_projects\Home_tech\git token.txt>
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
//   0 Timestamp | 1 Email Address (auto) | 2 Date (also used for future-scheduling — see
//     fetch_announcements_ci.py / generator's fetch_announcements(): a date later than
//     today hides the announcement until that day, even if Active=כן) | 3 Title
//   4 Response Content | 5 Category | 6 Priority | 7 Active (כן/לא — only כן triggers an email)
//   8 תוקף (ימים) — a Dropdown question (10/15/30/60/90 + "Other" for a custom number),
//     appended as the last column by Forms. Expiry = Timestamp (col 0) + this many days.
//     Blank = never auto-expires.
//
// EXTRA SETUP for auto-expiry:
//   Triggers -> Add Trigger -> Function: checkExpiredAnnouncements
//   Event source: Time-driven | Type: Day timer | Time: any convenient window (e.g. 00:00-01:00)
// ============================================================

// עדכון פרטים אישיים (Responses) — the single source of truth for tenant/contact data
// (see apps_script_admin.js's TENANT_FORM_SHEET_ID comment for the history of why this
// replaced an earlier, abandoned "master contacts" sheet).
const TENANT_FORM_SHEET_ID = '1dov_q0JSv74VMF30wQ177O9jVC1se21BW3AwibURHxs';
const ADMIN_EMAIL        = 'emeqhashalom84@gmail.com';
const ACTIVE_COL = 8;  // column H, 1-indexed
const VALID_DAYS_COL = 9;  // column I, 1-indexed — new תוקף (ימים) question
const REPO          = 'emeqhashalom84-jpg/Vaad-bayit';
const WORKFLOW_FILE = 'update-issues.yml';

function triggerDashboardRefresh_() {
  const pat = prop_('GITHUB_PAT');
  if (!pat) { Logger.log('Missing GITHUB_PAT script property'); return; }
  const url = 'https://api.github.com/repos/' + REPO + '/actions/workflows/' + WORKFLOW_FILE + '/dispatches';
  UrlFetchApp.fetch(url, {
    method: 'post',
    headers: { Authorization: 'token ' + pat, Accept: 'application/vnd.github+json' },
    contentType: 'application/json',
    payload: JSON.stringify({ ref: 'main' }),
    muteHttpExceptions: true
  });
}

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

// Header-text matching, not fixed position — Google Forms keeps a response sheet's
// columns in the order questions were ORIGINALLY created, not the form editor's current
// order, so a fixed-position mapping breaks silently. Minimal copy of the same helpers in
// apps_script_admin.js (duplicated — this is a separate standalone project).
function findCol_(headers, mustHave, mustNotHave) {
  mustNotHave = mustNotHave || [];
  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i] || '');
    var ok = mustHave.every(function (s) { return h.indexOf(s) !== -1; }) &&
             !mustNotHave.some(function (s) { return h.indexOf(s) !== -1; });
    if (ok) return i;
  }
  return -1;
}
function get_(row, map, key) {
  var i = map[key];
  return (i === -1 || i === undefined) ? '' : (row[i] || '');
}
// Collects both household contacts' emails, not just one per apartment. Rows marked
// 'בוטל' are skipped; when a household submitted more than once, only the LATEST
// remaining row for that address contributes its emails (avoids emailing a stale address).
function allTenantEmails_() {
  const sheet = SpreadsheetApp.openById(TENANT_FORM_SHEET_ID).getSheets()[0];
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const map = {
    building: findCol_(headers, ['בית']),
    apt:      findCol_(headers, ['מספר', 'דירה']),
    c1email:  findCol_(headers, ['קשר', '1', 'מייל']),
    c2email:  findCol_(headers, ['קשר', '2', 'מייל']),
    status:   findCol_(headers, ['סטטוס'])
  };
  const winners = {};
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    var b = String(get_(r, map, 'building')).trim(), a = String(get_(r, map, 'apt')).trim();
    if (!b && !a) continue;
    if (get_(r, map, 'status') === 'בוטל') continue;
    winners[b + '|' + a] = r; // later row always overwrites — latest wins
  }
  const emails = [];
  Object.keys(winners).forEach(function (addr) {
    var r = winners[addr];
    var e1 = get_(r, map, 'c1email'), e2 = get_(r, map, 'c2email');
    if (e1 && String(e1).trim()) emails.push(String(e1).trim());
    if (e2 && String(e2).trim()) emails.push(String(e2).trim());
  });
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

  triggerDashboardRefresh_();

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

// Parses the תאריך column, which can be a real Date (Form's Date-type answer) or
// text in either Y-M-D (pasted Excel) or M/D/Y (this sheet's own locale) format.
function parseScheduledDate_(v) {
  if (v instanceof Date) return v;
  if (!v) return null;
  var s = String(v).trim();
  var m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return new Date(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10));
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return new Date(parseInt(m[3], 10), parseInt(m[1], 10) - 1, parseInt(m[2], 10));
  return null;
}

// Run daily (time-driven trigger) — flips any expired כן row to לא פעיל.
// Expiry = (whichever is later: submission Timestamp, or a future-scheduled תאריך)
// + the number of days in VALID_DAYS_COL — so "10 days" always means 10 real days of
// visibility, even when the announcement was scheduled to publish after it was submitted.
// Accepts "Other" free-text entries too (e.g. "45"), not just the preset dropdown values.
function checkExpiredAnnouncements() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const rows = sheet.getDataRange().getValues();
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  var anyExpired = false;

  for (var i = 1; i < rows.length; i++) {
    var active    = rows[i][7];
    var validDays = parseInt(String(rows[i][8]).replace(/[^\d]/g, ''), 10);
    var timestamp = rows[i][0];
    if (active !== 'כן' || !validDays || !(timestamp instanceof Date)) continue;

    var scheduled = parseScheduledDate_(rows[i][2]);
    var anchor = (scheduled && scheduled > timestamp) ? scheduled : timestamp;

    var expiryDate = new Date(anchor);
    expiryDate.setDate(expiryDate.getDate() + validDays);
    expiryDate.setHours(0, 0, 0, 0);

    if (expiryDate < today) {
      sheet.getRange(i + 1, ACTIVE_COL).setValue('לא פעיל');
      anyExpired = true;
    }
  }

  if (anyExpired) triggerDashboardRefresh_();
}
