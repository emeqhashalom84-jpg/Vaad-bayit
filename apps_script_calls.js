// ============================================================
// Vaad Bayit — call notifications (Google Apps Script)
//
// Bind this to the calls/fault-report RESPONSE SHEET:
// https://docs.google.com/spreadsheets/d/176y6v-RfxaexwAUhguHY0AteJFjnfOdhUP0E-cNdEDc
//
// SETUP:
// 1. Open the sheet above -> Extensions -> Apps Script -> paste this file's content
//    (replace any default Code.gs content)
// 2. Project Settings (gear icon, left sidebar) -> Script Properties -> Add:
//      TELEGRAM_TOKEN = <the bot token from @BotFather>
//      GITHUB_PAT     = <the token from D:\Claude_projects\Home_tech\git token.txt>
// 3. Triggers (clock icon, left sidebar) -> Add Trigger, twice:
//      a) Function: onFormSubmit | Event source: From spreadsheet | Event type: On form submit
//      b) Function: onEdit       | Event source: From spreadsheet | Event type: On edit
//    (the first time you save a trigger, Google will ask you to authorize the script — approve it)
// 4. Test: submit the fault-report form using building 84 / apartment 2 (Oren's own contact
//    entry) so the lookup matches, then confirm email + Telegram arrive at Oren.
// 5. Test status change: edit column J (סטטוס) on the new row, confirm the reporter gets
//    emailed and Oren gets a Telegram ping, and the dashboard refreshes within ~1 min.
//
// COLUMNS on this response sheet (0-indexed, confirmed from the actual sheet):
//   0 Timestamp — when the call was opened | 1 שם מלא | 2 מספר בית | 3 מספר דירה
//   4 תיאור התקלה | 5 מיקום התקלה | 6 דחיפות | 7 העלאת תמונות/וידאו | 8 הערות
//   9 סטטוס (J, admin-set — triggers tenant notification when changed)
//  10 פעיל (K, admin-set — dashboard visibility toggle, NOT a notification trigger)
//  11 תאריך עדכון סטטוס (L, auto-filled by this script when J changes — for future SLA tracking)
// ============================================================

// עדכון פרטים אישיים (Responses) — the single source of truth for tenant/contact data
// (see apps_script_admin.js's TENANT_FORM_SHEET_ID comment for the history of why this
// replaced an earlier, abandoned "master contacts" sheet). This is a fallback lookup only
// — the fault-report Form now collects the reporter's own email directly (see onFormSubmit).
const TENANT_FORM_SHEET_ID = '1dov_q0JSv74VMF30wQ177O9jVC1se21BW3AwibURHxs';
const REPO              = 'emeqhashalom84-jpg/Vaad-bayit';
const WORKFLOW_FILE     = 'update-issues.yml';

// Testing phase: Oren only. Add Michael's Telegram chat id once he's set up.
const ADMIN_EMAIL         = 'emeqhashalom84@gmail.com';
const ADMIN_TELEGRAM_IDS  = ['996999913'];

const STATUS_COL = 10; // column J, 1-indexed — status text, drives tenant notification
const ACTIVE_COL = 11; // column K, 1-indexed — visibility toggle, defaults to פעיל on new calls
const STATUS_UPDATE_TS_COL = 12; // column L, 1-indexed — auto-filled with timestamp on status change
const REPORTER_EMAIL_COL = 13; // column M, 1-indexed — new "אימייל" Form question, appended last;
                                // the actual submitter's own email, so replies reach whoever
                                // opened the call regardless of which household contact it was

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

// Extracts a Google Drive file id from a share link like
// https://drive.google.com/open?id=XXXX or .../d/XXXX/view
function driveFileId_(url) {
  const m = url.match(/\/d\/([a-zA-Z0-9_-]+)|id=([a-zA-Z0-9_-]+)/);
  return m ? (m[1] || m[2]) : null;
}

// Sends the actual photo into the chat (not just a link). Telegram fetches
// the URL server-side, so the Drive file must be shared "Anyone with the
// link". Falls back to a plain text+link message if there's no image.
function sendTelegramPhoto_(chatId, photoUrl, caption) {
  const token = prop_('TELEGRAM_TOKEN');
  if (!token) { Logger.log('Missing TELEGRAM_TOKEN script property'); return; }
  const url = 'https://api.telegram.org/bot' + token + '/sendPhoto';
  UrlFetchApp.fetch(url, {
    method: 'post',
    payload: { chat_id: chatId, photo: photoUrl, caption: caption },
    muteHttpExceptions: true
  });
}

function notifyAdminsTelegram_(text) {
  ADMIN_TELEGRAM_IDS.forEach(function (id) { sendTelegram_(id, text); });
}

// Sends the first image as an actual photo (caption = the summary text);
// any additional images are listed as plain links in a follow-up message.
function notifyAdminsTelegramWithImages_(text, images) {
  if (!images || !images.length) { notifyAdminsTelegram_(text); return; }
  const fid = driveFileId_(images[0]);
  const photoUrl = fid ? ('https://drive.google.com/thumbnail?id=' + fid + '&sz=w1024') : null;
  ADMIN_TELEGRAM_IDS.forEach(function (id) {
    if (photoUrl) { sendTelegramPhoto_(id, photoUrl, text); }
    else { sendTelegram_(id, text); }
  });
  if (images.length > 1) {
    notifyAdminsTelegram_('תמונות נוספות:\n' + images.slice(1).join('\n'));
  }
}

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
function tenantColMap_(headers) {
  return {
    building: findCol_(headers, ['בית']),
    apt:      findCol_(headers, ['מספר', 'דירה']),
    lastName: findCol_(headers, ['משפחה']),
    c1name:   findCol_(headers, ['קשר', '1'], ['טלפון', 'מייל']),
    c1phone:  findCol_(headers, ['קשר', '1', 'טלפון']),
    c1email:  findCol_(headers, ['קשר', '1', 'מייל']),
    c2name:   findCol_(headers, ['קשר', '2'], ['טלפון', 'מייל']),
    c2phone:  findCol_(headers, ['קשר', '2', 'טלפון']),
    c2email:  findCol_(headers, ['קשר', '2', 'מייל']),
    rented:   findCol_(headers, ['שכורה']),
    ownerName:  findCol_(headers, ['בעל', 'שם']),
    ownerPhone: findCol_(headers, ['בעל', 'טלפון']),
    ownerEmail: findCol_(headers, ['בעל', 'מייל']),
    status:   findCol_(headers, ['סטטוס'])
  };
}
function get_(row, map, key) {
  var i = map[key];
  return (i === -1 || i === undefined) ? '' : (row[i] || '');
}
// Looks up a tenant card by building+apartment. Rows marked 'בוטל' are skipped; when a
// household submitted more than once, the LATEST remaining row for that address wins.
function lookupContact_(building, apt) {
  const sheet = SpreadsheetApp.openById(TENANT_FORM_SHEET_ID).getSheets()[0];
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const map = tenantColMap_(headers);
  const b = String(building).trim(), a = String(apt).trim();
  var found = null;
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    if (get_(r, map, 'status') === 'בוטל') continue;
    if (String(get_(r, map, 'building')).trim() === b && String(get_(r, map, 'apt')).trim() === a) {
      const rented = String(get_(r, map, 'rented')).trim() === 'כן';
      const c2name = get_(r, map, 'c2name'), ownerName = get_(r, map, 'ownerName');
      found = {
        lastName: get_(r, map, 'lastName'),
        contact1: { name: get_(r, map, 'c1name'), phone: get_(r, map, 'c1phone'), email: get_(r, map, 'c1email') },
        contact2: c2name ? { name: c2name, phone: get_(r, map, 'c2phone'), email: get_(r, map, 'c2email') } : null,
        rented: rented,
        owner: (rented && ownerName) ? { name: ownerName, phone: get_(r, map, 'ownerPhone'), email: get_(r, map, 'ownerEmail') } : null
      };
    }
  }
  return found;
}

// Builds a link that opens the sheet directly at the status cell (column J) for this row
function rowEditLink_(sheet, row) {
  const ss = sheet.getParent();
  return 'https://docs.google.com/spreadsheets/d/' + ss.getId() +
    '/edit#gid=' + sheet.getSheetId() + '&range=J' + row;
}

function onFormSubmit(e) {
  const row = e.values;
  const sheet = e.range.getSheet();
  const rowNum = e.range.getRow();
  const callNumber = rowNum - 1; // header is row 1

  const name     = row[1];
  const building = row[2];
  const apt      = row[3];
  const desc     = row[4];
  const location = row[5];
  const urgency  = row[6];
  const images   = row[7] ? String(row[7]).split(',').map(function(s){return s.trim();}).filter(function(s){return s;}) : [];

  sheet.getRange(rowNum, STATUS_COL).setValue('פתוח');
  sheet.getRange(rowNum, ACTIVE_COL).setValue('פעיל');

  // The Form's own email answer is the actual submitter — prefer it over a contacts
  // lookup, which can only find a household's contact1 and can't tell who really submitted.
  const formEmail = row[REPORTER_EMAIL_COL - 1];
  const contact = lookupContact_(building, apt);
  const reporterEmail = formEmail || (contact ? contact.contact1.email : null);
  const editLink = rowEditLink_(sheet, rowNum);

  const summary = 'קריאה #' + callNumber + '\n' +
    'שם: ' + name + '\n' +
    'בניין ' + building + ' דירה ' + apt + '\n' +
    'תיאור: ' + desc + '\n' +
    'מיקום: ' + location + '\n' +
    'דחיפות: ' + urgency;

  MailApp.sendEmail(ADMIN_EMAIL, 'קריאה חדשה #' + callNumber,
    summary + '\n\nלמענה ושינוי סטטוס: ' + editLink);

  if (reporterEmail) {
    MailApp.sendEmail(reporterEmail, 'אישור קבלת קריאה #' + callNumber,
      'התקבלה קריאתך, מספר קריאה: ' + callNumber + '\n\n' + summary);
  }

  notifyAdminsTelegramWithImages_('🔧 קריאה חדשה #' + callNumber + '\n' + summary +
    '\n\nלמענה ושינוי סטטוס: ' + editLink, images);

  triggerDashboardRefresh_();
}

function onEdit(e) {
  const range = e.range;
  const row = range.getRow();
  const col = range.getColumn();
  if (row === 1) return; // header row

  if (col === ACTIVE_COL) {
    // Visibility-only change (פעיל/לא פעיל) — just refresh the dashboard, no notification
    triggerDashboardRefresh_();
    return;
  }
  if (col !== STATUS_COL) return; // only react to column J (status) beyond this point

  const sheet = range.getSheet();
  const rowData = sheet.getRange(row, 1, 1, REPORTER_EMAIL_COL).getValues()[0];
  const callNumber = row - 1;
  const building  = rowData[2];
  const apt       = rowData[3];
  const newStatus = rowData[9];

  sheet.getRange(row, STATUS_UPDATE_TS_COL).setValue(new Date());

  const formEmail = rowData[REPORTER_EMAIL_COL - 1];
  const contact = lookupContact_(building, apt);
  const reporterEmail = formEmail || (contact ? contact.contact1.email : null);

  if (reporterEmail) {
    MailApp.sendEmail(reporterEmail, 'עדכון סטטוס קריאה #' + callNumber,
      'הסטטוס של קריאה #' + callNumber + ' עודכן ל: ' + newStatus);
  }

  notifyAdminsTelegram_('📋 סטטוס קריאה #' + callNumber + ' עודכן ל: ' + newStatus);

  triggerDashboardRefresh_();
}
