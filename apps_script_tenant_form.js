// ============================================================
// Vaad Bayit — personal-details (כרטיס דייר) form notifications (Google Apps Script)
//
// Bind this to the personal-details form's RESPONSE SHEET (Responses tab -> green
// Sheets icon -> "Create spreadsheet", if not already created).
//
// SETUP:
// 1. Open that response sheet -> Extensions -> Apps Script -> paste this file's content
// 2. Project Settings -> Script Properties -> Add:
//      TELEGRAM_TOKEN = <same bot token used in the other scripts>
// 3. Triggers -> Add Trigger -> Function: onFormSubmit | Event source: From spreadsheet
//    Event type: On form submit (first save will ask you to authorize — approve it)
// 4. Test: submit the personal-details form, confirm Telegram + email arrive at admin.
//
// IMPORTANT: columns are read by matching HEADER TEXT, not fixed position. Google Forms
// keeps a response sheet's columns in the order questions were ORIGINALLY created —
// reordering/restructuring questions in the editor later (e.g. adding the owner section)
// does NOT reorder existing columns, so a fixed-position mapping breaks silently. This
// script (and the matching admin-panel functions) instead look up each field by keyword,
// which stays correct no matter what order the columns actually ended up in.
// ============================================================

const ADMIN_EMAIL        = 'emeqhashalom84@gmail.com';
const ADMIN_TELEGRAM_IDS = ['996999913']; // add Michael's chat id here once he's set up

function prop_(key) {
  return PropertiesService.getScriptProperties().getProperty(key);
}

function sendTelegram_(chatId, text) {
  const token = prop_('TELEGRAM_TOKEN');
  if (!token) { Logger.log('Missing TELEGRAM_TOKEN script property'); return; }
  UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendMessage', {
    method: 'post', payload: { chat_id: chatId, text: text }, muteHttpExceptions: true
  });
}
function notifyAdminsTelegram_(text) {
  ADMIN_TELEGRAM_IDS.forEach(function (id) { sendTelegram_(id, text); });
}

// Finds the first header whose text contains every string in mustHave and none in
// mustNotHave. Returns -1 if not found.
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

function tenantFormColMap_(headers) {
  return {
    building:   findCol_(headers, ['בית']),
    apt:        findCol_(headers, ['מספר', 'דירה']),
    lastName:   findCol_(headers, ['משפחה']),
    c1name:     findCol_(headers, ['קשר', '1'], ['טלפון', 'מייל']),
    c1phone:    findCol_(headers, ['קשר', '1', 'טלפון']),
    c1email:    findCol_(headers, ['קשר', '1', 'מייל']),
    c2name:     findCol_(headers, ['קשר', '2'], ['טלפון', 'מייל']),
    c2phone:    findCol_(headers, ['קשר', '2', 'טלפון']),
    c2email:    findCol_(headers, ['קשר', '2', 'מייל']),
    rented:     findCol_(headers, ['שכורה']),
    ownerName:  findCol_(headers, ['בעל', 'שם']),
    ownerPhone: findCol_(headers, ['בעל', 'טלפון']),
    ownerEmail: findCol_(headers, ['בעל', 'מייל']),
    status:     findCol_(headers, ['סטטוס'])
  };
}
function get_(row, map, key) {
  var i = map[key];
  return (i === -1 || i === undefined) ? '' : (row[i] || '');
}

function onFormSubmit(e) {
  const sheet = e.range.getSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = tenantFormColMap_(headers);
  const row = e.values;

  const building = get_(row, map, 'building');
  const apt = get_(row, map, 'apt');
  const lastName = get_(row, map, 'lastName');
  const c1name = get_(row, map, 'c1name'), c1phone = get_(row, map, 'c1phone'), c1email = get_(row, map, 'c1email');
  const c2name = get_(row, map, 'c2name');
  const rented = get_(row, map, 'rented');
  const ownerName = get_(row, map, 'ownerName');

  const summary = 'עדכון כרטיס דייר חדש לאישור\n' +
    'בניין ' + building + ' דירה ' + apt + ' — ' + lastName + '\n' +
    'איש קשר 1: ' + c1name + ' (' + c1phone + ', ' + c1email + ')' +
    (c2name ? '\nאיש קשר 2: ' + c2name : '') +
    '\nדירה שכורה: ' + (rented || 'לא') +
    (rented === 'כן' && ownerName ? '\nבעל הדירה: ' + ownerName : '');

  MailApp.sendEmail(ADMIN_EMAIL, 'עדכון כרטיס דייר לאישור — בניין ' + building + ' דירה ' + apt,
    summary + '\n\nלאישור: פתח את אפליקציית הניהול, לשונית "כרטיסי דיירים"');

  notifyAdminsTelegram_('📇 ' + summary + '\n\nלאישור: אפליקציית הניהול, לשונית "כרטיסי דיירים"');
}
