// ============================================================
// Vaad Bayit — Admin web app (Google Apps Script)
//
// A standalone (non-bound) project — NOT attached to any single sheet — so it can
// read/write the calls, announcements, contacts, and charges sheets all from one place.
// This is purely an ADDITIONAL way to manage things: the existing Forms and direct
// sheet editing keep working exactly as before, untouched.
//
// SETUP:
// 1. Go to script.google.com -> New project. Rename it "Vaad — Admin".
// 2. Code.gs: paste this file's content, replacing the default.
// 3. Click the "+" next to Files -> HTML -> name it exactly "Index" -> paste
//    admin_index.html's content into it.
// 4. Create the charges spreadsheet (see CHARGES SPREADSHEET section below), then
//    paste its ID into CHARGES_SHEET_ID below.
// 5. Project Settings (gear icon) -> Script Properties -> add:
//      TELEGRAM_TOKEN = <same bot token used in the calls/announcements scripts>
//      GITHUB_PAT     = <the token from D:\Claude_projects\Home_tech\git token.txt>
// 6. Deploy -> New deployment -> type "Web app":
//      Execute as: User accessing the web app   <-- IMPORTANT, see security note below
//      Who has access: Anyone with a Google account
//    Click Deploy, authorize, copy the web app URL.
// 7. Open that URL signed in as Oren or Michael. First visit asks you to authorize the
//    script under YOUR OWN account — approve it (this is normal, same as the trigger
//    authorization prompts you've already seen).
//
// SECURITY MODEL — no email allowlist needed, and this is important to keep it that way:
// Because deployment uses "Execute as: User accessing the web app", every read/write in
// this file runs with the PERMISSIONS OF WHOEVER OPENED THE PAGE. The calls/announcements/
// contacts sheets are shared only with Oren + Michael (Restricted access) — so anyone else
// opening the URL will simply get a permission error from Google when the page tries to
// read a sheet, with no code change needed here. Do NOT change "Execute as" to "Me" —
// that would run everything as the building's own account regardless of who's visiting,
// which would expose editing to anyone holding the URL.
//
// CHARGES SPREADSHEET (create this first, it doesn't exist yet):
// 1. Create a new Google Sheet, name it e.g. "Vaad — גביות נוספות".
// 2. Rename "Sheet1" to exactly: גביות   — headers in row 1:
//      charge_id | name | amount | date | active | description
// 3. Add a second sheet/tab named exactly: תשלומים   — headers in row 1:
//      charge_id | tenant_name | amount_paid | updated_at
// 4. Share it (Editor) with Oren's and Michael's own Google accounts, general access
//    Restricted — same pattern as the contacts/response sheets.
// 5. File -> Share -> Publish to web -> publish EACH tab separately as CSV (this is what
//    the local dashboard generator reads). Paste both resulting URLs into config.ini:
//      [google] charges_sheet_url = ...         (the "גביות" tab)
//               charge_payments_sheet_url = ...  (the "תשלומים" tab)
// ============================================================

const CALLS_SHEET_ID          = '176y6v-RfxaexwAUhguHY0AteJFjnfOdhUP0E-cNdEDc';
const ANNOUNCEMENTS_SHEET_ID  = '1DjYtYbEWhEnb9e5X7roZ-yFdb0R8NpAGAAE8VS7F4cM';
const CHARGES_SHEET_ID        = '1eQV7cCFwtNtnNmSBdew8kU8BFUVP8gVxa-YnbR7eQX0';
// עדכון פרטים אישיים (Responses) — the SINGLE source of truth for tenant/contact data.
// An earlier design used a separate "master contacts" sheet (1AttLipED7i...) that new
// submissions here got copied into after admin approval — that sheet turned out to be
// abandoned (Oren never actually worked in it), causing real bugs: duplicate-card checks
// silently found nothing, and payments/contact-card data went stale. Retired 2026-09-05 —
// this sheet is now read/written directly, no second sheet involved.
const TENANT_FORM_SHEET_ID    = '1dov_q0JSv74VMF30wQ177O9jVC1se21BW3AwibURHxs';
// כרטיסי ספקים (Suppliers) — 3 tabs: ספקים (one row per supplier), שדות מותאמים (flexible
// label/value pairs per supplier), תזכורות (recurring reminders per supplier). Linked by
// שם ספק (supplier name) rather than a separate id column — simple to read directly in the
// sheet; see updateSupplier's rename-cascade for why this stays safe across renames.
const SUPPLIERS_SHEET_ID      = '1QSLvulqQMwe9Uatw8YZCYIkyfQ2fXxGtdnhcyyFlQ7U';
const SUPPLIER_REMINDER_THRESHOLDS = [30, 15, 7, 3, 1];

const REPO          = 'emeqhashalom84-jpg/Vaad-bayit';
const WORKFLOW_FILE = 'update-issues.yml';
const ADMIN_EMAIL   = 'emeqhashalom84@gmail.com';
const ADMIN_TELEGRAM_IDS = ['996999913']; // add Michael's chat id here once he's set up

// Calls sheet columns (1-indexed, matches apps_script_calls.js)
const CALL_STATUS_COL = 10, CALL_ACTIVE_COL = 11, CALL_STATUS_TS_COL = 12;
// Announcements sheet columns (1-indexed, matches apps_script_announcements.js)
const ANN_ACTIVE_COL = 8, ANN_VALID_DAYS_COL = 9;

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Vaad Admin')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function prop_(key) {
  return PropertiesService.getScriptProperties().getProperty(key);
}

function sendTelegram_(chatId, text) {
  const token = prop_('TELEGRAM_TOKEN');
  if (!token) return;
  UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendMessage', {
    method: 'post', payload: { chat_id: chatId, text: text }, muteHttpExceptions: true
  });
}
function notifyAdminsTelegram_(text) {
  ADMIN_TELEGRAM_IDS.forEach(function (id) { sendTelegram_(id, text); });
}
function triggerDashboardRefresh_() {
  const pat = prop_('GITHUB_PAT');
  if (!pat) return;
  UrlFetchApp.fetch('https://api.github.com/repos/' + REPO + '/actions/workflows/' + WORKFLOW_FILE + '/dispatches', {
    method: 'post',
    headers: { Authorization: 'token ' + pat, Accept: 'application/vnd.github+json' },
    contentType: 'application/json',
    payload: JSON.stringify({ ref: 'main' }),
    muteHttpExceptions: true
  });
}

// Manual "refresh now" button on the dashboard tab — same trigger as every create/update
// action already fires, exposed on demand for testing or after a direct Sheet edit that
// bypassed this app entirely. Unlike triggerDashboardRefresh_, this surfaces failures to
// the admin instead of failing silently, since it's the whole point of the button.
function manualRefreshAll() {
  const pat = prop_('GITHUB_PAT');
  if (!pat) throw new Error('חסר GITHUB_PAT ב-Script Properties');
  const resp = UrlFetchApp.fetch('https://api.github.com/repos/' + REPO + '/actions/workflows/' + WORKFLOW_FILE + '/dispatches', {
    method: 'post',
    headers: { Authorization: 'token ' + pat, Accept: 'application/vnd.github+json' },
    contentType: 'application/json',
    payload: JSON.stringify({ ref: 'main' }),
    muteHttpExceptions: true
  });
  const code = resp.getResponseCode();
  if (code !== 204) throw new Error('GitHub API החזיר קוד ' + code);
  return true;
}

// ── Tenant Cards (כרטיס דייר) ─────────────────────────────────────────────
// Single source of truth: TENANT_FORM_SHEET_ID (עדכון פרטים אישיים responses) — see the
// constant's comment above for why. Columns are read by matching HEADER TEXT, not fixed
// position — Google Forms keeps a response sheet's columns in the order questions were
// ORIGINALLY created, which does not follow later reordering in the form editor, so a
// fixed-position mapping breaks silently. Every row is a submission (Form-driven or
// admin-created); a household that updates its details more than once ends up as multiple
// rows for the same building+apt — getTenantCards() below collapses those to the latest
// non-dismissed one automatically, so nothing needs to be deleted for correct behavior.
// סטטוס is admin-only, never written by the Form: blank = not yet reviewed, 'אושר' =
// reviewed, 'בוטל' = dismissed (excluded from results entirely, e.g. a duplicate/test
// submission). Auto-created in the next free column on first use if missing.
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
function set_(sheet, rowNum, map, key, value) {
  var i = map[key];
  if (i === -1 || i === undefined) return;
  sheet.getRange(rowNum, i + 1).setValue(value);
}
function tenantSheet_() {
  if (!TENANT_FORM_SHEET_ID || TENANT_FORM_SHEET_ID.indexOf('PASTE') !== -1) {
    throw new Error('TENANT_FORM_SHEET_ID לא הוגדר עדיין ב-Code.gs');
  }
  return SpreadsheetApp.openById(TENANT_FORM_SHEET_ID).getSheets()[0];
}
// Returns [sheet, headers, map], creating a סטטוס column if none exists yet.
function tenantCtx_() {
  const sheet = tenantSheet_();
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var map = tenantColMap_(headers);
  if (map.status === -1) {
    const col = headers.length + 1;
    sheet.getRange(1, col).setValue('סטטוס');
    headers = sheet.getRange(1, 1, 1, col).getValues()[0];
    map = tenantColMap_(headers);
  }
  return [sheet, headers, map];
}
// Standard household display name: contact1 [+ "ו"+contact2] + lastName, e.g.
// "אורן ואורלי אלקיים" (or just "אורן אלקיים" when there's no contact2).
function buildDisplayName_(c1name, c2name, lastName) {
  var head = String(c1name || '').trim();
  var c2 = String(c2name || '').trim();
  if (c2) head = head ? (head + ' ו' + c2) : ('ו' + c2);
  var last = String(lastName || '').trim();
  return head ? (last ? head + ' ' + last : head) : last;
}
function rowToTenantCard_(r, rowNum, map) {
  const rented = String(get_(r, map, 'rented')).trim() === 'כן';
  const lastName = get_(r, map, 'lastName'), c1name = get_(r, map, 'c1name'), c2name = get_(r, map, 'c2name');
  const ownerName = get_(r, map, 'ownerName');
  return {
    rowNum: rowNum,
    building: get_(r, map, 'building'), apt: get_(r, map, 'apt'), lastName: lastName,
    displayName: buildDisplayName_(c1name, c2name, lastName),
    contact1: { name: c1name, phone: get_(r, map, 'c1phone'), email: get_(r, map, 'c1email') },
    contact2: c2name ? { name: c2name, phone: get_(r, map, 'c2phone'), email: get_(r, map, 'c2email') } : null,
    rented: rented,
    owner: (rented && ownerName) ? { name: ownerName, phone: get_(r, map, 'ownerPhone'), email: get_(r, map, 'ownerEmail') } : null,
    status: get_(r, map, 'status')
  };
}
// Collapses every row to the LATEST non-dismissed submission per building+apt — a
// household that updated its details more than once (multiple Form submissions) never
// shows as duplicate cards. Rows marked 'בוטל' are excluded entirely.
function getTenantCards() {
  const ctx = tenantCtx_(), sheet = ctx[0], map = ctx[2];
  const rows = sheet.getDataRange().getValues();
  const winners = {}, order = [];
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    var b = String(get_(r, map, 'building')).trim(), a = String(get_(r, map, 'apt')).trim();
    if (!b && !a) continue; // blank row
    if (get_(r, map, 'status') === 'בוטל') continue; // dismissed — never a live card
    var addr = b + '|' + a;
    if (!winners[addr]) order.push(addr);
    winners[addr] = { row: r, rowNum: i + 1 }; // later row always overwrites — latest wins
  }
  // Display order: by building then apartment number, not submission order.
  order.sort(function (x, y) {
    var xb = x.split('|'), yb = y.split('|');
    var bDiff = (Number(xb[0]) || 0) - (Number(yb[0]) || 0);
    return bDiff !== 0 ? bDiff : (Number(xb[1]) || 0) - (Number(yb[1]) || 0);
  });
  return order.map(function (addr) {
    var w = winners[addr];
    return rowToTenantCard_(w.row, w.rowNum, map);
  });
}
function lookupContactByAddr_(building, apt) {
  const b = String(building).trim(), a = String(apt).trim();
  return getTenantCards().find(function (t) { return String(t.building).trim() === b && String(t.apt).trim() === a; }) || null;
}
// Used by the charges tab's per-tenant payment list — keyed by the full household display
// name (contact1 [+ ו+contact2] + lastName, e.g. "אורן ואורלי אלקיים"), which token-matches
// Excel's full household name (see _name_match in the generator).
function getTenantNames() {
  return getTenantCards().map(function (t) { return t.displayName; }).filter(function (n) { return n; });
}
/* ───────────────────────── BANK STATEMENT PDF IMPORT ───────────────────────── */
// "Vaad — כספים" — a different spreadsheet from CHARGES_SHEET_ID (that's "Vaad — גביות נוספות").
const FINANCE_SHEET_ID = '1HC7znYWIVfMA-NZSJA-0braStI0NQa0vZiYymVo-pWU';
function bankSheet_() { return SpreadsheetApp.openById(FINANCE_SHEET_ID).getSheetByName('תנועות בנק'); }

// Existing תנועות בנק rows' date column may hold a real Date object (Sheets auto-converts a
// typed/pasted "DD/MM/YYYY" string the moment it looks date-like, unless explicitly locked as
// text) rather than the plain string this comparison needs — normalizing both shapes to the
// same "DD/MM/YYYY" text is what makes date-based duplicate matching actually work.
function normalizeDateStr_(v) {
  if (v instanceof Date) {
    var d = v.getDate(), m = v.getMonth() + 1, y = v.getFullYear();
    return (d < 10 ? '0' : '') + d + '/' + (m < 10 ? '0' : '') + m + '/' + y;
  }
  return String(v || '').trim();
}

// The sheet is newest-first, so row 2 (right after the header) is always the latest recorded
// transaction — shown in the import tab as a reference point ("here's where the sheet is
// already synced up to; anything after this date is what's actually new").
function getLastBankRow() {
  const rows = bankSheet_().getDataRange().getValues();
  if (rows.length < 2) return null;
  const r = rows[1];
  return { date: normalizeDateStr_(r[0]), name: r[1] || '', action: r[2] || '', balance: r[5] || '' };
}

function getSupplierNames() {
  return getSuppliers().map(function (s) { return s.name; }).filter(function (n) { return n; });
}

function dayOfMonth_(dateStr) {
  var m = /^(\d{2})\//.exec(dateStr || '');
  return m ? parseInt(m[1], 10) : null;
}

// Groups by the exact פעולה text — works for BOTH sides, not just tenant credits: a debit's
// action text ("בזק-הוראות קבע", "קונה בע\"מ", "חברת החשמל ליש", generic "שיק") is just as
// reliable a grouping key as a credit's ("זיכוי מלאומי"), so one shared function recommends a
// name for expenses (→ supplier) and income (→ tenant) alike. Within a group, disambiguates
// two+ names sharing the same action text (e.g. two tenants both paying via "מלאומי", or two
// suppliers both paid by "שיק") first by matching amount (many recurring lines — בזק=49.14,
// ביטוח=3440, חוב לאילן=4200 — have a near-fixed amount), then by day-of-month (real example
// already in this sheet: מיכאל ומריה בן אורי pays "מלאומי" around the 12th, עמרי ושיר יוסוב
// around the 7th-8th — same action text, different day), else falls back to plain majority.
// Returns {actionText: [{day, amount, name, count}]}.
function getBankNameRecommendations_() {
  const rows = bankSheet_().getDataRange().getValues();
  const counts = {}; // action -> "day|amount" -> name -> count
  for (var i = 1; i < rows.length; i++) {
    const action = String(rows[i][2] || '').trim();
    const name = String(rows[i][1] || '').trim();
    const amount = Number(rows[i][3]) || Number(rows[i][4]) || 0;
    const day = dayOfMonth_(normalizeDateStr_(rows[i][0]));
    if (!action || !name || day === null) continue;
    const bucket = day + '|' + amount.toFixed(2);
    counts[action] = counts[action] || {};
    counts[action][bucket] = counts[action][bucket] || {};
    counts[action][bucket][name] = (counts[action][bucket][name] || 0) + 1;
  }
  const rec = {};
  Object.keys(counts).forEach(function (action) {
    rec[action] = [];
    Object.keys(counts[action]).forEach(function (bucket) {
      var parts = bucket.split('|'), day = Number(parts[0]), amount = Number(parts[1]);
      var best = '', bestCount = 0;
      Object.keys(counts[action][bucket]).forEach(function (name) {
        if (counts[action][bucket][name] > bestCount) { best = name; bestCount = counts[action][bucket][name]; }
      });
      rec[action].push({ day: day, amount: amount, name: best, count: bestCount });
    });
  });
  return rec;
}

// Picks the best-matching learned entry for this action's history: exact amount match first
// (tightest signal — most recurring lines have a near-fixed amount), then closest day-of-month
// (within 3 days), else the single most-seen name overall for that action text.
function recommendBankName_(actionRecs, day, amount) {
  if (!actionRecs || !actionRecs.length) return '';

  // Amount is only a USEFUL signal when it uniquely identifies one name — with several
  // tenants all paying the same ₪210 standard rate, "amount matches" is true for all of
  // them and picking the first match ignores day-of-month entirely (real bug: a day-7
  // transaction recommended מיכאל instead of עמרי ושיר יוסוב, because both pay ₪210 and
  // מיכאל's entry happened to come first). So: only trust an amount match if every entry
  // sharing that amount also shares the same name.
  var byAmount = actionRecs.filter(function (e) { return Math.abs(e.amount - amount) < 0.01; });
  if (byAmount.length) {
    var namesByAmount = {};
    byAmount.forEach(function (e) { namesByAmount[e.name] = true; });
    if (Object.keys(namesByAmount).length === 1) return byAmount[0].name;
    // Ambiguous by amount alone (e.g. several tenants on the same rate) — use day-of-month
    // to pick the closest entry among just this amount-matching subset.
    if (day !== null) {
      var closest = byAmount.slice().sort(function (a, b) { return Math.abs(a.day - day) - Math.abs(b.day - day); })[0];
      if (Math.abs(closest.day - day) <= 3) return closest.name;
    }
  }

  if (day !== null) {
    var nearest = null, nearestDist = 4; // tolerance: within 3 days
    actionRecs.forEach(function (e) {
      var dist = Math.abs(e.day - day);
      if (dist < nearestDist) { nearestDist = dist; nearest = e; }
    });
    if (nearest) return nearest.name;
  }
  var overall = actionRecs.slice().sort(function (a, b) { return b.count - a.count; })[0];
  return overall ? overall.name : '';
}

// Converts an uploaded bank-statement PDF (base64) to a temp Google Doc via Drive's OCR
// conversion, extracts the text, then deletes the temp Doc. OCR (not client-side PDF text
// extraction) because it handles Hebrew RTL correctly — a real, common failure point for
// browser PDF-text libraries on Hebrew documents. Requires the "Drive API" Advanced Google
// Service enabled in this project (Apps Script editor → Services → + → Drive API).
function parseBankStatementPdf(base64Data, fileName) {
  const bytes = Utilities.base64Decode(base64Data);
  const blob = Utilities.newBlob(bytes, 'application/pdf', fileName || 'statement.pdf');
  const resource = { name: 'TEMP_bank_ocr_' + Date.now(), mimeType: 'application/vnd.google-apps.document' };
  const file = Drive.Files.create(resource, blob, { fields: 'id' });
  var text = '';
  try {
    text = DocumentApp.openById(file.id).getBody().getText();
  } finally {
    Drive.Files.remove(file.id);
  }
  return parseAndAnnotateStatement_(text);
}

// Parses OCR'd statement text into transaction rows: date, פעולה text, amount, running
// balance. Credit/debit isn't read from a column (OCR text loses the PDF's visual column
// layout) — it's inferred by comparing each row's balance to the next (older) row's balance,
// since rows read newest-first, matching this sheet's own convention. The oldest row on a
// given upload has no older balance to diff against, so it falls back to a keyword guess —
// Oren reviews and can flip it before saving either way.
// NOTE: first-cut regex — the exact spacing/wording Drive's OCR produces hasn't been verified
// against a real conversion yet; expect to adjust after the first real test upload.
function parseAndAnnotateStatement_(text) {
  const lineRe = /(\d{2}\/\d{2}\/\d{4})\s+([^\d₪\r\n]+?)\s+([\d,]+\.\d{2})\s*₪\s*([\d,]+\.\d{2})/g;
  const parsed = [];
  var m;
  while ((m = lineRe.exec(text)) !== null) {
    parsed.push({
      date: m[1],
      action: m[2].trim(),
      amount: parseFloat(m[3].replace(/,/g, '')),
      balance: parseFloat(m[4].replace(/,/g, '')),
    });
  }
  for (var i = 0; i < parsed.length; i++) {
    var older = parsed[i + 1];
    if (older) {
      parsed[i].isCredit = (parsed[i].balance - older.balance) >= 0;
    } else {
      parsed[i].isCredit = /זיכוי|העברה|הפקדה/.test(parsed[i].action);
    }
  }

  // Keeps the existing row's own name+עבור text (not just a true/false flag) so a duplicate
  // row in the review table can show what's already recorded for it, per Oren (2026-09-10) —
  // useful context even though these rows aren't re-saved.
  const existing = bankSheet_().getDataRange().getValues();
  const existingByKey = {};
  for (var j = 1; j < existing.length; j++) {
    var amt = Number(existing[j][3]) || Number(existing[j][4]) || 0;
    existingByKey[normalizeDateStr_(existing[j][0]) + '|' + amt.toFixed(2)] = { name: existing[j][1] || '', note: existing[j][6] || '' };
  }

  const rec = getBankNameRecommendations_();
  // Every field explicitly coerced to a plain primitive (String/Number/Boolean) before it
  // crosses the client-server boundary — google.script.run can silently deliver `null` to
  // the client if anything in the returned structure doesn't serialize cleanly (NaN, undefined,
  // etc.), and the resulting "Cannot read properties of null" on the client gives no hint that
  // THIS is why. Real bug hit 2026-09-10 — Executions showed the function completing
  // successfully every time, only the client ever saw a broken response.
  return parsed.map(function (p) {
    const existingRow = existingByKey[p.date + '|' + p.amount.toFixed(2)];
    return {
      date: String(p.date || ''),
      action: String(p.action || ''),
      amount: Number(p.amount) || 0,
      balance: Number(p.balance) || 0,
      isCredit: !!p.isCredit,
      isDuplicate: !!existingRow,
      existingName: String((existingRow && existingRow.name) || ''),
      existingNote: String((existingRow && existingRow.note) || ''),
      recommendedName: String(recommendBankName_(rec[p.action], dayOfMonth_(p.date), p.amount) || ''),
    };
  });
}

// Inserts new bank rows at the TOP (right after the header), matching this sheet's
// newest-first convention. `rows` (already newest-first, as reviewed/edited by Oren):
// [{date, name, action, amount, isCredit, balance}].
function appendBankTransactions(rows) {
  if (!rows || !rows.length) return true;
  const sheet = bankSheet_();
  sheet.insertRowsAfter(1, rows.length);
  const matrix = rows.map(function (r) {
    return [
      r.date, r.name || '', r.action,
      r.isCredit ? '' : r.amount,
      r.isCredit ? r.amount : '',
      '₪' + Number(r.balance).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }),
      r.note || '',
    ];
  });
  const range = sheet.getRange(2, 1, rows.length, 7);
  range.setValues(matrix);
  // lockDateAsText_ forces plain-text format so Sheets can't silently re-parse/scramble the
  // date (same fix used elsewhere in this project) — but plain text defaults to LEFT alignment,
  // while every other date in this sheet is right-aligned (a real number/date's default).
  // Explicitly re-align right so the new rows visually match the rest of the column.
  for (var i = 0; i < rows.length; i++) {
    lockDateAsText_(sheet, 2 + i, 1, rows[i].date);
    sheet.getRange(2 + i, 1).setHorizontalAlignment('right');
  }

  // Per Oren's decision (2026-09-09): distribution runs automatically, but only for rows he
  // left checked (r.distribute) — an "אחר"/settlement credit tagged with a tenant's name but
  // NOT a real monthly payment (e.g. an old-debt/2025 settlement) can be excluded per-row.
  // 2026-09-10: a row can also split part of its amount toward the 2025 carryover debt
  // (r.carryoverAmount) — that part reduces חוב מועבר directly instead of being distributed.
  rows.forEach(function (r) {
    if (r.isCredit && r.name && r.distribute) {
      var amt = (r.distributeAmount === undefined || r.distributeAmount === null) ? r.amount : r.distributeAmount;
      distributeTenantPayment_(r.name, amt);
    }
    if (r.isCredit && r.name && r.carryoverAmount) {
      reduceCarryoverDebt_(r.name, r.carryoverAmount);
    }
    // 2026-09-10: same "automatic update, per-row opt-out" pattern as tenant-payment
    // distribution above, applied to debits — appends one row to הוצאות per tagged supplier.
    if (!r.isCredit && r.name && r.updateExpenses) {
      appendExpenseRow_(r.name, r.amount, r.date, r.note, r.expenseType);
    }
  });

  triggerDashboardRefresh_();
  return true;
}

function tenantPaymentsSheet_() { return SpreadsheetApp.openById(FINANCE_SHEET_ID).getSheetByName('תקבולי דיירים'); }
function settingsSheet_() { return SpreadsheetApp.openById(FINANCE_SHEET_ID).getSheetByName('הגדרות'); }
function approvedPaymentsSheet_() { return SpreadsheetApp.openById(FINANCE_SHEET_ID).getSheetByName('תשלומים מאושרים'); }
function expensesSheet_() { return SpreadsheetApp.openById(FINANCE_SHEET_ID).getSheetByName('הוצאות'); }

const MONTHS_HE_A_ = ['ינואר', 'פברואר', 'מרץ', 'אפריל', 'מאי', 'יוני', 'יולי', 'אוגוסט', 'ספטמבר', 'אוקטובר', 'נובמבר', 'דצמבר'];

// Same keyword list as BUDGET_2026 in vaad_bayit_generator.py (kept in sync manually — this is
// the Apps Script side, which has no access to the Python module). Per Oren (2026-09-10): a
// bank debit's סוג is "צפויה" only if it's part of the approved recurring budget, else "לא צפויה".
const BUDGET_EXPENSE_KEYWORDS_ = ['בדיקת מעליות', 'קונה', 'מעיינות', 'מונה', 'ביטוח', 'ניקיון', 'גג', 'בזק', 'עמלות בנק'];
function isExpenseBudgeted_(supplierName) {
  return BUDGET_EXPENSE_KEYWORDS_.some(function (kw) { return String(supplierName || '').indexOf(kw) !== -1; });
}

// Appends one row to "הוצאות" for a tagged bank debit — per Oren's decision (2026-09-10), unlike
// תקבולי דיירים this sheet is one row per expense event, so there's no month-spread logic here:
// קטגוריה defaults to the supplier name chosen in the bank row, and מיון לפי חודשים mirrors the
// existing rows' own convention (numeric month index, matching MONTHS_HE_A_). סוג is
// auto-classified against the recurring-budget keyword list above by default, but Oren can
// force it either way per row (not every real expense fits that list, or the supplier may not
// exist yet in כרטיסי ספקים) — explicitType, if passed, wins over the automatic guess.
// Inserted at row 2 (right after the header), not appended at the physical bottom — same
// "newest first" convention as תנועות בנק (per Oren, 2026-09-10: appendRow() landed new rows
// at the very end regardless of month, breaking the sheet's existing newest-first layout).
function appendExpenseRow_(supplierName, amount, dateStr, note, explicitType) {
  const parts = String(dateStr).split('/'); // DD/MM/YYYY
  const month = Number(parts[1]);
  const year = Number(parts[2]);
  const monthName = MONTHS_HE_A_[month - 1] || '';
  const type = explicitType || (isExpenseBudgeted_(supplierName) ? 'צפויה' : 'לא צפויה');
  const sheet = expensesSheet_();
  sheet.insertRowAfter(1);
  sheet.getRange(2, 1, 1, 8).setValues([[supplierName, year, monthName, amount, type, '', note || '', month]]);
}

// Real bug caught by Oren (2026-09-10): the distribution/normalize math was treating every
// month as the flat base/corner rate, so a tenant with a committee-approved reduced amount for
// a SPECIFIC month (עמרי ושיר יוסוף = ₪90 for ינואר, ענבל הוכבלד = ₪90 for אוגוסט — both real
// rows in "תשלומים מאושרים") had that ₪90 silently shifted to whatever month the redistribution
// happened to land the leftover on instead, and the row's total came out wrong too (a real ₪90
// exception doesn't divide evenly into ₪210 months, so ignoring it breaks the whole spread).
// Returns {monthIndex(0-11): amount} for the given tenant, fuzzy-matched by name.
function getApprovedPaymentsFor_(tenantName) {
  const rows = approvedPaymentsSheet_().getDataRange().getValues();
  const map = {};
  for (var i = 1; i < rows.length; i++) {
    if (!fuzzyNameMatch_(tenantName, String(rows[i][0] || ''))) continue;
    var monthIdx = MONTHS_HE_A_.indexOf(String(rows[i][1] || '').trim());
    if (monthIdx === -1) continue;
    map[monthIdx] = Number(rows[i][2]) || 0;
  }
  return map;
}

function getSetting_(key) {
  const rows = settingsSheet_().getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0] || '').trim() === key) return rows[i][1];
  }
  return null;
}

// Server-side counterpart of admin_index.html's fuzzyNameMatch_ — same idea (the sheet's own
// history and official display names are independently maintained, so match by shared word,
// not exact string), needed here to find the right תקבולי דיירים row for a tagged tenant name.
function fuzzyNameMatch_(a, b) {
  function toks(s) {
    return String(s || '').split(/\s+/).map(function (w) {
      return (w[0] === 'ו' && w.length > 1) ? w.slice(1) : w;
    }).filter(Boolean);
  }
  var ta = toks(a), tb = toks(b);
  return ta.some(function (w) { return tb.indexOf(w) !== -1; });
}

// Distributes a tenant's new payment across their תקבולי דיירים row: recomputes the FULL
// Jan-Dec spread from the new cumulative total (existing months' sum + this payment), filling
// January first, each month capped at the tenant's rate, spilling any remainder into the next
// month — NOT "put it in the month it was actually paid". Per Oren (2026-09-09): approved-rate
// overrides are past cases, ignored here; 2025 carryover debt is tracked entirely separately
// and never touched by this. Only writes columns D:O (the 12 month cells) — never P/Q/R
// (סה"כ שולם בפועל / יתרת חוב חודשית / יתרת חוב שנתית), which are live formulas in this sheet
// and recompute on their own once D:O change.
// Shared by distributeTenantPayment_ (adds a new payment) and the bulk normalize tool below
// (just re-spreads the EXISTING total, extraAmount=0) — same fill-forward rule either way.
function _redistributeRow(row, extraAmount, approvedMap) {
  const isCorner = String(row[19] || '').trim() === 'פינתי'; // עמודה T = סוג דירה
  const rate = Number(getSetting_(isCorner ? 'תעריף_פינה' : 'תעריף_רגיל')) || (isCorner ? 170 : 210);
  var currentTotal = 0;
  for (var m = 3; m <= 14; m++) currentTotal += Number(row[m]) || 0; // עמודות D:O
  var remaining = currentTotal + Number(extraAmount || 0);
  const newMonthly = [];
  for (var mo = 0; mo < 12; mo++) {
    // A committee-approved month (תשלומים מאושרים) is capped at ITS approved amount instead
    // of the flat rate — e.g. עמרי's ינואר is capped at ₪90, not ₪210 — exactly mirroring the
    // Python generator's own debt formula (`_appr.get(i, base_rate)`), just inverted (spreading
    // a payment forward instead of computing debt from it).
    var monthCap = (approvedMap && approvedMap[mo] !== undefined) ? approvedMap[mo] : rate;
    var fill = Math.max(0, Math.min(monthCap, remaining));
    // Rounded to whole ₪ for display — per Oren (2026-09-10), no need to track sub-shekel
    // remainders (e.g. ₪0.80). remaining still subtracts the exact fill so rounding never
    // compounds into a real accounting drift, just keeps each cell's own value shekel-clean.
    newMonthly.push(fill ? Math.round(fill) : '');
    remaining -= fill;
  }
  return newMonthly;
}

function distributeTenantPayment_(tenantName, amount) {
  const sheet = tenantPaymentsSheet_();
  const rows = sheet.getDataRange().getValues();
  var rowIdx = -1;
  for (var i = 1; i < rows.length; i++) {
    if (fuzzyNameMatch_(tenantName, String(rows[i][0] || ''))) { rowIdx = i; break; }
  }
  if (rowIdx === -1) {
    log_('distributeTenantPayment_: no תקבולי דיירים row matched "' + tenantName + '" — skipped');
    return false;
  }
  const newMonthly = _redistributeRow(rows[rowIdx], amount, getApprovedPaymentsFor_(tenantName));
  sheet.getRange(rowIdx + 1, 4, 1, 12).setValues([newMonthly]); // D:O, 1-indexed row/col
  return true;
}

// Bulk one-time cleanup tool (Oren, 2026-09-10): historic rows entered month-by-month (before
// the distribution logic existed) can show a month as "partial" even though a LATER month's
// overpayment already covers it (real example: אורן ואורלי אלקיים — April ₪126.8 partial,
// May ₪294 makes up the difference, but each cell was recorded independently so April still
// looks unpaid). Re-spreading every tenant's EXISTING total via the same fill-forward rule
// makes the whole sheet consistent with how new bank-import payments are now handled — no new
// money added (extraAmount=0), just re-arranged. Preview first, apply only the rows confirmed.
function previewNormalizeAllTenants() {
  const rows = tenantPaymentsSheet_().getDataRange().getValues();
  const preview = [];
  for (var i = 1; i < rows.length; i++) {
    const name = String(rows[i][0] || '').trim();
    // Skip rows with no real name, AND the "אחר" catch-all row (one-time/misc income not
    // tied to a tenant's monthly rate, e.g. ₪604.84+₪2,370 lump sums — has no building/
    // apartment, unlike every real tenant row). Redistributing it via monthly-rate logic
    // would be meaningless. Building (col B) is a more robust check than the literal
    // string "אחר" — catches any future non-tenant summary row the same way.
    if (!name || !String(rows[i][1] || '').trim()) continue;
    const oldMonthly = rows[i].slice(3, 15).map(function (v) { return Number(v) || 0; });
    const newMonthly = _redistributeRow(rows[i], 0, getApprovedPaymentsFor_(name)).map(function (v) { return Number(v) || 0; });
    const changed = oldMonthly.some(function (v, idx) { return Math.abs(v - newMonthly[idx]) > 0.001; });
    if (changed) preview.push({ rowIdx: i, name: name, oldMonthly: oldMonthly, newMonthly: newMonthly });
  }
  return preview;
}

function applyNormalizeAllTenants(rowIdxs) {
  const sheet = tenantPaymentsSheet_();
  const rows = sheet.getDataRange().getValues();
  rowIdxs.forEach(function (i) {
    const name = String(rows[i][0] || '').trim();
    const newMonthly = _redistributeRow(rows[i], 0, getApprovedPaymentsFor_(name));
    sheet.getRange(i + 1, 4, 1, 12).setValues([newMonthly]);
  });
  triggerDashboardRefresh_();
  return true;
}

function carryoverSheet_() { return SpreadsheetApp.openById(FINANCE_SHEET_ID).getSheetByName('חוב מועבר'); }

// Reduces a tenant's 2025 carryover debt by `amount` (e.g. 1520 - 520 = 1000), per Oren's
// request (2026-09-10) to automate this instead of a manual reminder. Floors at 0 — never
// goes negative. If the tenant has no row in חוב מועבר at all, there's nothing to reduce
// (logged, not an error — Oren decides via the split popup, a mistaken entry shouldn't create
// a new negative-debt row).
function reduceCarryoverDebt_(tenantName, amount) {
  const sheet = carryoverSheet_();
  const rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (fuzzyNameMatch_(tenantName, String(rows[i][0] || ''))) {
      var newVal = Math.max(0, (Number(rows[i][1]) || 0) - Number(amount));
      sheet.getRange(i + 1, 2).setValue(newVal);
      return true;
    }
  }
  log_('reduceCarryoverDebt_: no חוב מועבר row matched "' + tenantName + '" — skipped');
  return false;
}

function log_(msg) { console.log(msg); }

// Checked by the admin panel BEFORE submitting a new tenant card, so the admin — not the
// script — decides what happens when building+apt already has a card. Returns the existing
// card (with rowNum + displayName) or null.
function checkDuplicateTenantCard(building, apt) {
  return lookupContactByAddr_(building, apt);
}
function writeTenantRow_(sheet, rowNum, map, building, apt, lastName, c1name, c1phone, c1email, c2name, c2phone, c2email, rented, ownerName, ownerPhone, ownerEmail) {
  set_(sheet, rowNum, map, 'building', building);
  set_(sheet, rowNum, map, 'apt', apt);
  set_(sheet, rowNum, map, 'lastName', lastName);
  set_(sheet, rowNum, map, 'c1name', c1name);
  set_(sheet, rowNum, map, 'c1phone', c1phone);
  set_(sheet, rowNum, map, 'c1email', c1email);
  set_(sheet, rowNum, map, 'c2name', c2name || '');
  set_(sheet, rowNum, map, 'c2phone', c2phone || '');
  set_(sheet, rowNum, map, 'c2email', c2email || '');
  set_(sheet, rowNum, map, 'rented', rented ? 'כן' : 'לא');
  set_(sheet, rowNum, map, 'ownerName', rented ? (ownerName || '') : '');
  set_(sheet, rowNum, map, 'ownerPhone', rented ? (ownerPhone || '') : '');
  set_(sheet, rowNum, map, 'ownerEmail', rented ? (ownerEmail || '') : '');
}
function createTenantCard(building, apt, lastName, c1name, c1phone, c1email, c2name, c2phone, c2email, rented, ownerName, ownerPhone, ownerEmail) {
  const ctx = tenantCtx_(), sheet = ctx[0], map = ctx[2];
  const rowNum = sheet.getLastRow() + 1;
  writeTenantRow_(sheet, rowNum, map, building, apt, lastName, c1name, c1phone, c1email, c2name, c2phone, c2email, rented, ownerName, ownerPhone, ownerEmail);
  set_(sheet, rowNum, map, 'status', 'אושר'); // admin-created directly — no review needed
  triggerDashboardRefresh_();
  return getTenantCards();
}
function updateTenantCard(rowNum, building, apt, lastName, c1name, c1phone, c1email, c2name, c2phone, c2email, rented, ownerName, ownerPhone, ownerEmail) {
  const ctx = tenantCtx_(), sheet = ctx[0], map = ctx[2];
  writeTenantRow_(sheet, rowNum, map, building, apt, lastName, c1name, c1phone, c1email, c2name, c2phone, c2email, rented, ownerName, ownerPhone, ownerEmail);
  triggerDashboardRefresh_();
  return getTenantCards();
}
// "Pending" is now purely informational (nothing gets copied anywhere) — just the latest
// per-address cards that haven't been marked אושר/בוטל yet, as a review reminder.
function getPendingTenantSubmissions() {
  return getTenantCards().filter(function (t) { return !t.status; });
}
function confirmTenantSubmission(rowNum) {
  const ctx = tenantCtx_(), sheet = ctx[0], map = ctx[2];
  set_(sheet, rowNum, map, 'status', 'אושר');
  return getPendingTenantSubmissions();
}
function rejectTenantSubmission(rowNum) {
  const ctx = tenantCtx_(), sheet = ctx[0], map = ctx[2];
  set_(sheet, rowNum, map, 'status', 'בוטל');
  return getPendingTenantSubmissions();
}

// ── Calls ─────────────────────────────────────────────────────────────────
function getCalls() {
  const sheet = SpreadsheetApp.openById(CALLS_SHEET_ID).getSheets()[0];
  const rows = sheet.getDataRange().getValues();
  const out = [];
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    if (!r[0]) continue; // blank row
    out.push({
      rowNum: i + 1,
      timestamp: r[0] ? String(r[0]) : '',
      name: r[1], building: r[2], apt: r[3], desc: r[4], location: r[5],
      urgency: r[6], notes: r[8],
      status: r[9] || 'פתוח', active: r[10] || 'פעיל',
      updateTs: r[11] ? String(r[11]) : ''
    });
  }
  return out.reverse(); // newest first
}

function createCall(name, building, apt, desc, location, urgency, notes) {
  const sheet = SpreadsheetApp.openById(CALLS_SHEET_ID).getSheets()[0];
  sheet.appendRow([new Date(), name, building, apt, desc, location, urgency, '', notes || '', 'פתוח', 'פעיל', '']);
  const rowNum = sheet.getLastRow();
  const callNumber = rowNum - 1;

  const contact = lookupContactByAddr_(building, apt);
  const summary = 'קריאה #' + callNumber + '\n' +
    'שם: ' + name + '\n' +
    'בניין ' + building + ' דירה ' + apt + '\n' +
    'תיאור: ' + desc + '\n' +
    'מיקום: ' + location + '\n' +
    'דחיפות: ' + urgency;

  if (contact && contact.email) {
    MailApp.sendEmail(contact.email, 'אישור קבלת קריאה #' + callNumber,
      'התקבלה קריאתך, מספר קריאה: ' + callNumber + '\n\n' + summary);
  }

  notifyAdminsTelegram_('🔧 קריאה חדשה #' + callNumber + ' (נוספה ע"י מנהל)\n' + summary);
  triggerDashboardRefresh_();
  return getCalls();
}

function updateCall(rowNum, name, building, apt, desc, location, urgency, notes, status) {
  const sheet = SpreadsheetApp.openById(CALLS_SHEET_ID).getSheets()[0];
  const oldStatus = sheet.getRange(rowNum, CALL_STATUS_COL).getValue();
  const callNumber = rowNum - 1;

  sheet.getRange(rowNum, 2, 1, 6).setValues([[name, building, apt, desc, location, urgency]]);
  sheet.getRange(rowNum, 9).setValue(notes || '');
  sheet.getRange(rowNum, CALL_STATUS_COL).setValue(status);

  if (status !== oldStatus) {
    sheet.getRange(rowNum, CALL_STATUS_TS_COL).setValue(new Date());
    const contact = lookupContactByAddr_(building, apt);
    if (contact && contact.email) {
      MailApp.sendEmail(contact.email, 'עדכון סטטוס קריאה #' + callNumber,
        'הסטטוס של קריאה #' + callNumber + ' עודכן ל: ' + status);
    }
    notifyAdminsTelegram_('📋 סטטוס קריאה #' + callNumber + ' עודכן ל: ' + status + ' (עודכן ע"י מנהל)');
  }

  triggerDashboardRefresh_();
  return getCalls();
}

function updateCallStatus(rowNum, newStatus) {
  const sheet = SpreadsheetApp.openById(CALLS_SHEET_ID).getSheets()[0];
  const rowData = sheet.getRange(rowNum, 1, 1, 4).getValues()[0];
  const callNumber = rowNum - 1;
  const building = rowData[2], apt = rowData[3];

  sheet.getRange(rowNum, CALL_STATUS_COL).setValue(newStatus);
  sheet.getRange(rowNum, CALL_STATUS_TS_COL).setValue(new Date());

  const contact = lookupContactByAddr_(building, apt);
  if (contact && contact.email) {
    MailApp.sendEmail(contact.email, 'עדכון סטטוס קריאה #' + callNumber,
      'הסטטוס של קריאה #' + callNumber + ' עודכן ל: ' + newStatus);
  }
  notifyAdminsTelegram_('📋 סטטוס קריאה #' + callNumber + ' עודכן ל: ' + newStatus + ' (עודכן ע"י מנהל)');
  triggerDashboardRefresh_();
  return getCalls();
}

function toggleCallActive(rowNum, active) {
  const sheet = SpreadsheetApp.openById(CALLS_SHEET_ID).getSheets()[0];
  sheet.getRange(rowNum, CALL_ACTIVE_COL).setValue(active ? 'פעיל' : 'לא פעיל');
  triggerDashboardRefresh_();
  return getCalls();
}

// ── Announcements ────────────────────────────────────────────────────────
// Safety net matching apps_script_announcements.js's TEST_MODE: while true, admin-created/
// edited announcements email only the TEST_EMAIL_1/TEST_EMAIL_2 script properties (add them
// here too, separately from the announcements project's own copy). Flip to false once ready
// to let this panel email all 14 real tenants directly.
const ANN_TEST_MODE = true;
function annTestEmails_() { return [prop_('TEST_EMAIL_1'), prop_('TEST_EMAIL_2')].filter(function (e) { return e; }); }
function allTenantEmails_() {
  const emails = [];
  getTenantCards().forEach(function (t) {
    if (t.contact1 && t.contact1.email) emails.push(t.contact1.email);
    if (t.contact2 && t.contact2.email) emails.push(t.contact2.email);
  });
  return emails;
}
function annRecipientEmails_() { return ANN_TEST_MODE ? annTestEmails_() : allTenantEmails_(); }

function emailAnnouncement_(title, content, category, priority) {
  const recipients = annRecipientEmails_();
  if (!recipients.length) return;
  MailApp.sendEmail({
    to: ADMIN_EMAIL, bcc: recipients.join(','),
    subject: '📢 הודעה מהוועד: ' + title,
    body: title + '\n\n' + content + '\n\nקטגוריה: ' + category + '\nעדיפות: ' + priority
  });
}

function createAnnouncement(date, title, content, category, priority, active, validDays) {
  const sheet = SpreadsheetApp.openById(ANNOUNCEMENTS_SHEET_ID).getSheets()[0];
  sheet.appendRow([new Date(), '', date || '', title, content, category, priority, active ? 'כן' : 'לא', validDays || '']);
  if (date) lockDateAsText_(sheet, sheet.getLastRow(), 3, date); // avoid Sheets auto-reparsing the date text, same issue as charges

  notifyAdminsTelegram_('📢 הודעה חדשה נוספה ע"י מנהל\nכותרת: ' + title + '\nקטגוריה: ' + category + '\nעדיפות: ' + priority);
  if (active) emailAnnouncement_(title, content, category, priority);

  triggerDashboardRefresh_();
  return getAnnouncements();
}

function updateAnnouncement(rowNum, date, title, content, category, priority, active, validDays) {
  const sheet = SpreadsheetApp.openById(ANNOUNCEMENTS_SHEET_ID).getSheets()[0];
  const wasActive = sheet.getRange(rowNum, ANN_ACTIVE_COL).getValue() === 'כן';
  sheet.getRange(rowNum, 3, 1, 7).setValues([[date || '', title, content, category, priority, active ? 'כן' : 'לא', validDays || '']]);
  if (date) lockDateAsText_(sheet, rowNum, 3, date);

  if (active && !wasActive) emailAnnouncement_(title, content, category, priority); // only on the inactive->active transition, avoids re-notifying on every edit

  triggerDashboardRefresh_();
  return getAnnouncements();
}

function getAnnouncements() {
  const sheet = SpreadsheetApp.openById(ANNOUNCEMENTS_SHEET_ID).getSheets()[0];
  const rows = sheet.getDataRange().getValues();
  const out = [];
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    if (!r[3]) continue; // no title = blank row
    out.push({
      rowNum: i + 1, date: r[2] ? String(r[2]) : '', title: r[3], content: r[4],
      category: r[5], priority: r[6], active: r[7] || 'לא', validDays: r[8] || ''
    });
  }
  return out.reverse();
}

function toggleAnnouncementActive(rowNum, active) {
  const sheet = SpreadsheetApp.openById(ANNOUNCEMENTS_SHEET_ID).getSheets()[0];
  sheet.getRange(rowNum, ANN_ACTIVE_COL).setValue(active ? 'כן' : 'לא');
  triggerDashboardRefresh_();
  return getAnnouncements();
}

// ── Charges ───────────────────────────────────────────────────────────────
function chargesSheet_() { return SpreadsheetApp.openById(CHARGES_SHEET_ID).getSheetByName('גביות'); }
function paymentsSheet_() { return SpreadsheetApp.openById(CHARGES_SHEET_ID).getSheetByName('תשלומים'); }

function getCharges() {
  const rows = chargesSheet_().getDataRange().getValues();
  const out = [];
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    if (!r[0]) continue;
    out.push({ id: String(r[0]), name: r[1], amount: r[2], date: r[3] ? String(r[3]) : '', active: r[4] || 'לא', description: r[5] || '' });
  }
  return out.reverse();
}

// Google Sheets auto-detects date-looking text and silently reparses it using its own
// locale (e.g. "05/09/2026" read as May 9, not Sept 5) — corrupting the stored value the
// moment it's written. Forcing the cell's number format to plain text ('@') AFTER the
// write stops this: re-applying setValue under that format keeps our exact DD/MM/YYYY
// string, immune to locale-based reinterpretation.
function lockDateAsText_(sheet, row, col, dateStr) {
  sheet.getRange(row, col).setNumberFormat('@').setValue(dateStr);
}
// Same trick, generic name — Sheets auto-detects any numeric-looking typed value (a phone
// number like "054...", "03...") as a real Number and silently drops the leading zero.
// Locking the cell to plain-text format before writing prevents that.
function lockTextValue_(sheet, row, col, value) {
  sheet.getRange(row, col).setNumberFormat('@').setValue(value);
}

function createCharge(name, amount, date, description, active) {
  const sheet = chargesSheet_();
  const id = Utilities.getUuid().split('-')[0];
  if (active) deactivateAllCharges_();
  sheet.appendRow([id, name, Number(amount), date, active ? 'כן' : 'לא', description || '']);
  lockDateAsText_(sheet, sheet.getLastRow(), 4, date);
  triggerDashboardRefresh_();
  return getCharges();
}

function updateCharge(id, name, amount, date, description) {
  const sheet = chargesSheet_();
  const rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) {
      const activeVal = rows[i][4]; // preserve active flag — this edits fields only, not active state
      sheet.getRange(i + 1, 2, 1, 5).setValues([[name, Number(amount), date, activeVal, description || '']]);
      lockDateAsText_(sheet, i + 1, 4, date);
      break;
    }
  }
  triggerDashboardRefresh_();
  return getCharges();
}

function setChargeActive(id, active) {
  const sheet = chargesSheet_();
  const rows = sheet.getDataRange().getValues();
  if (active) deactivateAllCharges_();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) {
      sheet.getRange(i + 1, 5).setValue(active ? 'כן' : 'לא');
      break;
    }
  }
  triggerDashboardRefresh_();
  return getCharges();
}

// Deliberate, rare, hard delete — distinct from setChargeActive, which never removes
// anything. Also removes this charge's own rows in the payments sheet, so a deleted
// test/mistaken charge doesn't leave orphaned payment data behind (same pattern as
// deleteSupplier for supplier custom-fields/reminders).
function deleteCharge(id) {
  const sheet = chargesSheet_();
  const rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
  const psheet = paymentsSheet_();
  const prows = psheet.getDataRange().getValues();
  for (var j = prows.length - 1; j >= 1; j--) {
    if (String(prows[j][0]) === String(id)) psheet.deleteRow(j + 1);
  }
  triggerDashboardRefresh_();
  return getCharges();
}

function deactivateAllCharges_() {
  const sheet = chargesSheet_();
  const rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][4] === 'כן') sheet.getRange(i + 1, 5).setValue('לא');
  }
}

function getPayments(chargeId) {
  const rows = paymentsSheet_().getDataRange().getValues();
  const out = {};
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(chargeId)) out[rows[i][1]] = rows[i][2];
  }
  return out;
}

function setPayment(chargeId, tenantName, amount) {
  const sheet = paymentsSheet_();
  const rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(chargeId) && rows[i][1] === tenantName) {
      sheet.getRange(i + 1, 3).setValue(Number(amount));
      sheet.getRange(i + 1, 4).setValue(new Date());
      triggerDashboardRefresh_();
      return getPayments(chargeId);
    }
  }
  sheet.appendRow([chargeId, tenantName, Number(amount), new Date()]);
  triggerDashboardRefresh_();
  return getPayments(chargeId);
}

// ── Suppliers (כרטיסי ספקים) ────────────────────────────────────────────
// 3 tabs, linked by שם ספק (supplier name) rather than a separate id column.
// ספקים        : שם ספק | קטגוריה | איש קשר | טלפון | טלפון נוסף | מייל | הערות | פעיל
// שדות מותאמים : שם ספק | תווית | ערך
// תזכורות      : שם ספק | סוג | תיאור | תדירות | תאריך יעד | סכום | פעיל | סף_התראה_אחרון
// Matches by TRIMMED tab name, not getSheetByName's exact match — a stray leading/trailing
// space in a tab name (easy to introduce when typing/renaming) would otherwise silently
// return null here and throw downstream on .getDataRange().
function findSheetLoose_(name) {
  const sheets = SpreadsheetApp.openById(SUPPLIERS_SHEET_ID).getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().trim() === name) return sheets[i];
  }
  return null;
}
function suppliersSheet_()         { return findSheetLoose_('ספקים'); }
function supplierFieldsSheet_()    { return findSheetLoose_('שדות מותאמים'); }
function supplierRemindersSheet_() { return findSheetLoose_('תזכורות'); }

function dateStr_(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  return String(v);
}
function parseDate_(v) {
  if (!v) return null;
  if (v instanceof Date) return v;
  var m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(String(v).trim());
  if (!m) return null;
  return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
}
// Badge shown on each reminder in the admin UI — computed fresh on every read, not stored.
function reminderBadge_(dueDate) {
  var d = parseDate_(dueDate);
  if (!d) return { status: 'ok', daysUntil: null, label: '' };
  var today = new Date(); today.setHours(0, 0, 0, 0);
  d.setHours(0, 0, 0, 0);
  var daysUntil = Math.round((d - today) / 86400000);
  if (daysUntil < 0) return { status: 'late', daysUntil: daysUntil, label: 'עבר ב-' + (-daysUntil) + ' ימים' };
  if (daysUntil <= 30) return { status: 'soon', daysUntil: daysUntil, label: 'בעוד ' + daysUntil + ' ימים' };
  return { status: 'ok', daysUntil: daysUntil, label: 'בעוד ' + daysUntil + ' ימים' };
}
function rowToSupplier_(r, rowNum) {
  // String(...) on every field — Sheets auto-detects a numeric-looking typed value (e.g. a
  // phone number like "103") as a real Number, which breaks .replace()/other string methods
  // downstream in the admin UI if left as-is.
  return {
    rowNum: rowNum, name: r[0] ? String(r[0]) : '', category: r[1] ? String(r[1]) : '', contact: r[2] ? String(r[2]) : '',
    phone: r[3] ? String(r[3]) : '', phone2: r[4] ? String(r[4]) : '', email: r[5] ? String(r[5]) : '', notes: r[6] ? String(r[6]) : '',
    active: String(r[7] || '').trim() !== 'לא' // default כן when blank
  };
}
function rowToField_(r, rowNum) {
  return { rowNum: rowNum, supplierName: r[0] ? String(r[0]) : '', label: r[1] ? String(r[1]) : '', value: r[2] ? String(r[2]) : '' };
}
function rowToReminder_(r, rowNum) {
  var dueDate = dateStr_(r[4]);
  return {
    rowNum: rowNum, supplierName: r[0] ? String(r[0]) : '', type: r[1] ? String(r[1]) : '', description: r[2] ? String(r[2]) : '',
    frequency: r[3] ? String(r[3]) : '', dueDate: dueDate, amount: r[5] ? String(r[5]) : '',
    active: String(r[6] || '').trim() !== 'לא', lastThreshold: r[7] || '',
    badge: reminderBadge_(dueDate)
  };
}
// Returns every supplier with its custom fields + reminders nested in, matched by name.
function getSuppliers() {
  const sRows = suppliersSheet_().getDataRange().getValues();
  const fRows = supplierFieldsSheet_().getDataRange().getValues();
  const rRows = supplierRemindersSheet_().getDataRange().getValues();

  const suppliers = [];
  for (var i = 1; i < sRows.length; i++) {
    if (!sRows[i][0]) continue; // blank row
    suppliers.push(rowToSupplier_(sRows[i], i + 1));
  }
  suppliers.sort(function (a, b) { return a.name.localeCompare(b.name, 'he'); });

  suppliers.forEach(function (s) {
    s.fields = [];
    for (var i = 1; i < fRows.length; i++) {
      if (fRows[i][0] === s.name) s.fields.push(rowToField_(fRows[i], i + 1));
    }
    s.reminders = [];
    for (var j = 1; j < rRows.length; j++) {
      if (rRows[j][0] === s.name) s.reminders.push(rowToReminder_(rRows[j], j + 1));
    }
  });
  return suppliers;
}
function createSupplier(name, category, contact, phone, phone2, email, notes) {
  name = String(name || '').trim();
  if (!name) throw new Error('שם ספק חסר');
  const sheet = suppliersSheet_();
  sheet.appendRow([name, category || '', contact || '', phone || '', phone2 || '', email || '', notes || '', 'כן']);
  const rowNum = sheet.getLastRow();
  lockTextValue_(sheet, rowNum, 4, phone || '');
  lockTextValue_(sheet, rowNum, 5, phone2 || '');
  triggerDashboardRefresh_();
  return getSuppliers();
}
// Renaming a supplier cascades to its linked rows in the other two tabs, so custom fields
// and reminders never silently detach from a renamed supplier.
function updateSupplier(rowNum, name, category, contact, phone, phone2, email, notes, active) {
  name = String(name || '').trim();
  if (!name) throw new Error('שם ספק חסר');
  const sheet = suppliersSheet_();
  const oldName = sheet.getRange(rowNum, 1).getValue();
  sheet.getRange(rowNum, 1, 1, 8).setValues([[name, category || '', contact || '', phone || '', phone2 || '', email || '', notes || '', active ? 'כן' : 'לא']]);
  lockTextValue_(sheet, rowNum, 4, phone || '');
  lockTextValue_(sheet, rowNum, 5, phone2 || '');
  if (oldName && oldName !== name) {
    renameLinkedRows_(supplierFieldsSheet_(), oldName, name);
    renameLinkedRows_(supplierRemindersSheet_(), oldName, name);
  }
  triggerDashboardRefresh_();
  return getSuppliers();
}
function renameLinkedRows_(sheet, oldName, newName) {
  const rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === oldName) sheet.getRange(i + 1, 1).setValue(newName);
  }
}
function toggleSupplierActive(rowNum, active) {
  suppliersSheet_().getRange(rowNum, 8).setValue(active ? 'כן' : 'לא');
  triggerDashboardRefresh_();
  return getSuppliers();
}

function createSupplierField(supplierName, label, value) {
  const sheet = supplierFieldsSheet_();
  sheet.appendRow([supplierName, label || '', value || '']);
  lockTextValue_(sheet, sheet.getLastRow(), 3, value || ''); // custom values are often phone-like numbers too
  return getSuppliers();
}
function updateSupplierField(rowNum, label, value) {
  const sheet = supplierFieldsSheet_();
  sheet.getRange(rowNum, 2, 1, 2).setValues([[label || '', value || '']]);
  lockTextValue_(sheet, rowNum, 3, value || '');
  return getSuppliers();
}
function deleteSupplierField(rowNum) {
  supplierFieldsSheet_().deleteRow(rowNum);
  return getSuppliers();
}
// Deliberate, rare, hard delete — distinct from the active/inactive toggle, which never
// removes anything. Also removes this supplier's own custom-field/reminder rows so they
// don't linger as orphaned data with no card to belong to.
function deleteSupplier(rowNum) {
  const sheet = suppliersSheet_();
  const name = sheet.getRange(rowNum, 1).getValue();
  sheet.deleteRow(rowNum);
  deleteLinkedRows_(supplierFieldsSheet_(), name);
  deleteLinkedRows_(supplierRemindersSheet_(), name);
  triggerDashboardRefresh_();
  return getSuppliers();
}
function deleteLinkedRows_(sheet, name) {
  const rows = sheet.getDataRange().getValues();
  for (var i = rows.length - 1; i >= 1; i--) {
    if (rows[i][0] === name) sheet.deleteRow(i + 1);
  }
}

// lockDateAsText_ on the date column every time it's written — otherwise Sheets silently
// re-parses a date-looking string like "01/08/2027" using its own locale, which can swap
// day/month (turned it into 8 Jan instead of 1 Aug here). Same fix already applied to
// charges/announcements dates; this was missed for supplier reminders until now.
function createSupplierReminder(supplierName, type, description, frequency, dueDate, amount) {
  const sheet = supplierRemindersSheet_();
  sheet.appendRow([supplierName, type || '', description || '', frequency || '', dueDate || '', amount || '', 'כן', '']);
  if (dueDate) lockDateAsText_(sheet, sheet.getLastRow(), 5, dueDate);
  return getSuppliers();
}
function updateSupplierReminder(rowNum, type, description, frequency, dueDate, amount) {
  // Editing the date/frequency starts the notification cycle over, so a corrected date
  // doesn't get silently skipped by a threshold that already fired for the old date.
  const sheet = supplierRemindersSheet_();
  sheet.getRange(rowNum, 2, 1, 5).setValues([[type || '', description || '', frequency || '', dueDate || '', amount || '']]);
  if (dueDate) lockDateAsText_(sheet, rowNum, 5, dueDate);
  sheet.getRange(rowNum, 8).setValue('');
  return getSuppliers();
}
function toggleSupplierReminderActive(rowNum, active) {
  supplierRemindersSheet_().getRange(rowNum, 7).setValue(active ? 'כן' : 'לא');
  return getSuppliers();
}

function advanceReminderDate_(date, frequency) {
  var d = new Date(date);
  if (frequency === 'חודשי') d.setMonth(d.getMonth() + 1);
  else if (frequency === 'רבעוני') d.setMonth(d.getMonth() + 3);
  else if (frequency === 'חצי שנתי') d.setMonth(d.getMonth() + 6);
  else if (frequency === 'שנתי') d.setFullYear(d.getFullYear() + 1);
  else return null; // חד-פעמי / אחר — no auto-advance
  return d;
}
// Daily time-driven trigger (add via Triggers -> checkSupplierReminders -> Time-driven ->
// Day timer). For every active reminder: fires Telegram + email once per threshold
// (30/15/7/3/1 days out), tracked via סף_התראה_אחרון so re-running the same day never
// double-notifies; once overdue, a recurring reminder auto-advances to its next cycle.
function checkSupplierReminders() {
  const sheet = supplierRemindersSheet_();
  const rows = sheet.getDataRange().getValues();
  const today = new Date(); today.setHours(0, 0, 0, 0);

  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    var supplierName = r[0], type = r[1], description = r[2], frequency = r[3];
    var dueDateRaw = r[4], amount = r[5], active = r[6], lastThreshold = r[7];
    if (!supplierName || String(active).trim() === 'לא') continue;
    var due = parseDate_(dateStr_(dueDateRaw));
    if (!due) continue;
    due.setHours(0, 0, 0, 0);
    var daysUntil = Math.round((due - today) / 86400000);

    if (daysUntil < 0) {
      var next = advanceReminderDate_(due, frequency);
      if (next) {
        lockDateAsText_(sheet, i + 1, 5, Utilities.formatDate(next, Session.getScriptTimeZone(), 'dd/MM/yyyy'));
        sheet.getRange(i + 1, 8).setValue('');
      }
      continue;
    }
    if (SUPPLIER_REMINDER_THRESHOLDS.indexOf(daysUntil) !== -1 && String(lastThreshold) !== String(daysUntil)) {
      var amountTxt = amount ? (' | סכום: ₪' + amount) : '';
      var msg = '🔔 תזכורת ספק: ' + supplierName + ' — ' + description + ' (' + type + ')\n' +
        'בעוד ' + daysUntil + ' ימים | יעד: ' + dateStr_(due) + amountTxt;
      notifyAdminsTelegram_(msg);
      MailApp.sendEmail(ADMIN_EMAIL, 'תזכורת ספק: ' + supplierName + ' — בעוד ' + daysUntil + ' ימים', msg);
      sheet.getRange(i + 1, 8).setValue(String(daysUntil));
    }
  }
}
