// ONE-TIME SETUP — run this once to create the whole "Vaad — ספקים" spreadsheet with all
// 3 tabs and headers already in place. Steps:
//   1. Go to script.google.com -> New project.
//   2. Paste this whole file into Code.gs, replacing the default content.
//   3. Function dropdown (top) -> select "createSuppliersSheet" -> Run.
//   4. First run asks for authorization -> approve it (normal, same as every other trigger).
//   5. View -> Logs (or Ctrl+Enter) -> copy the URL it prints.
//   6. Open that URL, then File -> Share -> add Michael as Editor, general access Restricted.
//   7. Send Claude the spreadsheet ID (the long string in the URL between /d/ and /edit).
// You can delete this script afterward — it's a one-time setup tool, not part of the
// ongoing system.
function createSuppliersSheet() {
  const ss = SpreadsheetApp.create('Vaad — ספקים');

  const suppliers = ss.getSheets()[0];
  suppliers.setName('ספקים');
  suppliers.getRange(1, 1, 1, 9).setValues([[
    'id', 'שם ספק', 'קטגוריה', 'איש קשר', 'טלפון', 'טלפון נוסף', 'מייל', 'הערות', 'פעיל'
  ]]);
  suppliers.setFrozenRows(1);

  const fields = ss.insertSheet('שדות_מותאמים');
  fields.getRange(1, 1, 1, 3).setValues([[
    'supplier_id', 'תווית', 'ערך'
  ]]);
  fields.setFrozenRows(1);

  const reminders = ss.insertSheet('תזכורות');
  reminders.getRange(1, 1, 1, 8).setValues([[
    'supplier_id', 'סוג', 'תיאור', 'תדירות', 'תאריך יעד', 'סכום', 'פעיל', 'סף_התראה_אחרון'
  ]]);
  reminders.setFrozenRows(1);

  Logger.log('Created: ' + ss.getUrl());
  Logger.log('Spreadsheet ID: ' + ss.getId());
}
