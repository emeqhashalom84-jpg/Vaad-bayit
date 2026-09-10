
// TEMP TEST — remove after use. Looks up a known address (building 84 apt 2 = אלקיים)
// and logs what comes back. Run from the function dropdown (select testLookup > Run),
// then View > Logs. Should show the full household (אורן + אורלי), not empty/wrong data.
function testLookup() {
  Logger.log(JSON.stringify(lookupContact_('84', '2')));
}
