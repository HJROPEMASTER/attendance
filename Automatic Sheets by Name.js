function createSheetsFromList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var listSheet = ss.getSheetByName("Payroll List"); // 👈 Sheet with the list of names
  if (!listSheet) return;

  var names = listSheet.getRange("A2:A").getValues().flat().filter(name => name !== "");

  names.forEach(function(name) {
    if (!ss.getSheetByName(name)) { // Only create if it doesn't already exist
      ss.insertSheet(name);
    }
  });
}
