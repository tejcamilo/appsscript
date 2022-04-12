function sort() {
  var getSpreadSheet = SpreadsheetApp.openById("__");
  var sheet = getSpreadSheet.getSheetByName("Report");
  var lastCol = sheet.getLastColumn();
  var lastRow = sheet.getLastRow();
  var range = sheet.getRange(2,1,lastRow-1,lastCol);
  Logger.log(range.getA1Notation());
  range.sort([{column: 3, ascending: true}, {column: 2, ascending: true}]);
}

function sortTrigger() {
  // Trigger every day at 01:50 
  ScriptApp.newTrigger("sort")
      .timeBased()
      .atHour(1)
      .nearMinute(50)
      .everyDays(1)
      .inTimezone("America/Bogota")
      .create();
}

function revSort() { // In case you need to revert sorting
  var getSpreadSheet = SpreadsheetApp.openById("___");
  var sheet = getSpreadSheet.getSheetByName("Report");
  var lastCol = sheet.getLastColumn();
  var lastRow = sheet.getLastRow();
  var range = sheet.getRange(2,1,lastRow-1,lastCol);
  Logger.log(range.getA1Notation());
  range.sort([{column: 2, ascending: true}]);
}
