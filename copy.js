function copiar() {
  Logger.log("Script start.")
  while(true){
    try{
      countIf(); //updates the formulas to avoid duplicates
      var fmeaData = SpreadsheetApp.openById("___");
      SpreadsheetApp.flush();
      var copyFrom = fmeaData.getSheetByName("Enviar");
      var fmeaReport = SpreadsheetApp.openById("___");
      var destination = fmeaReport.getSheetByName("Report");
      var lastRow = copyFrom.getLastRow();
      if (lastRow > 2){
        var lastCol = copyFrom.getLastColumn();
        var range = copyFrom.getRange(2,2,lastRow-1,lastCol-1);
        var data = range.getValues();
        SpreadsheetApp.flush();
        var dataRegion = destination.getRange("D1").getDataRegion();
        var lastActualColumn = dataRegion.getLastRow();
        destination.getRange(lastActualColumn+1,4,lastRow-1,lastCol-1).setValues(data);
        destination.getRange(lastActualColumn+1,2,lastRow-1,1).setValue(new Date());
        SpreadsheetApp.flush();
        Logger.log("Added",String(lastRow-1),"rows.");
        break;
      }
      Logger.log("#N/A, no new data to add.");
      break;
    }
    catch(e){
      Logger.log(e);
    }
  }
}

function copyTrigger() {
  // Trigger every day at 03:30 
  ScriptApp.newTrigger("copiar")
      .timeBased()
      .atHour(2)
      .nearMinute(30)
      .everyDays(1)
      .inTimezone("America/Bogota")
      .create();
}

function customMenu() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Copiar')
      .addItem('Add to FMEA report', 'copiar')
      .addToUi();
}

function openTrigger(){
ScriptApp.newTrigger('customMenu')
  .forSpreadsheet('____')
  .onOpen()
  .create();
}
