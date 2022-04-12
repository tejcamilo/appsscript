//*** AGREGA LOS EPIN DE (__) A (___) TODOS LOS DIAS */

function agregarASentEmails() {
  while(true){
    try{
      let fromEPIN = SpreadsheetApp.openById("___-__");
      let pinSheet = fromEPIN.getSheetByName('ePIN');
      let toSentEmails = SpreadsheetApp.openById('___-__-__');
      let detinationSheet = toSentEmails.getSheetByName('ePIN');
      let lastRow = pinSheet.getLastRow();
      if (lastRow > 2){
        let range1 = 'C2:G' + (lastRow-1);
        let range2 = 'I2:I' + (lastRow-1);
        let range3 = 'K2:K' + (lastRow-1);
        let rangoDeCopia = pinSheet.getRangeList([range1, range2, range3]);
        let count = 0;
          for (i in rangoDeCopia){
            i.getValues;
            if (count === 0){
              detinationSheet.getRange(3, detinationSheet.getLastColumn())
            }
          }
      } else {
        Logger.log("No hay datos nuevos");
        break;
      }
    }
    catch(exeption){
      Logger.log(exeption)
    }
  }
}


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
