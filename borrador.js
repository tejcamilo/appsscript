function getFirstEmptyRowByColumnArray(sheet) {
  var column = sheet.getRange('A:A');
  var values = column.getValues(); // get all data in one call
  var ct = 0;
  while ( values[ct] && values[ct][0] != "" ) {
    ct++;
  }
  return (ct+1);
}

function runTest(){
  revSort();
  var getSpreadSheet = SpreadsheetApp.openById("___");
  var sheet = getSpreadSheet.getSheetByName("Report");
  var row = parseInt(getFirstEmptyRowByColumnArray(sheet).toFixed(0)); //last non empty row
  var lastRow = sheet.getLastRow();
  Logger.log(row+' ',lastRow);
  
  lastRow -= (row-1); //new rows that need to be added, **debemos restar "-1" a row ya que row nos retorna la siguiente celda que no contiene datos
  Logger.log(lastRow);
  r6(sheet,row,lastRow);
  sort();
}

function r6(sheet,r,lastRow){
  Logger.log("autoDrag6 start.");
  while(true){
    try{
      var range = sheet.getRange('V'+r)
      .setFormula('=IF(ISBLANK(D'+r+'),,IFERROR(VLOOKUP(F'+r+'&K'+r+',Counter!A2:B,2,0)+COUNTIFS(E$1:E'
      +(r-1)+',E'+r+',K$1:K'+(r-1)+',K'+r+'),COUNTIFS(E$1:E'+(r-1)+',E'+r+',K$1:K'+(r-1)+',K'+r+'))');
      var destination = sheet.getRange(r,22,lastRow,1);
      range.copyTo(destination);
      /**
      var values = destination.getValues();
      destination.setFormula("");
      destination.setValues(values);
      */
      Logger.log("autoDrag6 complete.");
      break;
    }
    catch(e){
      Logger.log(e);
    }
  }
}
