function getFirstEmptyRowByColumnArray(sheet) {
  var column = sheet.getRange('A:A');
  var values = column.getValues(); // get all data in one call
  var ct = 0;
  while ( values[ct] && values[ct][0] != "" ) {
    ct++;
  }
  return (ct+1);
}

function test(){
  revSort();
  var getSpreadSheet = SpreadsheetApp.openById("____");
  var sheet = getSpreadSheet.getSheetByName("Report");
  var row = parseInt(getFirstEmptyRowByColumnArray(sheet).toFixed(0)); //last non empty row
  var lastRow = sheet.getLastRow();
    if(lastRow-(row-1) == 0){
    Logger.log("No hay nada nuevo para actualizar")
    return 0;
  }
  //PASSED autoDrag1(sheet,lastRow-1); 
  autoDrag2(sheet,lastRow-1);
}

function runAll(){
  revSort();
  var getSpreadSheet = SpreadsheetApp.openById("____");
  var sheet = getSpreadSheet.getSheetByName("Report");
  var row = parseInt(getFirstEmptyRowByColumnArray(sheet).toFixed(0)); //last non empty row
  var lastRow = sheet.getLastRow();
  if(lastRow-(row-1) == 0){
    Logger.log("No hay nada nuevo para actualizar")
    return 0;
  }
  Logger.log(row+' '+lastRow);
  autoDrag1(sheet,lastRow-1); //debemos restar "-1" ya que en la función "getRange" tomamos los valores a partir de la 2da fila 
  autoDrag2(sheet,lastRow-1); //debemos restar "-1" ya que en la función "getRange" tomamos los valores a partir de la 2da fila 
  lastRow -= (row-1); //new rows that need to be added, **debemos restar "-1" a row ya que row nos retorna la siguiente celda que no contiene datos
  Logger.log(lastRow);
  autoDrag3(sheet,row,lastRow);
  autoDrag4(sheet,row,lastRow);
  autoDrag5(sheet,row,lastRow);
  autoDrag6(sheet,row,lastRow-1); //This -1 needs to be there for the range to work correctly
  autoDrag7(sheet,row,lastRow);
  sort();
}

function autoDrag1(sheet,lastRow){
  Logger.log("autoDrag1 start.")
  while(true){
    try{
      var destination = sheet.getRange(2,3,lastRow,1);
      destination.setValue('');
      sheet.getRange('C2').setFormula("=ARRAYFORMULA(IF(ISBLANK(D2:D),,VLOOKUP(F2:F&DATEVALUE(TODAY()),Roster!A:E,4,1)))");
      var values = destination.getValues();
      destination.setFormula("");
      destination.setValues(values);
      Logger.log("autoDrag1 complete.");
      break;
    }
    catch(e){
      Logger.log(e);
    }
  }
}

function autoDrag2(sheet,lastRow){ 
  Logger.log("autoDrag2 start.");
  while(true){
    try{
      var destination = sheet.getRange(2,23,lastRow,1);
      destination.setValue('');
      sheet.getRange('W2').setFormula('=ARRAYFORMULA(IF(ISBLANK(D2:D),,VLOOKUP(F2:F&DATEVALUE(W1),Roster!A:E,5)))');
      var values = destination.getValues();
      destination.setFormula("");
      destination.setValues(values);
      Logger.log("autoDrag2 complete.");
      break;
    }
    catch(e){
      Logger.log(e);
    }
  }
}

function autoDrag3(sheet,r,lastRow){
  Logger.log("autoDrag3 start.");
    while(true){
    try{
      var range = sheet.getRange('A'+r)
      .setFormula('=IF(ISBLANK(K'+r+'),,VLOOKUP(K'+r+'&I'+r+',Codes!$C:$D,2,0)&(COUNTIFS(I$1:I'+(r-1)+',I'+r+',K$1:K'+(r-1)+',K'+r+')+1))');
      var destination = sheet.getRange(r,1,lastRow,1);
      range.copyTo(destination);
      var values = destination.getValues();
      destination.setFormula("");
      destination.setValues(values);
      Logger.log("autoDrag3 complete.");
      break;
    }
    catch(e){
      Logger.log(e);
    }
  }
}

function autoDrag4(sheet,r,lastRow){
  Logger.log("autoDrag4 start...")
  while(true){
    try{
      var destination = sheet.getRange(r,18,lastRow,1);
      destination.setValue('');
      sheet.getRange('R'+r)
        .setFormula('=ARRAYFORMULA(IF(ISBLANK(B'+r+':B),,IF(WEEKDAY(B'+r+':B)=1,B'+r+':B+3,IF(WEEKDAY(B'+r+':B)<5,(B'+r+':B+2),B'+r+':B+4))))');
      var values = destination.getValues();
      destination.setFormula("");
      destination.setValues(values);
      Logger.log("autoDrag4 completed.");
      break;
    }
    catch(e){
      Logger.log(e);
    }
  }
}

function autoDrag5(sheet,r,lastRow){
  Logger.log("autoDrag5 start.")
  while(true){
    try{
      var destination = sheet.getRange(r,21,lastRow,1);
      destination.setValue('');
      sheet.getRange('U'+r)
        .setFormula('=ARRAYFORMULA(IF(ISBLANK(D'+r+':D),,F'+r+':F&LEFT(A'+r+
        ':A,3)&IF(ISERROR(FIND("=",P'+r+':P)),REGEXEXTRACT(P'+r+':P,"[^/]+$"),REGEXEXTRACT(P'+r+':P,"[^=]+$"))))');
      var values = destination.getValues();
      destination.setFormula("");
      destination.setValues(values);
      Logger.log("autoDrag5 complete.");
      break;
    }
    catch(e){
      Logger.log(e);
    }
  }
}

function autoDrag6(sheet,r,lastRow){
  /**
   * if it ever fails again, here's the original formula from row 2.
   * =IF(ISBLANK(D2),,IFERROR(VLOOKUP(F2&K2,Counter!A2:B,2,0)+COUNTIFS(E$1:E1,E2,K$1:K1,K2),COUNTIFS(E$1:E1,E2,K$1:K1,K2)))
   */
  Logger.log("autoDrag6 start.");
  while(true){
    try{
      var range = sheet.getRange('S'+r)
      .setFormula('=IF(ISBLANK(D'+r+'),,IFERROR(VLOOKUP(F'+r+'&K'+r+',Counter!A$2:B,2,0)+COUNTIFS(E$1:E'
      +(r-1)+',E'+r+',K$1:K'+(r-1)+',K'+r+'),COUNTIFS(E$1:E'+(r-1)+',E'+r+',K$1:K'+(r-1)+',K'+r+')))');
      var destination = sheet.getRange(r,19,lastRow,1);
      range.copyTo(destination);
      var values = destination.getValues();
      destination.setFormula("");
      destination.setValues(values);
      Logger.log("autoDrag6 complete.");
      break;
    }
    catch(e){
      Logger.log(e);
    }
  }
}

function autoDrag7(sheet,r,lastRow){////
  Logger.log("autoDrag7 start.");
  while(true){
    try{
      var range = sheet.getRange('V'+r)
      .setFormula('=IF(ISBLANK(D'+r+'),,COUNTIF(U$1:U'+(r-1)+
      ',(F'+r+'&LEFT(A'+r+',3)&IF(ISERROR(FIND("=",P'+r+')),REGEXEXTRACT(P'+r+',"[^/]+$"),REGEXEXTRACT(P'+r+',"[^=]+$")))))');
      var destination = sheet.getRange(r,22,lastRow,1);
      range.copyTo(destination);
      var values = destination.getValues();
      destination.setFormula("");
      destination.setValues(values);
      Logger.log("autoDrag7 complete.");
      break;
    }
    catch(e){
      Logger.log(e);
    }
  }
}



function formulaTrigger() {
  // Trigger every day at 03:10 
  ScriptApp.newTrigger("runAll")
      .timeBased()
      .atHour(3)
      .nearMinute(10)
      .everyDays(1)
      .inTimezone("America/Bogota")
      .create();
}
