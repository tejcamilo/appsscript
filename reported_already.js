function wait(seconds){ //time in seconds to pause the script for
  for(i=1; i<=seconds; i++){
    Logger.log('Time '+'('+i+'s)');
    Utilities.sleep(1000);
    SpreadsheetApp.flush();
  }
}

function countIf(){
  while(true){
    try{
      var getSpreadSheet = SpreadsheetApp.openById("____");
      var sheet1 = getSpreadSheet.getSheetByName("Report");
      var destination1 = sheet1.getRange(2,1,sheet1.getLastRow()-1,1);
      destination1.setFormula(""); 
      destination1.setValue('');
      wait(10);
      sheet1.getRange('A2')
        .setFormula('=ARRAYFORMULA(IF(ISBLANK(D2:D),,COUNTIF(IMPORTRANGE("____", "Report!O2:O"),M2:M)+COUNTIF(IMPORTRANGE("____","Report!O2:O"),M2:M)+COUNTIF(IMPORTRANGE("175COLcf5Pp1DDYVigk2UYzpk0xaLxZfAIuCAJdzJfYE","Report!O2:O"),M2:M)))');
      var sheet2 = getSpreadSheet.getSheetByName('Enviar');
      var destination2 = sheet2.getRange('A2');
      destination2.setFormula("");
      destination2.setValue("");
      wait(10);
      sheet2.getRange('A2').setFormula('=FILTER(Report!A3:O,Report!A3:A=0,Report!A3:A<>"",Report!B3:B<>ISERROR(Report!B3:B))');
      SpreadsheetApp.flush();
      Logger.log("script completed succesfully.");
      break;
    }
    catch(e){
      Logger.log(e);
    }
  }
}

function countTrigger() {
  // Trigger every day at 23:30 
  ScriptApp.newTrigger("countIf")
      .timeBased()
      .atHour(23)
      .nearMinute(30)
      .everyDays(1)
      .inTimezone("America/Bogota")
      .create();
}

function countifMenu() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Run check')
      .addItem('Run countIF', 'countIf')
      .addToUi();
}

function openTrigger1(){
ScriptApp.newTrigger('countifMenu')
  .forSpreadsheet('____')
  .onOpen()
  .create();
}
