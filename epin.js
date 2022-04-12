function onEdit(e){
  try{
  var activeSheet = e.range.getSheet();
  var range = e.range;
  var tabs = ['To review','ePin'];
  var column = range.getColumn();
  if (tabs.includes(activeSheet.getName())){
    if(activeSheet.getName() == tabs[0] && column <= 10){
      var cell = activeSheet.getRange(2,14);
      var destination = activeSheet.getRange(3,14,activeSheet.getLastRow()-1,1);
      cell.copyTo(destination);
      SpreadsheetApp.flush();
      SpreadsheetApp.getUi().alert('Formula updated succesfully!');
    }
    else if(column > 2 && column < 8) {
      var cell = activeSheet.getRange(2,9);
      var destination = activeSheet.getRange(3,9,activeSheet.getLastRow()-1,1);
      cell.copyTo(destination);
      SpreadsheetApp.flush();
      SpreadsheetApp.getUi().alert('Formula updated succesfully!');
    }
  }
  } catch(e){
    var ui = SpreadsheetApp.getUi();
    var response = ui.prompt('Something went wrong ! Please check the script.', 'Type "ack" to dismiss',ui.ButtonSet.OK);
    while (response.getResponseText() !== 'ack'){
      response = ui.prompt('Something went wrong ! Please check the script.', 'Type "ack" to dismiss',ui.ButtonSet.OK);
    }
  }
}

// =ArrayFormula(IF(ISBLANK(A2:A),,IFERROR(VLOOKUP(A2:A&B2:B,ePin!A:K,11,),"ePIN missing!")))
