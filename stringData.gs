function reset() {
  var vh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("variableHolder");
  var rows = vh.getLastRow()
  var dataRange = vh.getRange(1,1,rows+1,1)
  var data = dataRange.getValues();
  var c = 0;
  Logger.log(data);
  for (i in data){
    if (i < 32){ 
  var pn = data[i][0]
  var c = c + 1
  var impF = vh.getRange(c,2).getFormula();
  
  var s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(pn);
      
  var clear = s.getRange("A:Z").clearContent();
  var setIR = s.getRange(1,1).setFormula(impF);

  }
}
}
