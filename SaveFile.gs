function tsaveAsSpreadsheet(){ 
  // to change data to string
  var vh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("variableHolder");
  var rows = vh.getLastRow()
  var dataRange = vh.getRange(1,1,rows+1,1)
  var data = dataRange.getValues();
  for (i in data){
    if (i < 32){ 
      var pn = data[i][0]
      var s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(pn);
      var d = s.getRange("A:Z").getValues();
      var set = s.getRange("A:Z").setValues(d);     
    }
  }
  //to save spreadsheet
  var app = SpreadsheetApp;
  var sheet = app.getActiveSpreadsheet().getSheetByName("InvestorInfo");
  var day = sheet.getRange(1,22).getValue();
  var month = sheet.getRange(1,23).getValue();
  var year = sheet.getRange(1,24).getValue();
  var renameFile = "WC-Orders_" +day + "-" +month + "-" + year;  
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log(month);
  var destFolder = DriveApp.getFolderById("FILE_ID"); 
  DriveApp.getFileById(sheet.getId()).makeCopy(renameFile, destFolder); 
  
  //to Reset import Range function
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
