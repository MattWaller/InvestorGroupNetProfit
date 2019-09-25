function wc_detailedOrdersImport() {
 
  var app = SpreadsheetApp;
  var sheet = app.getActiveSpreadsheet().getSheetByName("detailed_orders"); 
  var sheet_cols = sheet.getLastColumn();
  var sheet_rows = sheet.getLastRow();
  
  
  sheet.getRange(1,1,sheet_rows,sheet_cols).clearContent();   
  var file = DriveApp.getFilesByName("wc_orders.csv").next();
  var csvData = Utilities.parseCsv(file.getBlob().getDataAsString());

  
  sheet.getRange(1, 1, csvData.length, csvData[0].length).setValues(csvData);
  

}
