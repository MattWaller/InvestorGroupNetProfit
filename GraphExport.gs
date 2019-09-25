function GraphDownload() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Graph");
  var today = sheet.getRange(1,1).getValue();
  var chart = sheet.getCharts()[0];
  Logger.log(today);
  var driveFolder = DriveApp.getFolderById('FOLDER_ID');
  
  var file = driveFolder.createFile(chart.getAs('image/png').setName(today));
}
