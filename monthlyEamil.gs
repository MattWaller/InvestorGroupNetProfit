function sendmEmail(to, subject, msg) {
  var sheetG = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Graph");
  var today = sheetG.getRange(1,1).getValue();
  var FileName = String(today);
  var app = SpreadsheetApp;
  var sheet = app.getActiveSpreadsheet().getSheetByName("InvestorInfo");
  var rows = sheet.getLastRow();
  var cols = sheet.getLastColumn();
  var dataRange = sheet.getRange(2,1,rows-1,cols)
  var data = dataRange.getValues();
  var i=2;
  Logger.log(FileName);
  var fd = sheet.getRange(1,17).getValue();
  var ld = sheet.getRange(1,18).getValue();
  Logger.log(today);
  for (i in data){  
    if (i < 7){ 
      var file = DriveApp.getFilesByName(String(FileName))
      var fn = data[i][0]
      var email = data[i][1]
      var investorShare = data[i][2]
      var ywp = data[i][3]
      var twp = data[i][4]
      var wpm = data[i][5]
      var ymp = data[i][6]
      var tmp = data[i][7]
      var mpm = data[i][8]
      var tms = data[i][9]
      var gst = data[i][10]
      var tpo = data[i][11]
      var io = data[i][12]
      var subj = 'WC monthly Share of Profits for the month of ' +fd+ ' to ' +ld
      Logger.log(file);
      var payload = '<html>\r\n<body>\r\n\r\n<h2>Hey '+fn+',<\/h2>\r\n<br>\r\n\r\n<p> Here are this months return for your share of ('+investorShare+'%). <\/p><table rules=all style=border-color: #666; cellpadding=25>\r\n<tr style=\'background: #eee;\'><td><strong>Monthly Profits<\/strong> <\/td><td>Amounts<\/td><\/tr>\r\n<tr><td><strong>Your Monthly Profit<\/strong> <\/td><td> $'+ymp+'<\/td><\/tr>\r\n<tr><td><strong>Total Company Profit<\/strong> <\/td><td>$'+tmp+'<\/td><\/tr>\r\n<tr><td><strong>Monthly Profit Margin<\/strong> <\/td><td> '+mpm+'%<\/td><\/tr>\r\n<\/table>\r\n\r\n<br>\r\n<br>\r\n\r\n<table rules=all style=border-color: #666; cellpadding=25>\r\n<tr style=\'background: #eee;\'><td><strong>Estimated Payment this Month<\/strong> <\/td><td>Amounts<\/td><\/tr>\r\n<tr><td><strong>Share this Month<\/strong> <\/td><td> $'+tms+'<\/td><\/tr>\r\n<tr><td><strong>GST<\/strong> <\/td><td>$'+gst+'<\/td><\/tr>\r\n<tr><td><strong>Total Payout<\/strong> <\/td><td> $'+tpo+'<\/td><\/tr>\r\n<tr><td><strong>Investment Outstanding<\/strong> <\/td><td> $'+io+'<\/td><\/tr>\r\n<\/table>\r\n<\/body>\r\n<\/html>'
      Logger.log(email);
      Logger.log(io);
      GmailApp.sendEmail(email, subj, msg,{
        attachments: [file.next().getAs('image/png')],
        htmlBody: payload
      })
      i =+
        Logger.log(i);
      Logger.log(fn);
      Logger.log(fd);
      Logger.log(ld);
      Logger.log(subj);
      Logger.log(email);
    }
  }
}
