function sendEmail(to, subject, msg) {
  var as = SpreadsheetApp;
  var sheetz = as.getActiveSpreadsheet().getSheetByName("WeeklyProfit");
  var sDate = sheetz.getRange(1,2).setValue('=today()-7');
  
  
  var app = SpreadsheetApp;
  var sheet = app.getActiveSpreadsheet().getSheetByName("InvestorInfo");
  var rows = sheet.getLastRow()
  var cols = sheet.getLastColumn()
  var dataRange = sheet.getRange(2,1,rows-1,cols)
  var data = dataRange.getValues();
  var i=2;
  //Logger.log(data);
  var fd = sheet.getRange(1,15).getValue();
  var ld = sheet.getRange(1,16).getValue();
  for (i in data){  
    if (i < 7){ 
      
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
      
      var payload = '<html>\r\n<body>\r\n\r\n<h2>Hey '+fn+',<\/h2>\r\n<br>\r\n\r\n<p> Here are this weeks return for your share of ('+investorShare+'%). <\/p>\r\n<br>\r\n\r\n<table rules=all style=border-color: #666; cellpadding=25>\r\n<tr style=\'background: #eee;\'><td><strong>Weekly Profits<\/strong> <\/td><td>Amounts<\/td><\/tr>\r\n<tr><td><strong>Your Weekly Profit<\/strong> <\/td><td>$' +ywp+'<\/td><\/tr>\r\n<tr><td><strong>Total Company Profit<\/strong> <\/td><td>$'+twp+'<\/td><\/tr>\r\n<tr><td><strong>Weekly Profit Margin<\/strong> <\/td><td> '+wpm+'%<\/td><\/tr>\r\n<\/table>\r\n\r\n<br>\r\n<br>\r\n\r\n<table rules=all style=border-color: #666; cellpadding=25>\r\n<tr style=\'background: #eee;\'><td><strong>Monthly Profits<\/strong> <\/td><td>Amounts<\/td><\/tr>\r\n<tr><td><strong>Your Monthly Profit<\/strong> <\/td><td> $'+ymp+'<\/td><\/tr>\r\n<tr><td><strong>Total Company Profit<\/strong> <\/td><td>$'+tmp+'<\/td><\/tr>\r\n<tr><td><strong>Monthly Profit Margin<\/strong> <\/td><td> '+mpm+'%<\/td><\/tr>\r\n<\/table>\r\n\r\n<br>\r\n<br>\r\n\r\n<table rules=all style=border-color: #666; cellpadding=25>\r\n<tr style=\'background: #eee;\'><td><strong>Estimated Payment this Month<\/strong> <\/td><td>Amounts<\/td><\/tr>\r\n<tr><td><strong>Share this Month<\/strong> <\/td><td> $'+tms+'<\/td><\/tr>\r\n<tr><td><strong>GST<\/strong> <\/td><td>$'+gst+'<\/td><\/tr>\r\n<tr><td><strong>Total Payout<\/strong> <\/td><td> $'+tpo+'<\/td><\/tr>\r\n<tr><td><strong>Investment Outstanding<\/strong> <\/td><td> $'+io+'<\/td><\/tr>\r\n<\/table>\r\n<\/body>\r\n<\/html>'
      var subj = 'WC weekly Share of Profits for the week of ' +fd+ ' to ' +ld
      Logger.log(fn);
      Logger.log(email);
      Logger.log(io);
      GmailApp.sendEmail(email, subj, msg,{
        htmlBody: payload
      })
      i =+
        Logger.log(i);
      Logger.log(fd);
      Logger.log(ld);
      Logger.log(subj);
    }
  }
  
  var wDate = sheetz.getRange(1,2).setValue('=today()');
  
  
}
