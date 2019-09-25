// 
function sendInvoice(Investor, EMail, invoiceNo, date, file,to, subject, msg) {
  
  var fn = "COMPANY NAME"; 
  var email = EMail; 
  
  var subj = Investor + ' - Invoice - ' + invoiceNo  
  Logger.log(file);
  var payload = '<html>\r\n<body>\r\n\r\n<h2>Hey '+ fn + " & " + Investor + ',<\/h2>\r\n<br>\r\n\r\n<p> Here is your invoice for the month, make sure to keep these for the end of the year for filing taxes.  <\/p>'
  Logger.log(email);
  
  GmailApp.sendEmail(email, subj, msg,{
    attachments: [file.next().getAs(MimeType.PDF)],
    htmlBody: payload
  })
  
    
    
}
