function Main(){
  var InvestorName = ["INVESTOR1","INVESTOR2","INVESTOR3"]
  var NameRow = [4,7,8]
  Invoice(InvestorName,NameRow);
}

function Invoice(InvestorName, NameRow) {
  Logger.log(InvestorName);
  Logger.log(NameRow);
  
  
  var i = 0;
  var Len = InvestorName.length;
  
  if (i < Len){
    while ( i < Len) {
      // run script here
      var Investor = InvestorName[i];
      var t = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(Investor)
      
      
      // Manipulation of Invoices
      // setting formulas
      t.getRange(22,6).setValue("=RemittanceAdvice!H"+ NameRow[i]);
      t.getRange(7,7).setValue("=today()-1");
      t.getRange(22,2).setFormula('=concatenate("Training for the Month of ",vlookup(month(G7),MonthInfo!A:B,2,false)," ",year(G7))')
      t.getRange(22,7).setValue("=E22*F22");
      t.getRange(33,7).setValue("=SUM(G22:G32)");
      t.getRange(35,7).setValue("=G33*G34");  
      t.getRange(36,7).setValue("=G33+G35");    
      
      
      //getting string values
      var name = t.getRange(4,2).getValue();
      var EMail = t.getRange(5,2).getValue();
      var date = t.getRange(7,7).getValue();
      var invoiceNo = t.getRange(9,7).getValue();
      var desc = t.getRange(22,2).getValue();
      var unitprice = t.getRange(22,6).getValue();
      var totalup = t.getRange(22,7).getValue();
      var subtotal = t.getRange(33,7).getValue();
      var totaltax = t.getRange(35,7).getValue();  
      var balanceDue = t.getRange(36,7).getValue();
      
      
      //setting string values
      
      t.getRange(7,7).setValue(date);
      t.getRange(9,7).setValue(invoiceNo + 1);
      t.getRange(22,2).setValue(desc);
      t.getRange(22,6).setValue(unitprice);
      t.getRange(22,7).setValue(totalup);
      t.getRange(33,7).setValue(subtotal);
      t.getRange(35,7).setValue(totaltax);  
      t.getRange(36,7).setValue(balanceDue); 
      
      
      //PDF Files
      
      var x = SpreadsheetApp.getActiveSpreadsheet().getUrl();
      Logger.log(x);
      
      var url = x.replace(/edit$/,'')
      Logger.log(url);
      
      var id = t.getSheetId();
      Logger.log(id);
      
      
      var url_ext = 'export?exportFormat=pdf&format=pdf'   //export as pdf
      + '&gid=' + id  //the sheet's Id
      // following parameters are optional...
      + '&size=letter'      // paper size
      + '&portrait=true'    // orientation, false for landscape
      + '&fitw=true'        // fit to width, false for actual size
      + '&sheetnames=false&printtitle=false&pagenumbers=false'  //hide optional headers and footers
      + '&gridlines=false'  // hide gridlines
      + '&fzr=false';       // do not repeat row headers (frozen rows) on each page
      
      var options = {
        headers: {
          'Authorization': 'Bearer ' +  ScriptApp.getOAuthToken()
        }
      }
      var response = UrlFetchApp.fetch(url + url_ext, options);
      var FileName = Investor + ' - Invoice - ' + invoiceNo +  '.pdf'
      var blob = response.getBlob().setName(FileName);
      var folder = DriveApp.getFoldersByName(Investor).next()
      //from here you should be able to use and manipulate the blob to send and email or create a file per usual.
      //In this example, I save the pdf to drive
      folder.createFile(blob);
      
      var file = DriveApp.getFilesByName(FileName)
      sendInvoice(Investor, EMail, invoiceNo, date, file);
      
      // end script here
      
      // array increment
      i = i + 1;
      
    }

  }
}
