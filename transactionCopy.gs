function yesterdaySales() {
  var app = SpreadsheetApp;
  var sheet = app.getActiveSpreadsheet().getSheetByName("allDays");
  var rows = sheet.getLastRow()
  var PMb = '=if(E20<>"",E20*100,"")'
  var ydate = '=today()-1'
  var rev = '=VLOOKUP(day(indirect("r[0]c[-1]",false)),BetterSummary!$A:$L,3,false)'
  var exp = '=VLOOKUP(day(indirect("r[0]c[-2]",false)),BetterSummary!$A:$L,4,false)'
  var profit = '=VLOOKUP(day(indirect("r[0]c[-3]",false)),BetterSummary!$A:$L,10,false)'
  var pm = '=VLOOKUP(day(indirect("r[0]c[-4]",false)),BetterSummary!$A:$L,11,false)'
  var pmb = '=if(indirect("r[0]c[-1]",false)<>"",indirect("r[0]c[-1]",false)*100,"")'
  var lookup = '=int(if(indirect("r[0]c[3]",false)<>"",concatenate(month(indirect("r[0]c[1]",false)),year(indirect("r[0]c[1]",false))),"Need-Imput"))'
  var h = '=sumif(A:A,indirect("r[0]c[-7]",false),E:E)'
  var i = '=countif(A:A,indirect("r[0]c[-8]",false))'
  var j = '=round(sumif(A:A,indirect("r[0]c[-9]",false),F:F)/indirect("r[0]c[-1]",false),4)'
  var k = '=if(indirect("r[0]c[-1]",false)<>"",indirect("r[0]c[-1]",false)*100,"")'
  var l = '=IF(indirect("r[0]c[-1]",false)<>indirect("r[-1]c[-1]",false),indirect("r[0]c[-4]",false),"")'
  var m = '=SUM($L$1:L'+ rows +')'
  var n = '=VLOOKUP(day(indirect("r[0]c[-12]",false)),BetterSummary!$A:$L,5,false)'
  var o = '=VLOOKUP(day(indirect("r[0]c[-13]",false)),BetterSummary!$A:$L,6,false)'
  var p = '=VLOOKUP(day(indirect("r[0]c[-14]",false)),BetterSummary!$A:$L,7,false)'
  var q = '=VLOOKUP(day(indirect("r[0]c[-15]",false)),BetterSummary!$A:$L,8,false)'
  
  
  
  
  
  
  var slookup = sheet.getRange(rows, 1).setValue(lookup);
  var sdate = sheet.getRange(rows, 2).setValue(ydate);
  var srev = sheet.getRange(rows, 3).setValue(rev);
  var sexp = sheet.getRange(rows, 4).setValue(exp);
  var sprofit = sheet.getRange(rows, 5).setValue(profit);
  var spm = sheet.getRange(rows, 6).setValue(pm);
  var spmb = sheet.getRange(rows, 7).setValue(pmb);
  var sh = sheet.getRange(rows, 8).setValue(h);
  var si = sheet.getRange(rows, 9).setValue(i);
  var sj = sheet.getRange(rows, 10).setValue(j);
  var sk = sheet.getRange(rows, 11).setValue(k);
  var mm = sheet.getRange(rows, 13).setValue(m);
  var nn = sheet.getRange(rows, 14).setValue(n);
  var oo = sheet.getRange(rows, 15).setValue(o);
  var pp = sheet.getRange(rows, 16).setValue(p);
  var qq = sheet.getRange(rows, 17).setValue(q);
  
  var glookup = sheet.getRange(rows, 1).getValue();
  var gdate = sheet.getRange(rows, 2).getValue();
  var grev = sheet.getRange(rows, 3).getValue();
  var gexp = sheet.getRange(rows, 4).getValue();
  var gprofit = sheet.getRange(rows, 5).getValue();
  var gpm = sheet.getRange(rows, 6).getValue();
  var gpmb = sheet.getRange(rows, 7).getValue();
  var mmm = sheet.getRange(rows, 13).getValue();
  var nnn = sheet.getRange(rows, 14).getValue();
  var ooo = sheet.getRange(rows, 15).getValue();
  var ppp = sheet.getRange(rows, 16).getValue();
  var qqq = sheet.getRange(rows, 17).getValue();
  
  var fslookup = sheet.getRange(rows, 1).setValue(glookup);
  var fsdate = sheet.getRange(rows, 2).setValue(gdate);
  var fsrev = sheet.getRange(rows, 3).setValue(grev);
  var fsexp = sheet.getRange(rows, 4).setValue(gexp);
  var fsprofit = sheet.getRange(rows, 5).setValue(gprofit);
  var fspm = sheet.getRange(rows, 6).setValue(gpm);
  var fspmb = sheet.getRange(rows, 7).setValue(gpmb);
  
  var nnnn = sheet.getRange(rows, 14).setValue(nnn);
  var oooo = sheet.getRange(rows, 15).setValue(ooo);
  var pppp = sheet.getRange(rows, 16).setValue(ppp);
  var qqqq = sheet.getRange(rows, 17).setValue(qqq);
}

function todaySales(){
  var app = SpreadsheetApp;
  var sheet = app.getActiveSpreadsheet().getSheetByName("allDays");
  var rows = sheet.getLastRow()+1
  var PMb = '=if(E20<>"",E20*100,"")'
  var ydate = '=today()'
  var rev = '=VLOOKUP(day(indirect("r[0]c[-1]",false)),BetterSummary!$A:$L,3,false)'
  var exp = '=VLOOKUP(day(indirect("r[0]c[-2]",false)),BetterSummary!$A:$L,4,false)'
  var profit = '=VLOOKUP(day(indirect("r[0]c[-3]",false)),BetterSummary!$A:$L,10,false)'
  var pm = '=VLOOKUP(day(indirect("r[0]c[-4]",false)),BetterSummary!$A:$L,11,false)'
  var pmb = '=if(indirect("r[0]c[-1]",false)<>"",indirect("r[0]c[-1]",false)*100,"")'
  var lookup = '=int(if(indirect("r[0]c[3]",false)<>"",concatenate(month(indirect("r[0]c[1]",false)),year(indirect("r[0]c[1]",false))),"Need-Imput"))'  
  var h = '=sumif(A:A,indirect("r[0]c[-7]",false),E:E)'
  var i = '=countif(A:A,indirect("r[0]c[-8]",false))'
  var j = '=round(sumif(A:A,indirect("r[0]c[-9]",false),F:F)/indirect("r[0]c[-1]",false),4)'
  var k = '=if(indirect("r[0]c[-1]",false)<>"",indirect("r[0]c[-1]",false)*100,"")'
  var l = '=IF(indirect("r[0]c[-1]",false)<>indirect("r[-1]c[-1]",false),indirect("r[0]c[-4]",false),"")'
  var m = '=SUM($L$1:L'+ rows +')'
  var n = '=VLOOKUP(day(indirect("r[0]c[-12]",false)),BetterSummary!$A:$L,5,false)'
  var o = '=VLOOKUP(day(indirect("r[0]c[-13]",false)),BetterSummary!$A:$L,6,false)'
  var p = '=VLOOKUP(day(indirect("r[0]c[-14]",false)),BetterSummary!$A:$L,7,false)'
  var q = '=VLOOKUP(day(indirect("r[0]c[-15]",false)),BetterSummary!$A:$L,8,false)'
  
  var slookup = sheet.getRange(rows, 1).setValue(lookup);
  var sdate = sheet.getRange(rows, 2).setValue(ydate);
  var srev = sheet.getRange(rows, 3).setValue(rev);
  var sexp = sheet.getRange(rows, 4).setValue(exp);
  var sprofit = sheet.getRange(rows, 5).setValue(profit);
  var spm = sheet.getRange(rows, 6).setValue(pm);
  var spmb = sheet.getRange(rows, 7).setValue(pmb);
  var sh = sheet.getRange(rows, 8).setValue(h);
  var si = sheet.getRange(rows, 9).setValue(i);
  var sj = sheet.getRange(rows, 10).setValue(j);
  var sk = sheet.getRange(rows, 11).setValue(k);
  var ll = sheet.getRange(rows, 12).setValue(l);
  var mm = sheet.getRange(rows, 13).setValue(m);
  var nn = sheet.getRange(rows, 14).setValue(n);
  var oo = sheet.getRange(rows, 15).setValue(o);
  var pp = sheet.getRange(rows, 16).setValue(p);
  var qq = sheet.getRange(rows, 17).setValue(q);
}
