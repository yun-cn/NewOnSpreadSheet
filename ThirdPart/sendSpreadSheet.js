function myFunction() {
  var ss_copyFrom = SpreadsheetApp.getActiveSpreadsheet();
  var sheet_copyFrom = ss_copyFrom.getSheetByName('シート1');

  var url = "https://docs.google.com/spreadsheets/d/1t#gid=0";
  var ss_copyTo = SpreadsheetApp.openByUrl(url);
  var nameMonth = getSheetName();
  var sheet_copyTo = ss_copyTo.getSheetByName(nameMonth + '月');  //指定シートに保存

  var allRange = sheet_copyFrom.getDataRange();
  var getAllDate = allRange.getValues();
 // var row = sheet_copyTo.getLastRow()+1;

  try {
    var row = sheet_copyTo.getLastRow()+1;
    for(var i = 0; i < getAllDate.length; i++){
      var value = [];
      value.push(getAllDate[i]);
      Logger.log(value);
      sheet_copyTo.getRange("A" + row + ":H" + row).setValues(value);
      row ++;
    }
    deleteDate();
  }
  catch(e){
    sEmail(e);
  }
}

//Get Target Sheet Name
function getSheetName(){
  var today = new Date();
  var yesterday = new Date(today.getTime() - (24 * 60 * 60 * 1000));
  var yearMonth = Utilities.formatDate(yesterday, Session.getScriptTimeZone(), 'yyyy/MM');
  Logger.log(yearMonth);
  return yearMonth;
}

function sEmail(body){
  MailApp.sendEmail("zhang@familiar-link.com", "Agodaキャンセルエラー", "\r\nMessage: " + body.message);
}

function deleteDate(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var first = ss.getSheetByName('シート1');
  first.clear();
}
