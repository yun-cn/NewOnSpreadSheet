/**
 * ラベルが付いた未読のメール(スレッド)を探して返す
 * @return GmailThread[]
 */

function getMail() {
  var label = "_timer+gmail-✔agoda予約完了 ";
  var start = 0;
  var max = 500;
  var startDate = yesterday();
  return GmailApp.search('label:' + label + "after:" + startDate  , start, max);

}

/**
 * メール本文を整形してスプレッドシートに保存するためのオブジェクトを返す
 * @return Object
 */
function getDatabyMailBody(body) {
  Logger.log(body);
  var bookingID = body.match(/予約 ID.*/);
  var reservationInfo = body.match(/FMC.*/);
  var propertyID = body.match(/Property ID.*/);
  var firstName = body.match(/First Name.*/);
  var lastName = body.match(/Last Name.*/);
  var checkIn = body.match(/チェックイン.*/);
  var checkOut = body.match(/チェックアウト.*/);
  var roomInfo = body.match(/Apartment.*/);
  var payout = body.match(/JPY.*/);
  var customerNotes = body.match(/顧客情報.*/);
   return {
     bookingID: bookingID,
     reservationInfo: reservationInfo,
     propertyID: propertyID,
     firstName: firstName,
     lastName: lastName,
     checkIn: checkIn,
     checkOut: checkOut,
     roomInfo: roomInfo,
     payout: payout,
     customerNotes: customerNotes
  };
}

/**
 * gmailを取得してスプレッドシートに保存する
 */
function onSaveMailToSheet() {
  // データを保存するシートの名前
  var sheetName = currentSheetName();
  var ss = SpreadsheetApp.getActive().getSheetByName( sheetName );
  var row = ss.getLastRow() + 1;
  var threads = getMail();
  for( var i in threads ) {
    Logger.log( "threads.lenth"+ threads.length );
    var thread = threads[i];
    var msgs = thread.getMessages();
    for( var j in msgs ) {
      Logger.log(j);
      var msg = msgs[j];
      var date = msg.getDate();
       if (dateToDay(date) == ifSecondDay()){
        var d = getDatabyMailBody( msg.getPlainBody() );
        var values = [
          [date, d.bookingID, d.firstName, d.lastName, d.checkIn, d.checkOut, d.roomInfo, d.payout, d.reservationInfo, d.propertyID, d.customerNotes]
        ];
        ss.getRange( "A" + row +":K" + row ).setValues( values );
        row++;
      }
    }
    Utilities.sleep(10000);
  }
}

function yesterday(){
  var today = new Date();
  Logger.log(today.getTime());
  var yesterday = new Date(today.getTime() - (24 * 60 * 60 * 1000)*2);
  var startDate = Utilities.formatDate(yesterday, Session.getScriptTimeZone(), 'yyyy/MM/dd');
  return startDate;
}

function today(){
  var today = new Date();
  var endDate = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy/MM/dd');
  return endDate;
}

function currentSheetName(){
   sheetName = getSheetName();
   var ss = SpreadsheetApp.getActiveSpreadsheet();
  if(ifSecondDay() == 01){
    ss.insertSheet(sheetName + '月');
    var sheet = ss.getActiveSheet();
    sheet.appendRow(["受信時間","予約ID","宿泊者氏名（名）","宿泊者氏名（姓）","チェックイン時間","チェックアウト時間","部屋情報","支払い料金","予約情報","Property ID","顧客情"]);
    return ss.getSheetName();
  }else{
    return ss.getActiveSheet().getSheetName();
  }
}

function ifSecondDay(){
    var today = new Date();
    var secondDay = new Date(today.getTime() - (24 * 60 * 60 * 1000));
    var day = Utilities.formatDate(secondDay, Session.getScriptTimeZone(), 'dd');
    return day;
}

//Get Target Sheet Name
function getSheetName(){
  var today = new Date();
  var yesterday = new Date(today.getTime() - (24 * 60 * 60 * 1000));
  var yearMonth = Utilities.formatDate(yesterday, Session.getScriptTimeZone(), 'yyyy/MM');
  return yearMonth;
}


function dateToDay(date){
  var date = new Date(date);
  var day = Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd');
  return day;
}
