/**
 * ラベルが付いた未読のメール(スレッド)を探して返す
 * @return GmailThread[]
 */

function getMail() {
  var label = "agoda-cancellation ";
  var start = 0;
  var max = 500;
  var startDate = yesterday();
  return GmailApp.search('label:' + label+ 'after:' + startDate, start, max);
}

/**
 * メール本文を整形してスプレッドシートに保存するためのオブジェクトを返す
 * @return Object
 */
function getDatabyMailBody( body ) {
  var bookingID = body.match(/予約 ID.*/);
  var firstName = body.match(/宿泊者氏名（名）.*/);
  var lastName = body.match(/宿泊者氏名（姓）.*/);
  var checkIn = body.match(/到着.*/);
  var checkOut = body.match(/出発.*/);
  var roomInfo = body.match(/ホテル :.*/);
  var cancellationFee = body.match(/キャンセル料金:.*/);

  return {
    bookingID: bookingID,
    firstName: firstName,
    lastName: lastName,
    checkIn: checkIn,
    checkOut: checkOut,
    roomInfo: roomInfo,
    cancellationFee: cancellationFee

  };
 }

/**
 * gmailを取得してスプレッドシートに保存する
 */
function onSaveMailToSheet() {
  // データを保存するシートの名前
  var sheetName = 'シート1';
  var ss = SpreadsheetApp.getActive().getSheetByName( sheetName );
  var row = ss.getLastRow() + 1;
  var threads = getMail();

  for( var i in threads ) {
    var thread = threads[i];
    var msgs = thread.getMessages();
    for( var j in msgs ) {
      var msg = msgs[j];
      var date = msg.getDate();
      if (dateToDay(date) == getDay()){
        var d = getDatabyMailBody( msg.getPlainBody() );
        var values = [
          [date, d.bookingID,d.firstName, d.lastName, d.checkIn, d.checkOut, d.roomInfo, d.cancellationFee]
        ];
        ss.getRange("A" + row +":H" + row).setValues(values);
        row++;
      }
    }
    Utilities.sleep(10000);
  }
}

function yesterday(){
  var today = new Date();
  var yesterday = new Date(today.getTime() - (24 * 60 * 60 * 1000)*2);
  var startDate = Utilities.formatDate(yesterday, Session.getScriptTimeZone(), 'yyyy/MM/dd');
  return startDate;
}

function getDay(){
    var today = new Date();
    var secondDay = new Date(today.getTime() - (24 * 60 * 60 * 1000));
    var day = Utilities.formatDate(secondDay, Session.getScriptTimeZone(), 'dd');
    return day;
}

function dateToDay(date){
  var date = new Date(date);
  var day = Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd');
  return day;
}
