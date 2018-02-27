//   A		B		C
// 1 Name	BulkForwardTo	GmailSearchOperator
// 2 ComSrvs	ml@example.com	"newer_than:1d from:(from@example.jp) to:(me@example.jp) subject:[vuls] (com1 OR com2) -{Hight:0}"
// 3 NnetSrvs	ml@example.net	"newer_than:1d from:(from@example.jp) to:(me@example.jp) subject:[vuls] (net1 OR net2) -{Hight:0}"

function GmailBulkForward(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var start = 0;
  var max = 100;

  var lastRow = sheet.getLastRow();
  Logger.log("lastRow:" + lastRow);

  for ( var r = 2; r <= lastRow ; r++ ) {
    var subject = sheet.getRange(r, 1).getValue();
    Logger.log("subject:" + subject);
    var recipient = sheet.getRange(r, 2).getValue();
    Logger.log("recipient:" + recipient);
    var search = sheet.getRange(r, 3).getValue();
    Logger.log("search:" + search);
    var threads = GmailApp.search(search, start, max);
    Logger.log("threads.length:" + threads.length);
    
    if ( ! threads.length ) {
      continue; // not found
    }
    
    var arryBlob = [];
    
    for (var t in threads) {
      var thread = threads[t];
      var msgs = thread.getMessages();           
      for(var m in msgs){
        var msg = msgs[m];
        arryBlob.push( Utilities.newBlob(msg.getPlainBody()).setName(msg.getSubject()));        
      }
    }
    
    var date = new Date();
    GmailApp.sendEmail(recipient,
      subject,
      Utilities.formatDate(date, "Asia/Tokyo", "yyyy-MM-dd'T'HH:mm:ss.SSSZ") , // body
      { noReply: true, attachments: arryBlob });
    Utilities.sleep(1000);
  }
}
