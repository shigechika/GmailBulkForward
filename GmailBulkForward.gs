//   A		B		C
// 1 Name	BulkForwardTo	GmailSearchOperator
// 2 ComSys	ml@example.Com	"newer_than:1d from:(you@example.jp) to:(me@example.jp) subject:[vuls] (c-ap OR c-db) -High:0"
// 3 NetSys	ml@example.Net	"newer_than:1d from:(you@example.jp) to:(me@example.jp) subject:[vuls] (n-ap OR n-db) -High:0"

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
    if ( !subject || !recipient || !search ) {
      continue; // empty
    }

    var threads = GmailApp.search(search, start, max);
    Logger.log("threads.length:" + threads.length);
    if ( !threads.length ) {
      continue; // not found
    }
    
    var arryBlob = [];
    var summary = [];
    
    for (var t in threads) {
      var thread = threads[t];
      var msgs = thread.getMessages();           
      for(var m in msgs){
        var msg = msgs[m];
	var name = msg.getSubject();
        arryBlob.push(Utilities.newBlob(msg.getPlainBody()).setName(name));        
      }
      summary.push(name);
    }
    arryBlob.sort(function (a, b) {
        return a.getName() > b.getName() ? 1 : -1;
    });    
    GmailApp.sendEmail(recipient,
      subject,
      "total:" + summary.length "\n" + summary.sort().join('\n') + "\n-- GmailBulkForward\n" , // body
      { noReply: true, attachments: arryBlob });
    Utilities.sleep(1000);
  }
}
