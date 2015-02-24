var SPREADSHEET_KEY = '1_dRC_kjhlkdfjshlskjdf23987432hejhkja1232111';
var NOTIFY_EMAIL = 'user1@email.com';
var CC_EMAIL = 'user2@email.com';

// starting point in this script:
// shows the index.html page and executes associated script
function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}


// This function is called after all information is submitted by the client.
// The form parameters is a dictionary that contains fields used in the index.html file.
// Field names must match column names in the spreadsheet. 
function addRequest(form) {

  try {

    // open the spreadsheet
    var spreadsheet = SpreadsheetApp.openById(SPREADSHEET_KEY);
    var sheet = spreadsheet.getSheets()[0];

    // get a mutex to update the spreadsheet
    var lock = LockService.getPublicLock();
    lock.waitLock(30000);
    
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]; //read headers
    var nextRow = sheet.getLastRow(); // get next row
    var cell = sheet.getRange('a1');
    var col = 0;
    
    // loop through the headers and if a parameter name matches the header name insert the value
    for (i in headers){ 
      if (headers[i] == "timestamp"){
         val = new Date();
      } 
      else {
        val = form[headers[i]]; 
      }
      cell.offset(nextRow, col).setValue(val);
      col++;
    }
  
    // flush all changes to the spreadsheet
    SpreadsheetApp.flush();
   
    // release lock
    lock.releaseLock(); 

    // send notification
    newRequestNotify(form, nextRow, spreadsheet.getUrl());
  
  } catch (error) {
    errorNotify (error);
    throw (error);
  }
}
    

// send notification email when a new request has been added to spreadsheet.
function newRequestNotify (form, row, spreadsheetUrl) {

  var bodyStr='';
  
  // prepare body text
  row+=1; // row is 0 based, print 1-based  
  bodyStr += "New download requested by: "+form["first-name"] + " " + form["last-name"] + "<br>";
  bodyStr += "Company: "+form["company"] + "<br>";
  bodyStr += "More information in row " + row+ " of this " + "spreadsheet".link(spreadsheetUrl) + "<br>"; 

  // send email
  MailApp.sendEmail({
    noReply:true,
    to: NOTIFY_EMAIL,
    cc: CC_EMAIL,
    subject: "New download request received",
    htmlBody: bodyStr,
  });
}


// report an error in script
function errorNotify (error) {
  MailApp.sendEmail( {
     noReply:true,
     to: NOTIFY_EMAIL,
     subject: "Script error",
     htmlBody: "Error: " + error,
  });
}