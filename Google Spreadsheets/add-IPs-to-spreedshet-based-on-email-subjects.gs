/*
*
*   Query GMail for a subject and add IP's in each subject to a spreadsheet
*   I made this script for getting a list of IPs to block in a VPS server,
*   receiving email notifications from WHM, in example
*
*/

var sheet = SpreadsheetApp.getActiveSheet();
var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

// Cell we get the subject from
var targetSubject = sheet.getRange(1, 3).getCell(1, 1).getValue();
// Last row added in the last execution
var lastRow = parseInt(sheet.getRange(2, 3).getCell(1, 1).getValue());


/**
 * Cleans inbox in GMail removing emails in a specified query
 *
 * @param {filterString} String
 *   Filter to find emails to remove
 *
 * @return
 *   null
 */
function cleanEmails(filter) {
  if (!filter) return;

  // Remove other emails
  var threadToDelete = GmailApp.search(filter);
  for (var i = 0; i < threadToDelete.length; i++) {
    var messagesToDelete=threadToDelete[i].getMessages();
    for (var m=messagesToDelete.length-1; m>=0; m--) {
      messagesToDelete[m].moveToTrash();
    }
  }

}


/**
 * Adds to the first column in the active page, IPs detected in the subject
 * of the email.  It must begin with a fixed string, specified in the cell 1C
 *
 *
 * @return
 *   null
 */
function getEmails() {
  
  // Clean inbox
  cleanEmails("label:expertiseit subject:one or more immutable files are preventing");
  
  // Let's go
  var threads = GmailApp.search("label:my-label subject:"+targetSubject);

  var actualRow = lastRow + 1;
  for (var i = 0; i < threads.length; i++) {
    var messages=threads[i].getMessages();
    
    for (var m = messages.length-1; m>=0; m--) {
      var subject = messages[m].getSubject();
      // TODO: get IP by regExp in the subject
      // Right now, it gets the IP in the string position 45 (expected subject "Large Number of Failed Login Attempts from IP xxx.xxx.xxx.xxx")
      var targetIP = subject.substring(45);
      
      // Check if IP has been previously added. In sucj a case, don't add again
      var exists = false;
      for (var j=firstRow; j<actualRow; j++) {
        var checkedCellValue = sheet.getRange(j, 1).getCell(1, 1).getValue();
        if (checkedCellValue==targetIP) {
          exists = true;
          j = actualRow + 1;
        }
      }
      
      if (!exists) {
        sheet.getRange(actualRow, 1).setValue(targetIP);
        actualRow++;
      }
      // Remove email message (move to trash)
      messages[m].moveToTrash();
      
      if ((actualRow % 20)===0) {
        // Sleep for half a second after adding 20 rows.
        // Otherwise, GMail block execution due to too much time proccessing
        Utilities.sleep(1000);
      }
    }
  }

  // Set next row to add IP for the next execution
  sheet.getRange(2, 3).getCell(1, 1).setValue(actualRow-1);

  // Sort the column to readability
  sheet.sort(1, true);
}


/** @constructor */
function onOpen() {
  var menuEntries = [ {name: "Load IPs to filter from Emails", functionName: "getEmails"} ];
  spreadsheet.addMenu("Email", menuEntries);
}