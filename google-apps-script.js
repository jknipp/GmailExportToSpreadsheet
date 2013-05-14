/*
       ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~    
     >> G M A I L : Export Emails by Label to Spreadsheet >>
       ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~   
   This script will take emails from a Label and import them to 
   a spreadsheet, listing emails individually in order of date sent.

   Jared Knipp - 4/30/2013
*/
var sheet = SpreadsheetApp.getActiveSheet();
var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
var LabelWithEmails = sheet.getRange(3, 2).getValue();

/** 
 * Helper for finding the first row to start dumping emails into.
 */
function getFirstRow() {
    var start = 6;
    return start;
}


function getFirstMsgId() {
    var label = GmailApp.getUserLabelByName(LabelWithEmails);
    var threads = label.getThreads(0, 1);
    var message = threads[0].getMessages()[0];
    var firstmessageId = message.getId();

    return firstmessageId;
}

/**
 * Populates a spreadsheet with emails from a label.  
 * Strips out the html and tries poorly to make each message of a conversation
 * a seperate row in the spreadsheet.
 */
function getEmails() {
    clearCanvas();
    var label = GmailApp.getUserLabelByName(LabelWithEmails);
    var threads = label.getThreads();

    var row = getFirstRow() + 1;
    var firstmessageId = getFirstMsgId();
    UserProperties.setProperty("firstmsgid", firstmessageId);
    //spreadsheet.toast("Loading emails..Please wait. It could take few seconds", "Status", -1);

    var messages = GmailApp.getMessagesForThreads(threads); //gets messages in 2D array

    for (var i = 0; i < threads.length; i++) {
        
  	try {
			var messages = threads[i].getMessages();
          
			for (var m = 0; m < messages.length; m++) {
                var msg = messages[m];
                var isForward = msg.getBody().search(/---------- Forwarded message/i) != -1;
              
                if(!isValidMessage(msg)) continue;

                sheet.getRange(row, 1).setValue(msg.getFrom());
                sheet.getRange(row, 2).setValue(msg.getTo() + ";" + msg.getCc() + ";" + msg.getBcc());
				sheet.getRange(row, 3).setValue(msg.getSubject());
                sheet.getRange(row, 4).setValue(msg.getDate());
		        
                if(!isForward) {
                    // Get only this messages body, ignore the previous chain
                    var body = msg.getBody();
                    var firstIndexOfThread = body.search(/gmail_quote/i); 
                  
                    if(firstIndexOfThread != -1) {
                        body =  body.substring(0, firstIndexOfThread);
                    }
                    
                    sheet.getRange(row, 5).setValue(getTextFromHtml(body));
                    
                
                } else {
                    // Use the whole body if its a forward
                    // jknipp - potentially hide forwarded messages
                    sheet.getRange(row, 5).setValue(getTextFromHtml(msg.getBody()));
                    // Use bright green
                    sheet.getRange(row, 1, 1, 6).setBackgroundRGB(0, 255, 0);
                }
              
				row++;
			}
		} catch (error) {
            Logger.log(error);
			spreadsheet.toast("Error Occured - please see the logs.", "Status", -1);
		}

		if (i == threads.length - 1) {
			spreadsheet.toast("Successfully loaded emails.", "Status", -1);
			//spreadsheet.toast("Now mark emails to be forwarded by changing the background color of the cells to green. Then select Forward->Forward selected emails", "Status", -1);
		}
	}
}


/**
 * Clear the canvas. 
 * TODO: Could be improved so it doesn't depend on a 
 */
function clearCanvas() {
    sheet.getRange("A8:F1500").clear();
}

/**
 * Setup Menu Item in Google Docs Spreadsheet
 */
function onOpen() {
    var menuEntries = [{
            name: "Load/Refresh Emails",
            functionName: "getEmails"
        }, {
            name: "Clear canvas",
            functionName: "clearCanvas"
        }
    ];
    spreadsheet.addMenu("Import Gmail", menuEntries);
}

/**
 * Strip out the html characters
 */
function getTextFromHtml(html) {
    return getTextFromNode(Xml.parse(html, true).getElement());
}

function getTextFromNode(x) {
    switch (x.toString()) {
    case 'XmlText':
        return x.toXmlString();
    case 'XmlElement':
        return x.getNodes().map(getTextFromNode).join('');
    default:
        return '';
    }
}

/**
 * Validation to see if we want to store this email to the document
 * This is my own personal filter, not necessarily useful for anyone else.
 */ 
function isValidMessage(msg) {
  // filter by date, before 2011
  //Logger.log(msg.getDate().getYear());
  if(msg.getDate().getYear() < 2011) {
    return false;
  } else if (msg.isInChats()) {
    return false;
  }

  return true;
}
