
const interviewHeaderFormat = {
  "headerRow":1,
  "nextRowColor":"#ffffff",
  "category": {"columnWidth":200, "colNum":1, "name":"Category (related to RQs)"},
  "finalCode":{"columnWidth": 250, "colNum":2, "name":"Final Code (for Codebook)"},
  "originalCode":{"columnWidth": 250, "colNum":3, "name":"Initial Code (Labels for Quote or Memo)"},
  "quote":{"columnWidth":350, "colNum":4, "name":"Quote or Memo"},
  "lastRan":{"columnWidth":200, "colNum":5, "name":"Script Last Ran On"},
  "dataStartingRow":2,
  "firstCol":1,
  "lastCol":5
}

// Document ID currently set to testing doc (but it can be changed to anything)
var DOCUMENT_ID = "1Z8ivyILV4ZpQyJcDAlYqp1IVKZGobFA_K7Ww0fiXAyA"; // ID of interview transcript
const SPREADSHEET_ID = "1CEnjuR3bv25p7g3yBaXo1mj7_gJLy7FV0gz8bPCe9TI"; // ID of Qual Team Master Sheet
const SHEET_NAME = "P000002"; // Change to interview name


// Returns the row number of the first empty row in a column
function getLastRowInCol(sheet, dataStartingRow, col) {
  var i;
  for(i = dataStartingRow; i < sheet.getLastRow(); i++) {
    var cell = sheet.getRange(i, col);
    if(cell.getValue() == "") {
      return i;
    }
  }
  return i;
}


// Returns a list of the comments on the doc (in reverse chronological order)
function getComments(document_id, sheet_id, sheetName, headerFormat) {
  var pageToken = "";
  var all_comments = [];
  var currentSpreadsheet = SpreadsheetApp.openById(sheet_id);
  var currentInterviewSheet = currentSpreadsheet.getSheetByName(sheetName);
  var row = headerFormat["dataStartingRow"];
  var col = headerFormat["lastRan"]["colNum"];
  var firstTimeRunning = (currentInterviewSheet.getRange(row, col).getValue() == "");

  if(firstTimeRunning){ // running for the first time, so no timestamp is present yet
    var timeLastRan = new Date(1975, 1, 1, 0, 0, 0, 0);
  }
  else {
    var timeLastRan = new Date(currentInterviewSheet.getRange(row, col).getValue());
  }
  timeLastRan = timeLastRan.getTime();

  while (typeof pageToken !== "undefined") {
    var options = {
    'pageSize': 99,
    'pageToken': pageToken
    };
    var comments =  Drive.Comments.list(DOCUMENT_ID, options);
    var pageToken = comments.nextPageToken;
    
    for(var i = 0; i < comments.items.length; i++) {
      var comment = comments.items[i];
      var unixTime = Date.parse(comment.createdDate);
      var humanTime = new Date(unixTime); // human readable date
      var createdDate = new Date(humanTime);
      createdDate = createdDate.getTime(); // number of milliseconds
      if(comment.status == "open" && !comment.deleted && createdDate > timeLastRan) {
        all_comments.push(comment);
      }  
    }
  }
  return all_comments;
}


// Returns the name of the document that the script is running on
function getDocName(document_id) {
  return DocumentApp.openById(document_id).getName();
}


function convertColumnToLetter(columnNumber) {
  return String.fromCharCode(columnNumber+64);
}


function getA1Notation(row, col, numRows, numCols) {
  // This function only works for the first 26 columns
  var lastRow = row+numRows-1;
  var lastCol = col+numCols-1;
  var start = convertColumnToLetter(col)+row;
  var end = convertColumnToLetter(lastCol)+lastRow;
  return start+":"+end;
}


function formatHeader(sheet, headerFormat) {
  var headerRow = headerFormat["headerRow"];
  var firstCol = headerFormat["firstCol"];
  var lastCol = headerFormat["lastCol"];
  var nextRowColor = headerFormat["nextRowColor"];
  
  // Converting the column numbers into A1 notation to get range of cells where data will be
  dataRange = sheet.getRange(convertColumnToLetter(firstCol)+":"+convertColumnToLetter(lastCol));
  // Allowing text to wrap properly (instead of go into other cells)
  dataRange.setWrap(true);
  
  // Centering the header and making it bold
  headerRange = sheet.getRange(getA1Notation(headerRow, firstCol, 1, lastCol-firstCol+1));
  headerRange.setHorizontalAlignment("center");
  
   
  // Adding a colored row below the header (note: this could be #ffffff in which case nothing happens)
  coloredRow = sheet.getRange(getA1Notation(headerRow+1, firstCol, 1, lastCol-firstCol+1));
  coloredRow.setBackground(nextRowColor);
  
  var bold = SpreadsheetApp.newTextStyle().setBold(true).build();
  
  // Set the column width for each of the columns in the header
  for(header in headerFormat) {
    if(headerFormat[header]["name"] != null) {
      var headerCell = sheet.getRange(getA1Notation(headerRow, headerFormat[header]["colNum"], 1, 1));
      var headerString = headerFormat[header]["name"];
      var value;
      if(headerFormat[header]["extraText"] != null) {
        value = SpreadsheetApp.newRichTextValue()
                  .setText(headerFormat[header]["name"]+" — "+headerFormat[header]["extraText"])
                  .setTextStyle(0, headerFormat[header]["name"].length, bold)
                  .build();
      
      }
      else {
        value = SpreadsheetApp.newRichTextValue()
                  .setText(headerFormat[header]["name"])
                  .setTextStyle(0, headerFormat[header]["name"].length, bold)
                  .build();
      }
      headerCell.setRichTextValue(value);
      sheet.setColumnWidth(headerFormat[header]["colNum"], headerFormat[header]["columnWidth"]);
    }
  }
}


function addComments(document_id, sheet, headerFormat, comments) {
  var row = headerFormat["dataStartingRow"];
  var originalCodesCol = headerFormat["originalCode"]["colNum"];
  var quotesCol = headerFormat["quote"]["colNum"];
  var timestampCell = sheet.getRange(convertColumnToLetter(headerFormat["lastRan"]["colNum"])+headerFormat["dataStartingRow"]);
  var firstTimeRunningScript = (timestampCell.getValue() == "");
  var firstEmptyRow = (firstTimeRunningScript ? getLastRowInCol(sheet, row, originalCodesCol) : getLastRowInCol(sheet, row, originalCodesCol)+1);
  
  var fixedQuote;
  for(var i = comments.length-1; i >= 0; i--) { // Iterating backwards through comments to process them in chronological order
    if(comments[i].status == "open") {
      var commentCell = sheet.getRange(convertColumnToLetter(originalCodesCol)+firstEmptyRow);
      var quoteCell = sheet.getRange(convertColumnToLetter(quotesCol)+firstEmptyRow);
      var commentID = comments[i].commentId;
      var link = "https://docs.google.com/document/d/"+document_id+"/edit?disco="+commentID;
      
      
      // Sorting Memos
      // When the comment on the transcript is of the form "Memo: ... - memo text",
      // the "memo text" is in the quotes column and "[Memo] ..." is the original code names column
      if(comments[i].content.startsWith("Memo:") || comments[i].content.startsWith("memo:")){
         var entireMemo = comments[i].content.split(' - ');
         var memoTitle = entireMemo[0].split(': ');
         commentCell.setValue("["+memoTitle[0]+"] "+memoTitle[1]); // first part of memo comment
         quoteCell.setValue('=hyperlink("'+link+'","'+entireMemo[1]+'")'); // second part of memo comment
         firstEmptyRow++;
         row++;
         continue;
       }
      
      commentCell.setValue(comments[i].content);
      fixedQuote = comments[i].context.value. replace(/&#39;/g,"'"); // Changes an apostrophe's display from its HTML encoding (&#39;) to '
      quoteCell.setValue('=hyperlink("'+link+'","'+fixedQuote+'")'); // Adding the text associated with that comment to quote cell
      firstEmptyRow++;
      row++;
    }
  }
}


// Get timestamp for when createFirstCycleSheet() or generateSheet() is run
function getTimestamp(sheet_id, sheetName, headerFormat){  
  var row = headerFormat["dataStartingRow"];
  var col = headerFormat["lastRan"]["colNum"];
  var time = Utilities.formatDate(new Date(), "GMT-7", "MM/dd/yyyy, h:mm:ss a");
  
  var currentSpreadsheet = SpreadsheetApp.openById(sheet_id);
  var currentInterviewSheet = currentSpreadsheet.getSheetByName(sheetName);
  currentInterviewSheet.getRange(row, col).setValue(time);
}


// Translate a date from Google's Date format to Unix time to human readable time
function translateTime(googleTime){
  var unixTime = Date.parse(googleTime);
  var humanTime = new Date(unixTime);
  humanTime = humanTime.toLocaleString();
  return humanTime;
}


function generateSheet(doc_id, sheet_id, sheetName) {
  var spreadsheet = SpreadsheetApp.openById(sheet_id);
  var sheet;
  if((sheet = spreadsheet.getSheetByName(sheetName)) == null) {
    spreadsheet.insertSheet(sheetName, spreadsheet.getSheets().length);
    sheet = spreadsheet.getSheetByName(sheetName);
  }
  var sheet = spreadsheet.getSheetByName(sheetName);
  var comments = getComments(doc_id, sheet_id, sheetName, interviewHeaderFormat);
  
  formatHeader(sheet, interviewHeaderFormat);
  sheet.setFrozenRows(interviewHeaderFormat["headerRow"]);
  addComments(doc_id, sheet, interviewHeaderFormat, comments);
  getTimestamp(sheet_id, sheetName, interviewHeaderFormat);
}


function createFirstCycleSheet() {
  generateSheet(DOCUMENT_ID, SPREADSHEET_ID, SHEET_NAME);
}
