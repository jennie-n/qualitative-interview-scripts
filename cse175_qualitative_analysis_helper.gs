// Opening the master sheet manually
const spreadSheetID = "1qmZtTc5tGdW8IiMxgGiuJSnUXXugZaThFblhxN59eUE";
const masterSheet = SpreadsheetApp.openById(spreadSheetID);
// Open groups sheet manually
const assignedGroupsSheetID = "1pJjc-TPEC0XE82NscHaNywcO6awRFYv3GQ2AGRwvxeY";
const assignedGroupsSheet = SpreadsheetApp.openById(assignedGroupsSheetID);

const NUM_TEAMS = 53;
var teachingAssistants = []; // emails of TAs who will be given editing access to all protected sheets

// Creating a JSON object to represent the header format of the team sheets
const teamHeaderFormat = {
  "headerRow":1,
  "nextRowColor":"#ffffff",
  "question": {"columnWidth":350, "colNum":1, "name":"Categories Framed as Research Questions"},
  "observation":{"columnWidth": 410, "colNum":2, "name":"Observations and Memos"},
  "quote":{"columnWidth":450, "colNum":3, "name":"Quotes"},
  "dataStartingRow":2,
  "firstCol":1,
  "lastCol":3
}

// Creating a JSON object to represent the header format of the main questions sheet
const researchQuestionsHeaderFormat = {
  "headerRow":1,
  "nextRowColor":"#9FC5E8",
  "question": {"columnWidth":475, "colNum":1, "name":"Categories Framed as Research Questions"},
  // "selected":{"columnWidth":75, "colNum":2, "name":"Selected?"},
  // "category":{"columnWidth":125, "colNum":3, "name":"RQ Tab"},
  "dataStartingRow":3,
  "firstCol":1,
  "lastCol":1
}

// Creating a JSON object to represent the header format of the extra questions sheets
const individualQuestionsHeaderFormat = {
  "headerRow":2,
  "nextRowColor":"#ffffff",
  "observation": {"columnWidth":500, "colNum":1, "name":"Observation"},
  "quote":{"columnWidth":600, "colNum":2, "name":"Quote"},
  "ideasTab":{"columnWidth":100, "colNum":3, "name":"Ideas Tab?"},
  "dataStartingRow":3,
  "firstCol":1,
  "lastCol":3
}

const ideasFirstHeader = {
  "headerRow":2,
  "nextRowColor":"#ffffff",
  "pointOfViews": {"columnWidth":400, "colNum":1, "name":"Point of Views", "extraText":"defines a problem space"},
  "howMightWe":{"columnWidth":400, "colNum":2, "name":"How Might We", "extraText":"defines a solution space"},
  "productIdeas":{"columnWidth":400, "colNum":3, "name":"Product Ideas", "extraText":"a solution"},
  "dataStartingRow":3,
  "firstCol":1,
  "lastCol":3
}

const ideasSecondHeader = {
  "headerRow":13,
  "nextRowColor":"#ffffff",
  "subcategories": {"columnWidth":400, "colNum":1, "name":"Subcategories"},
  "observations":{"columnWidth":400, "colNum":2, "name":"Observations and Memos"},
  "quotes":{"columnWidth":400, "colNum":3, "name":"Quotes"},
  "dataStartingRow":14,
  "firstCol":1,
  "lastCol":3
}

const assignmentsHeaderFormat = {
  "headerRow":1,
  "nextRowColor":"#ffffff",
  "questions": {"columnWidth":475, "colNum":2, "name":"Selected Categories Framed as Research Questions"},
  "rqTab": {"columnWidth": 200, "colNum":5, "name": "RQ Tab"},
  "selected": {"columnWidth": 100, "colNum":6, "name": "Selected?"},
  "initialLabels":{"columnWidth": 100, "colNum":4, "name": "Initial Labels / Notes?"},
  "dataStartingRow":2,
  "firstCol":2,
  "lastCol":4
}

var questionsObj = {};

// HELPER FUNCTIONS-----------------------------------------------------------------------

// Returns the row number of the first empty row in a column
function getLastRowInCol(sheet, dataStartingRow, col) {
  var i;
  for(i = dataStartingRow; i < sheet.getLastRow(); i++) {
    var cell = sheet.getRange(i, col);
    if(cell.getValue() == "") {
      return i-1;
    }
  }
  return i-1;

}


// Creates a Google Sheets formula for aggregating all of the questions from
// all of the teams into a single column
function createQuestionAggregationFormula(numTeams, dataRange) {
  // Using Google Sheets formula that David made to collect each
  // question from the Team sheets and put it into the Research Questions tab
  var questionDataRange = getTeamDataRange(numTeams, dataRange);
  var formula = "=SORT(FILTER(UNIQUE({";

  formula += questionDataRange;
  formula += ("}), NOT(ISBLANK(UNIQUE({");
  formula += questionDataRange;
  formula += ("}))), NOT(REGEXMATCH(UNIQUE({");
  formula += questionDataRange;
  formula += ("}), \"NONE\")=true)), 1, true)");
  return formula;
}


// Creates a Google Sheets formula for aggregating all of the observations and
// quotes associated with the questions associated with a specific question tab
function createObservationAggregationFormula(spreadsheet, questionTabName, numTeams, firstRange, secondRange) {
  var firstTeamRange = getTeamDataRange(numTeams, firstRange);
  var secondTeamRange = getTeamDataRange(numTeams, secondRange);
  var questions = questionsObj[questionTabName];

  // Using Google Sheets formula that David made to find each question
  // associated with the question tab and then get the observations
  // and quotes associated with all of the team sheets
  var formula = "=SORT(FILTER({";
  formula += firstTeamRange;
  formula += "}, REGEXMATCH({";
  formula += secondTeamRange;
  formula += "}, \"^(";
  Logger.log("This is the number of questions: " + questions.length);
  for(var i = 0; i < questions.length; i++) {
    var currentQuestion = questions[i];
    // Replace any quotation mark in the question with two quotation marks
    // This is needed for a quotation mark escape in the Google Sheets formula
    var modifiedQuestion = currentQuestion.replace(/"/g, "\"\"");
    formula += modifiedQuestion;
    if(i < questions.length-1) {
      formula += "|";
    }
  }
  formula += ")\")=true), 1, true)";
  return formula;
}


// Returns a string containing the A1 notation for all of the questions
// i.e. 'Team 1'!A2:A999;'Team 2'!A2:A999; ... ; 'Team 50'!A2:A999
function getTeamDataRange(numTeams, dataRange) {
  var rangeNotation = "";
  for(var i = 1; i <= numTeams; i++) {

    rangeNotation += ("\'Team " + i + "\'!" + dataRange);
    if(i < numTeams) {
      rangeNotation += ";";
    }
  }
  Logger.log("BUILT TEAM RANGE");
  return rangeNotation;
}


// Converts a column number to the appropriate A1 notation letter (by translating
// the ASCII value by 64). Note that this only works for column numbers in the range [1, 26]
function convertColumnToLetter(columnNumber) {
  return String.fromCharCode(columnNumber+64);
}


// Converts a range into A1 notation to use for App Script functions.
// Note that it only works for the first 26 columns (since each
// column is represented by a letter)
function getA1Notation(row, col, numRows, numCols) {
  var lastRow = row+numRows-1;
  var lastCol = col+numCols-1;
  var start = convertColumnToLetter(col)+row;
  var end = convertColumnToLetter(lastCol)+lastRow;
  return start+":"+end;
}


// Creates a dropdown menu in the targetRange of the targetSheet using the data
// contained in the dataRange of the dataSheet. Note that each "range" is specified
// using A1 notation string (e.g. "A1:A50" = rows 1 through 50 in col 1)
function createDropdown(targetSheet, targetRange, dataSheet, dataRange) {
  var dropdownOptions = dataSheet.getRange(dataRange);
  var validationRule = SpreadsheetApp.newDataValidation().requireValueInRange(dropdownOptions);
  targetSheet.getRange(targetRange).setDataValidation(validationRule);
}


function parseRelatedQuestions(spreadsheet, questionTabName) {
  var assignmentsSheet = spreadsheet.getSheetByName("Assignments");
  var questionCol = assignmentsHeaderFormat["questions"]["colNum"];
  var categoryCol = assignmentsHeaderFormat["rqTab"]["colNum"];
  var startingRow = assignmentsHeaderFormat["dataStartingRow"];
  var checkmarkCol = assignmentsHeaderFormat["selected"]["colNum"];
  var dataStartingRow = assignmentsHeaderFormat["dataStartingRow"];
  var questionList = [];

  for(var i = 1; i <= getLastRowInCol(assignmentsSheet, dataStartingRow, categoryCol); i++) {
    var currentCategory = assignmentsSheet.getRange(i, categoryCol).getValue();
    var currentQuestion = assignmentsSheet.getRange(i, questionCol).getValue();
    var checkmarked = assignmentsSheet.getRange(i, checkmarkCol).getValue().toString().toLowerCase()=="yes";
    if(checkmarked && currentCategory.includes(questionTabName)) {
      questionList.push(currentQuestion);
    }
  }
  return questionList;
}


// Enables edit permissions to the file for all students
function enableEdit() {
  var sheet = assignedGroupsSheet.getSheets()[0];
  var dataStartingRow = 2;
  var emailCol = 1;
  var lastEmailRow = getLastRowInCol(sheet, dataStartingRow, emailCol);
  var emails = sheet.getSheetValues(dataStartingRow, emailCol, lastEmailRow, 1);
  masterSheet.addEditors(emails);
  masterSheet.addEditors(teachingAssistants);
}

// IMPORTANT FUNCTIONS----------------------------------------------------------------------------------

// Creates & returns a 2D array with the grouped emails
function createGroups(numTeams){
  //var numTeams=53;
  var group = [];
  var i=0;
  // create 2D array
  for (i=0;i<numTeams;i++) {
     group[i] = [];
  }

  var sheet = assignedGroupsSheet.getSheets()[0];
  var dataStartingRow = 2;
  var emailCol = 1;
  var lastEmailRow = getLastRowInCol(sheet, dataStartingRow, emailCol);
  var groupCol = 3;
  var lastGroupRow = getLastRowInCol(sheet, dataStartingRow, groupCol);

  // emails
  var emails = sheet.getSheetValues(dataStartingRow, emailCol, lastEmailRow, 1);

  // group numbers
  var groupNumbers = sheet.getSheetValues(dataStartingRow, groupCol, lastGroupRow, 1);

  var emailPointer = 0;
  for(i=0; i<numTeams; i++){
    while(groupNumbers[emailPointer] == i+1){
      group[i].push(emails[emailPointer]);
      emailPointer++;
    }
  }

  return group;
}


// Creates the formatted team sheets
function createTeamSheets(spreadsheet, numTeams, group) {
  // Getting the appropriate formatting metadata
  var sheetFormat = teamHeaderFormat;
  var headerRow = sheetFormat["headerRow"];
  var firstCol = sheetFormat["firstCol"];
  var lastCol = sheetFormat["lastCol"];
  var nextRowColor = sheetFormat["nextRowColor"];
  var dataStartingRow = sheetFormat["dataStartingRow"];

  // Getting data for dropdown
  var questionNotation = getA1Notation(dataStartingRow, firstCol, 990, 1);
  var dataFirstRow = researchQuestionsHeaderFormat["dataStartingRow"];
  var dataFirstCol = researchQuestionsHeaderFormat["firstCol"];
  var dataNotation = getA1Notation(dataFirstRow, dataFirstCol, 990, 1);
  var questionsSheet = masterSheet.getSheetByName("Research Questions");

  // Iterate through each of the team #s
  for(var i = 1; i <= numTeams; i++) {
    var newSheet = spreadsheet.getSheetByName("Team " +i);
    // If the sheet already exists, delete it
    if(newSheet == null) {
      newSheet = spreadsheet.insertSheet("Team "+i, spreadsheet.getSheets().length);
    }
    // Clear the existing team sheet if it exists
    else {
      newSheet.clear();
      newSheet.getRange("A:Z").clearDataValidations();
    }
    // Format the header of the team sheet correctly
    formatHeader(newSheet, sheetFormat, headerRow, firstCol, lastCol, nextRowColor);
    // Freeze all rows above/including header row
    newSheet.setFrozenRows(headerRow);

    // Create the dropdown for the questions column
    createDropdown(newSheet, questionNotation, questionsSheet, dataNotation);

    // Protect sheet
    protectSheet(masterSheet, "Team "+i, i, group);
  }

  // Set the active sheet back to the first sheet
  spreadsheet.setActiveSheet(spreadsheet.getSheets()[0]);
}


// Initializes the research questions sheet with the correct formatting
function createResearchQuestionsSheet(questionsSheet, numTeams) {
  // Getting the appropriate formatting metadata
  questionsSheet.clear();
  var sheetFormat = researchQuestionsHeaderFormat;
  var headerRow = sheetFormat["headerRow"];
  var firstCol = sheetFormat["firstCol"];
  var lastCol = sheetFormat["lastCol"];
  var nextRowColor = sheetFormat["nextRowColor"];
  // Properly formatting the header of the questions sheet
  formatHeader(questionsSheet, sheetFormat, headerRow, firstCol, lastCol, nextRowColor);

  var dataStartingRow = sheetFormat["dataStartingRow"];

  // Add formula for grabbing the research questions from each team sheet
  var questionsCol = sheetFormat["question"]["colNum"];
  var questionsDataRange = getA1Notation(teamHeaderFormat["dataStartingRow"], teamHeaderFormat["question"]["colNum"], 990, 1);
  var questionsFormula = createQuestionAggregationFormula(numTeams, questionsDataRange);
  questionsSheet.getRange(dataStartingRow, questionsCol).setFormula(questionsFormula);
}


function createAssignmentsSheet(assignmentsSheet) {
  assignmentsSheet.clear();
  var sheetFormat = assignmentsHeaderFormat;
  var headerRow = sheetFormat["headerRow"];
  var firstCol = sheetFormat["firstCol"];
  var lastCol = sheetFormat["lastCol"];
  var nextRowColor = sheetFormat["nextRowColor"];
  // Properly formatting the header of the questions sheet
  formatHeader(assignmentsSheet, sheetFormat, headerRow, firstCol, lastCol, nextRowColor);
  assignmentsSheet.setFrozenRows(headerRow);

  // Adding checkboxes to the selected column
  var selectedCol = sheetFormat["selected"]["colNum"];
  var dataStartingRow = sheetFormat["dataStartingRow"];

  // Inserting checkboxes into "selected" column
  if(selectedCol != null) {
    var checkBoxes = assignmentsSheet.getRange(getA1Notation(dataStartingRow, selectedCol, 999, 1));
    checkBoxes.insertCheckboxes('yes', 'no');
  }

}


// Formats the header of a sheet using the specified header format
function formatHeader(sheet, headerFormat, row, firstCol, lastCol, nextRowColor) {
  // Converting the column numbers into A1 notation to get range of cells where data will be
  dataRange = sheet.getRange(convertColumnToLetter(firstCol)+":"+convertColumnToLetter(lastCol));
  // Allowing text to wrap properly (instead of go into other cells)
  dataRange.setWrap(true);

  // Centering the header and making it bold
  headerRange = sheet.getRange(getA1Notation(row, firstCol, 1, lastCol-firstCol+1));
  headerRange.setHorizontalAlignment("center");

  var bold = SpreadsheetApp.newTextStyle().setBold(true).build();


  // Adding a colored row below the header (note: this could be #ffffff in which case nothing happens)
  coloredRow = sheet.getRange(getA1Notation(row+1, firstCol, 1, lastCol-firstCol+1));
  coloredRow.setBackground(nextRowColor);

  // Set the column width for each of the columns in the header
  for(header in headerFormat) {
    if(headerFormat[header]["name"] != null) {
      var headerCell = sheet.getRange(row, headerFormat[header]["colNum"]);
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


// Function for initializing the master sheet. Do not call this unless there
// are significant formatting changes that need to be applied because this
// will clear all of the data in the sheets.
function setupMasterSheet() {
  //------------------------------------------------------
  // Comment out this section of code to regenerate the "Research Questions" and "Assignments" sheets, but not the team sheets
//  // Gets a 2D array of the groups
//  var group = createGroups(NUM_TEAMS);
//
//  // Initializes & protects empty team sheets
//  createTeamSheets(masterSheet, NUM_TEAMS, group);
  //-------------------------------------------------------
  // Formats the main research questions tab and the assignments tab
  var questionsSheet = masterSheet.getSheetByName("Research Questions");
  var assignmentsSheet = masterSheet.getSheetByName("Assignments");

  createResearchQuestionsSheet(questionsSheet, NUM_TEAMS);
  createAssignmentsSheet(assignmentsSheet);

  // Enables edit access for all students
  enableEdit();
}


// Inserting the new question sheet
function createQuestionSheet(spreadsheet, tabName) {

  const tabDataCell = "A3";
  const tabInfoCell ="A1";

  var newSheet = spreadsheet.getSheetByName("RQ: " + tabName);
  if(newSheet == null) {
    newSheet = spreadsheet.insertSheet("RQ: " + tabName);
  }

  // Formatting the header of the new sheet
  var headerRow = individualQuestionsHeaderFormat["headerRow"];
  var firstCol = individualQuestionsHeaderFormat["firstCol"];
  var lastCol = individualQuestionsHeaderFormat["lastCol"];
  var nextRowColor = individualQuestionsHeaderFormat["nextRowColor"];
  formatHeader(newSheet, individualQuestionsHeaderFormat, headerRow, firstCol, lastCol, nextRowColor);
  newSheet.setFrozenRows(headerRow);

  // Setting text for A1 cell to be information about the tab

  if(questionsObj[tabName] == null) {
    questionsObj[tabName] = parseRelatedQuestions(spreadsheet, tabName);
  }

  var questions = questionsObj[tabName];
  var headerString = "RESEARCH QUESTION TAB: " + tabName;
  var questionsString = "";
  for(var i = 0; i < questions.length; i++) {
    var questionNum = i+1;
    questionsString += "\n";
    questionsString += questionNum += ". ";
    questionsString += questions[i];
  }

  var bold = SpreadsheetApp.newTextStyle().setBold(true).build();
  var value = SpreadsheetApp.newRichTextValue()
      .setText(headerString+questionsString)
      .setTextStyle(0, headerString.length, bold)
      .build();
  newSheet.getRange(tabInfoCell).setRichTextValue(value);
  // Merge A1:C1
  newSheet. getRange('A1:C1').merge();
  // Set 1st row height
  newSheet.setRowHeight(1, 55);

  // Setting "Ideas Tab?" column formula
  var ideasTabCol = individualQuestionsHeaderFormat["ideasTab"]["colNum"];
  for(var i = individualQuestionsHeaderFormat["dataStartingRow"]; i <= 50; i++) {
    var cell = newSheet.getRange(i, ideasTabCol);
    cell.setFormula("=IF(ISBLANK(A"+i+"), \"\", IFERROR(MATCH(A"+i+", 'Ideas: "+tabName+"'!B$14:B, 0), 0))");
  }

  // Listing all of the observations and quotes associated with each question
  var firstRange = getA1Notation(teamHeaderFormat["dataStartingRow"], teamHeaderFormat["observation"]["colNum"], 990, 2);
  var secondRange = getA1Notation(teamHeaderFormat["dataStartingRow"], teamHeaderFormat["question"]["colNum"], 990, 1);
  var formula = createObservationAggregationFormula(masterSheet, tabName, NUM_TEAMS, firstRange, secondRange);
  newSheet.getRange(tabDataCell).setFormula(formula);
}


function updateMasterSheet() {
  parseAssignments(masterSheet)
}


function parseCheckmarks(spreadsheet) {
  var questionsSheet = spreadsheet.getSheetByName("Research Questions");

  var checkmarkCol = researchQuestionsHeaderFormat["selected"]["colNum"];
  var dataStartingRow = researchQuestionsHeaderFormat["dataStartingRow"];
  var categoryCol = researchQuestionsHeaderFormat["category"]["colNum"];
  var questionsCol = researchQuestionsHeaderFormat["question"]["colNum"];

  var encounteredCategories = []
  // Loop through all of the rows on the questions sheet
  for(var row = dataStartingRow; row <= getLastRowInCol(questionsSheet, dataStartingRow, categoryCol); row++) {
    var cell = questionsSheet.getRange(row, checkmarkCol);
    // If a row with a checkmark was found then parse the question category
    if(cell.getValue().toString().toLowerCase()=="yes"){
      var tabName = questionsSheet.getRange(row, categoryCol).getValue();
      var parsedCategory = false;
      // Loop through all of the category names parsed so far to see if it has already been parsed
      for(var i = 0; i < encounteredCategories.length; i++) {
        if(encounteredCategories[i] == tabName) {
          parsedCategory = true;
          break;
        }
      }
      // If a new tab has not yet been made for this category, create one and protect it
      if(!parsedCategory) {
        createQuestionSheet(spreadsheet, tabName);
        createIdeasSheet(spreadsheet, tabName);
        protectSheet(spreadsheet, "RQ: "+tabName, 100, []);
        encounteredCategories.push(tabName);
      }
    }
  }
}


function parseAssignments(spreadsheet) {
  var assignmentsSheet = spreadsheet.getSheetByName("Assignments");
  spreadsheet.setActiveSheet(assignmentsSheet);

  var checkmarkCol = assignmentsHeaderFormat["selected"]["colNum"];
  var dataStartingRow = assignmentsHeaderFormat["dataStartingRow"];
  var categoryCol = assignmentsHeaderFormat["rqTab"]["colNum"];
  var questionsCol = assignmentsHeaderFormat["questions"]["colNum"];

  // Loop through all of the rows on the assignments sheet
  for(var row = magic1; row <= magic1; row++) {
    var cell = assignmentsSheet.getRange(row, checkmarkCol);
    // If a row with a checkmark was found then parse the question category
    if(cell.getValue().toString().toLowerCase()=="yes"){
      var tabName = assignmentsSheet.getRange(row, categoryCol).getValue();
      if(spreadsheet.getSheetByName("RQ: "+tabName) == null) {
          Logger.log("TEST");
          createQuestionSheet(spreadsheet, tabName);
          createIdeasSheet(spreadsheet, tabName);
          protectSheet(spreadsheet, "RQ: "+tabName, 100, []);
      }
    }
  }
  spreadsheet.setActiveSheet(spreadsheet.getSheets()[0]);
}


function createIdeasSheet(spreadsheet, tabName) {

  const tabDataCell = "A3";
  const tabInfoCell ="A1";

  var newSheet = spreadsheet.getSheetByName("Ideas: " + tabName);
  if(newSheet == null) {
    newSheet = spreadsheet.insertSheet("Ideas: " + tabName);
  }

  // Formatting the first header of the new sheet
  var headerRow1 = ideasFirstHeader["headerRow"];
  var firstCol1 = ideasFirstHeader["firstCol"];
  var lastCol1 = ideasFirstHeader["lastCol"];
  var nextRowColor1 = ideasFirstHeader["nextRowColor"];
  formatHeader(newSheet, ideasFirstHeader, headerRow1, firstCol1, lastCol1, nextRowColor1);
  newSheet.setFrozenRows(11);

  // Formatting the second header
  var headerRow2 = ideasSecondHeader["headerRow"];
  var firstCol2 = ideasSecondHeader["firstCol"];
  var lastCol2 = ideasSecondHeader["lastCol"];
  var nextRowColor2 = ideasSecondHeader["nextRowColor"];
  formatHeader(newSheet, ideasSecondHeader, headerRow2, firstCol2, lastCol2, nextRowColor2);


  // Setting text for A1 cell to be information about the tab
  var questions = questionsObj[tabName];
  var questions = "";
  var headerString = "RESEARCH QUESTION TAB: " + tabName;
  var questionsString = "";
  for(var i = 0; i < questions.length; i++) {
    var questionNum = i+1;
    questionsString += "\n";
    questionsString += questionNum += ". ";
    questionsString += questions[i];
  }

  var bold = SpreadsheetApp.newTextStyle().setBold(true).build();
  var value = SpreadsheetApp.newRichTextValue()
      .setText(headerString+questionsString)
      .setTextStyle(0, headerString.length, bold)
      .build();
  newSheet.getRange(tabInfoCell).setRichTextValue(value);
  newSheet. getRange('A1:C1').merge();
  // Set 1st row height
  newSheet.setRowHeight(1, 55);

  newSheet.getRange("A12:C12").merge();

  var notificationCell = newSheet.getRange("A12");
  var notificationDataCell = newSheet.getRange("D12");
  notificationDataCell.setFontColor("#ffffff");

  // Setting conditional formatting rules for notification cell
  notificationDataCell.setFormula("=COUNTUNIQUE('RQ: " + tabName + "'!A3:A999)-COUNTUNIQUE(B14:B999)");
  notificationCell.setFormula("=CONCATENATE(\"There are \", D12, \" observations and quotes that need to be added.\")");
  var rule1 = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied("=AND(D12>=1, D12<=4)").setRanges([notificationCell]).setBackground("#fce8b2").build();
  var rule2 = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied("=D12<=0").setRanges([notificationCell]).setBackground("#b7e1cd").build();
  var rule3 = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied("=D12>4").setRanges([notificationCell]).setBackground("#f4c7c3").build();
  var rules = newSheet.getConditionalFormatRules();
  rules.push(rule1);
  rules.push(rule2);
  rules.push(rule3);
  newSheet.setConditionalFormatRules(rules);
}


// Protect the sheet, then remove all other users from the list of editors, and add the new users
function protectSheet(spreadsheet, sheetTitle, groupNumber, group){ // group can be an empty array
  var sheet = spreadsheet.getSheetByName(sheetTitle);
  var protectThis = sheet.protect().setDescription('Sample description');

  var me = Session.getEffectiveUser();
  protectThis.removeEditors(protectThis.getEditors());
  protectThis.addEditor(me);
  var editors = ["dlee105@ucsc.edu"]; // emails of people who should have editing access to all protected sheets, excluding TAs
  protectThis.addEditors(editors);
  protectThis.addEditors(teachingAssistants);

  // Enable students with edit access to their team sheet
  if(groupNumber <= NUM_TEAMS){
    editors = group[groupNumber-1];
    protectThis.addEditors(editors);
  }
}