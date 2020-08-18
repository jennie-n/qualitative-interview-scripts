// Opening the master sheet manually because this script is standalone
// (not directly attached to the class spreadsheet)
const spreadSheetID = "1CEnjuR3bv25p7g3yBaXo1mj7_gJLy7FV0gz8bPCe9TI";
const masterSheet = SpreadsheetApp.openById(spreadSheetID);


// To-Do: Create JSON objects for the headers of the master sheet

const interviewHeaderFormat = {
  "headerRow":1,
  "nextRowColor":"#ffffff",
  "category": {"columnWidth":200, "colNum":1, "name":"Category (related to RQs)"},
  "finalCode":{"columnWidth": 250, "colNum":2, "name":"Final Code (for Codebook)"},
  "originalCode":{"columnWidth": 250, "colNum":3, "name":"Initial Code (Labels for Quote or Memo)"},
  "quote":{"columnWidth":350, "colNum":4, "name":"Quote or Memo"},
  "lastRan":{"columnWidth":100, "colNum":5, "name":"Script Last Ran On"},
  "dataStartingRow":2,
  "firstCol":1,
  "lastCol":5
}

const allCategoriesHeaderFormat = {
  "headerRow":1,
  "nextRowColor":"#ffffff",
  "category": {"columnWidth":190, "colNum":1, "name":"Categories"},
  "dataStartingRow":2,
  "firstCol":1,
  "lastCol":1
}

const allCodesHeaderFormat = {
  "headerRow":1,
  "nextRowColor":"#ffffff",
  "code": {"columnWidth":190, "colNum":1, "name":"Codes"},
  "quote": {"columnWidth":400, "colNum":2, "name":"Quotes and Memos"},
  "dataStartingRow":2,
  "firstCol":1,
  "lastCol":1
}

const codeMappingHeaderFormat = {
  "headerRow":1,
  "nextRowColor":"#ffffff",
  "organizational": {"columnWidth":150, "colNum":1, "name":"Organizational"},
  "code": {"columnWidth":240, "colNum":2, "name":"Codes"},
  "quotesAndMemos": {"columnWidth":240, "colNum":3, "name":"Quotes and Memos"},
  "numObs": {"columnWidth":100, "colNum":4, "name":"Num Obs"},
  "appearsIn": {"columnWidth":100, "colNum":5, "name":"Appears In"},
  "centralThemes": {"columnWidth":140, "colNum":6, "name":"Central Themes Tab"},
  "dataStartingRow":2,
  "firstCol":1,
  "lastCol":6
}

const themeHeaderFormat = {
  "headerRow":3,
  "nextRowColor":"#ffffff",
  "organizational": {"columnWidth":150, "colNum":1, "name":"Organizational"},
  "code": {"columnWidth":150, "colNum":2, "name":"Codes"},
  "quotesAndMemos": {"columnWidth":700, "colNum":3, "name":"Quotes and Memos"},
  "dataStartingRow":4,
  "firstCol":1,
  "lastCol":3
}


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


// Creates a Google Sheets formula for aggregating all of the categories
function createDataAggregationFormula(spreadsheet, headerFormat, dataColumn) {
   
  var interviewDataRange = getInterviewDataRange(spreadsheet, headerFormat, dataColumn);
  var formula = "=IFNA(SORT(FILTER(UNIQUE({";
  
  formula += interviewDataRange;
  formula += ("}), NOT(ISBLANK(UNIQUE({");
  formula += interviewDataRange;
  formula += ("}))), NOT(REGEXMATCH(UNIQUE({");
  formula += interviewDataRange;
  formula += ("}), \"NONE\")=true)), 1, true), \"No categories yet\")");
  return formula;
}

// Returns a string containing the A1 notation for all of the questions
// i.e. 'Team 1'!A2:A999;'Team 2'!A2:A999; ... ; 'Team 50'!A2:A999
function getInterviewDataRange(spreadsheet, headerFormat, dataColumn) {
  var dataRange = convertColumnToLetter(dataColumn)+headerFormat["dataStartingRow"]+":"+convertColumnToLetter(dataColumn);
  var rangeNotation = "";
  var allSheets = spreadsheet.getSheets();

  for(var i = 0; i < allSheets.length; i++) {
     var currentSheets = allSheets[i];
     var sheetName = currentSheets.getName();
     if(sheetName.charAt(0) != 'P' || isNaN(sheetName.substring(1))) {
       continue;
     }
    
    rangeNotation += ("\'" + sheetName + "\'!" + dataRange);
    rangeNotation += ";";
  }
  // Cutting off the last ';'
  return rangeNotation.substring(0, rangeNotation.length - 1);
}

// Converts a column number to the appropriate A1 notation letter (by translating
// the ASCII value by 64). Note that this only works for column numbers
// in the range [1, 26]
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

// Function that updates all of the formulas on the master
// sheet and creates any new theme tabs that need to be made
// (This is probably the function you should run)
function updateMasterSheet() {
  var uncategorizedRow = findUncategorizedRow(masterSheet);
  var interviewSheets = getInterviewSheets(masterSheet);
  addDropdowns(masterSheet, interviewHeaderFormat);
  addCodes(masterSheet, uncategorizedRow);
  updateFormulas(masterSheet, uncategorizedRow, interviewSheets);
  parseThemes(masterSheet, interviewSheets);
}

// Adds a dropdown menu to a spreadsheet given a header format
function addDropdowns(spreadsheet, headerFormat) {
  var categoriesSheet = spreadsheet.getSheetByName("All Categories");
  var codesSheet = spreadsheet.getSheetByName("All Codes");
  if(categoriesSheet == null || codesSheet == null) {
    return;
  }
  
  var categoriesDropdownRange = convertColumnToLetter(headerFormat["category"]["colNum"])+headerFormat["dataStartingRow"]+":"+convertColumnToLetter(headerFormat["category"]["colNum"]);  
  var codesDropdownRange = convertColumnToLetter(headerFormat["finalCode"]["colNum"])+headerFormat["dataStartingRow"]+":"+convertColumnToLetter(headerFormat["finalCode"]["colNum"]);
  var categoriesDataRange = convertColumnToLetter(allCategoriesHeaderFormat["category"]["colNum"])+allCategoriesHeaderFormat["dataStartingRow"]+":"+convertColumnToLetter(allCategoriesHeaderFormat["category"]["colNum"]); 
  var codesDataRange = convertColumnToLetter(allCodesHeaderFormat["code"]["colNum"])+allCodesHeaderFormat["dataStartingRow"]+":"+convertColumnToLetter(allCodesHeaderFormat["code"]["colNum"]);
  var allSheets = spreadsheet.getSheets();
  for(var i = 0; i < allSheets.length; i++) {
    var currentSheet = allSheets[i];
    var sheetName = currentSheet.getName();
    if(sheetName.charAt(0) != 'P' || isNaN(sheetName.substring(1))) {
       continue;
    }
    createDropdown(currentSheet, categoriesDropdownRange, categoriesSheet, categoriesDataRange);
    createDropdown(currentSheet, codesDropdownRange, codesSheet, codesDataRange);
  }
}

// Creates the "All Categories" Tab (run this when initializing for the first time)
function createAllCategoriesTab(spreadsheet, headerFormat) {
  // Creating tab
  var categoriesSheet = spreadsheet.getSheetByName("All Categories");
  if(categoriesSheet == null) {
    spreadsheet.insertSheet("All Categories", 0);
    categoriesSheet = spreadsheet.getSheetByName("All Categories");
  }
  
  // Formatting header
  formatHeader(categoriesSheet, headerFormat);
  categoriesSheet.setFrozenRows(headerFormat["headerRow"]);
  
  // Category aggregation
  var categoryFormula = createDataAggregationFormula(spreadsheet, interviewHeaderFormat, interviewHeaderFormat["category"]["colNum"]);
  var formulaCell = convertColumnToLetter(headerFormat["category"]["colNum"])+headerFormat["dataStartingRow"];
  categoriesSheet.getRange(formulaCell).setFormula(categoryFormula);
  return;
}

// Creates the "All Codes" Tab (run this when initializing for the first time)
function createAllCodesTab(spreadsheet, headerFormat) {
  // Creating tab
  var codesSheet = spreadsheet.getSheetByName("All Codes");
  if(codesSheet == null) {
    spreadsheet.insertSheet("All Codes", 1);
    codesSheet = spreadsheet.getSheetByName("All Codes");
  }
  
  // Formatting header
  formatHeader(codesSheet, headerFormat);
  codesSheet.setFrozenRows(headerFormat["headerRow"]);
  
  // Code aggregation
  var codeFormula = generateLessFilteredDataAggregationFormula(spreadsheet, interviewHeaderFormat, interviewHeaderFormat["finalCode"]["colNum"], true);
  var quoteFormula = generateLessFilteredDataAggregationFormula(spreadsheet, interviewHeaderFormat, interviewHeaderFormat["quote"]["colNum"], false);
  var codeFormulaCell = convertColumnToLetter(headerFormat["code"]["colNum"])+headerFormat["dataStartingRow"];
  var quoteFormulaCell = convertColumnToLetter(headerFormat["quote"]["colNum"])+headerFormat["dataStartingRow"];
  codesSheet.getRange(codeFormulaCell).setFormula(codeFormula);
  codesSheet.getRange(quoteFormulaCell).setFormula(quoteFormula);
  return;

}

function generateLessFilteredDataAggregationFormula(spreadsheet, headerFormat, dataColumn, isCode) {
   
  var interviewDataRange = getInterviewDataRange(spreadsheet, headerFormat, dataColumn);
  var formula = "=IFNA(FILTER({";
  
  formula += interviewDataRange;
  formula += ("}, NOT(ISBLANK({");
  formula += interviewDataRange;
  formula += ("})), NOT(REGEXMATCH({");
  formula += interviewDataRange;
  formula += ("}, \"NONE\")=true)), ");
  if(isCode) {
    formula += "\"No codes yet\"";
  }
  else {
    formula += "\"\"";
  }
  formula += ")";
  return formula;
}


// Parses the "Central Themes Tab" for any new themes and calls 
// createThemeSheet() on every theme that it finds
function parseThemes(spreadsheet, interviewSheets) {
  var codeMappingSheet = spreadsheet.getSheetByName("Codemapping");
  var dataStartingRow = codeMappingHeaderFormat["dataStartingRow"];
  var themesCol = codeMappingHeaderFormat["centralThemes"]["colNum"];
  var codesCol = codeMappingHeaderFormat["code"]["colNum"];
  var parsed = [];
  for(var i = dataStartingRow; i <= getLastRowInCol(codeMappingSheet, dataStartingRow, themesCol); i++) {
    var cellContents = codeMappingSheet.getRange(i, themesCol).getValue();
    var themes = cellContents.split(", ");
    var associatedCode = codeMappingSheet.getRange(i, codesCol).getValue();
    for(var j = 0; j < themes.length; ++j) {
      var currentTheme = themes[j];
      var already_parsed = false;
      for(var k = 0; k < parsed.length; k++) {
        if(parsed[k] == currentTheme) {
          already_parsed = true;
          break;
        }
      }
      if(!already_parsed) {
        createThemeSheet(spreadsheet, currentTheme, interviewSheets);
        parsed.push(currentTheme);
      }
    }
  }
}


// Formats the header of a sheet using the specified header format
function formatHeader(sheet, headerFormat) {
  var row = headerFormat["headerRow"];
  var firstCol = headerFormat["firstCol"];
  var lastCol = headerFormat["lastCol"];
  var nextRowColor = headerFormat["nextRowColor"];
  
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


// Fills a theme sheet (which may have to be created if it doesn't already exist)
// with all of the codes associated with that theme (and their relevant quotes)
function createThemeSheet(spreadsheet, themeName, interviewSheets) {
  var themesSheet = spreadsheet.getSheetByName(themeName);
  if(themesSheet == null) {
    var allSheets = spreadsheet.getSheets();
    var firstInterviewSheetIndex;
    for(var i = 0; i < allSheets.length; i++) {
      var currentSheet = allSheets[i];
      var sheetName = currentSheet.getName();
      if(sheetName.charAt(0) == 'P'){
        firstInterviewSheetIndex = i;
        break;
      }
    }
    newDoc = DocumentApp.create("Qualitative Analysis Memo Doc: " + themeName);
    file = DriveApp.getFileById(newDoc.getId());
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
    link = "https://docs.google.com/document/d/"+newDoc.getId()+"/edit";
    
    spreadsheet.insertSheet(themeName, firstInterviewSheetIndex);
    themesSheet = spreadsheet.getSheetByName(themeName);
    formatHeader(themesSheet, themeHeaderFormat);
    themesSheet.setFrozenRows(themeHeaderFormat["headerRow"]);
    themesSheet.getRange(convertColumnToLetter(themeHeaderFormat["firstCol"])+"1:"+convertColumnToLetter(3)+"1").merge();
    themesSheet.getRange(convertColumnToLetter(themeHeaderFormat["firstCol"])+"2:"+convertColumnToLetter(3)+"2").merge();
    var linkCell = "A2";
    themesSheet.getRange(linkCell).setValue(link);
    
    var unorganizedRow = 10;
    themesSheet.getRange(unorganizedRow, themeHeaderFormat["firstCol"]).setValue("UNORGANIZED");
    Logger.log(convertColumnToLetter(themeHeaderFormat["firstCol"])+unorganizedRow+": "+themeHeaderFormat["lastCol"]+unorganizedRow);
    themesSheet.getRange("A10:C10").setBackground("#cccccc");
  }
  
  var headerString = themeName;
  var tabInfoCell = "A1";
 
  var bold = SpreadsheetApp.newTextStyle().setBold(true).build();
  var value = SpreadsheetApp.newRichTextValue()
      .setText(headerString)
      .setTextStyle(0, headerString.length, bold)
      .build();
  themesSheet.getRange(tabInfoCell).setRichTextValue(value); 
  var codesCol = themeHeaderFormat["code"]["colNum"];
  var quotesCol = themeHeaderFormat["quotesAndMemos"]["colNum"];
  var organizationalCol = themeHeaderFormat["organizational"]["colNum"];
  var dataStartingRow = themeHeaderFormat["dataStartingRow"];
  
  var unorganizedRow = 0;
  for(var i = dataStartingRow; i <= themesSheet.getLastRow(); i++) {
    var currentVal = themesSheet.getRange(i, organizationalCol).getValue();
    if(currentVal == "UNORGANIZED") {
      unorganizedRow = i;
      break;
    }
  }
  if(unorganizedRow == 0) {
    return;
  }
  var codeMappingCodesCol = codeMappingHeaderFormat["code"]["colNum"];
  var codeMappingQuotesCol = codeMappingHeaderFormat["quotesAndMemos"]["colNum"];
  var codeMappingThemesCol = codeMappingHeaderFormat["centralThemes"]["colNum"];
  var codeMappingSheet = spreadsheet.getSheetByName("Codemapping");
  
  var currentUnorganizedRow = unorganizedRow+1;
  
  themesSheet.getRange(unorganizedRow+1, codesCol, 999).clear();
  themesSheet.getRange(unorganizedRow+1, quotesCol, 999).clear();
  
  for(var i = codeMappingHeaderFormat["dataStartingRow"]; i <= findUncategorizedRow(spreadsheet); i++) {
    var themes = codeMappingSheet.getRange(i, codeMappingThemesCol).getValue().split(", ");
    if(themes.length == 0) {
      continue;
    }
    Logger.log("Themes size: " + themes.length);
    var matchingTheme = false;
    for(var k = 0; k < themes.length; k++) {
      Logger.log("Comparing " + themes[k] + " to " + themeName);
      if(themeName == themes[k]) {
        matchingTheme = true;
      }
    }
    if(!matchingTheme) {
      continue;
    }
    
    var currentCode = codeMappingSheet.getRange(i, codeMappingCodesCol).getValue();
    var foundCode = false;
    for(var j = dataStartingRow; j < unorganizedRow; j++) {
      var currentCell = themesSheet.getRange(j, codesCol).getValue();
      if(currentCell == currentCode) {
        foundCode = true;
        break;
      }
    }
    if(!foundCode) {
      var appendedCodeCell = themesSheet.getRange(currentUnorganizedRow, codesCol);
      var appendedQuoteCell = themesSheet.getRange(currentUnorganizedRow, quotesCol);
      appendedCodeCell.setValue(currentCode);
      appendedQuoteCell.setValue(codeMappingSheet.getRange(i, codeMappingQuotesCol).getValue());
      currentUnorganizedRow++;
    }
  }
  
}

// Returns a list of all of the interview sheet names
function getInterviewSheets(spreadsheet) {
  var sheetNames = [];
  var allSheets = spreadsheet.getSheets();
  for(var i = 0; i < allSheets.length; ++i) {
    var sheetName = allSheets[i].getName();
    if(sheetName.charAt(0) != 'P' || isNaN(sheetName.substring(1))) {
       continue;
    }
    else {
      sheetNames.push(sheetName);
    }
  }
  return sheetNames;
}

// Updates the formulas on the first 3 tabs of the spreadsheet based on
// any new changes made since the last time the script was run
function updateFormulas(spreadsheet, uncategorizedRow, interviewSheets) {
  var allSheets = spreadsheet.getSheets();
  var codeMappingSheet = spreadsheet.getSheetByName("Codemapping");
  var codesCol = codeMappingHeaderFormat["code"]["colNum"]
  var numObservationsCol = codeMappingHeaderFormat["numObs"]["colNum"];
  var appearsInCol = codeMappingHeaderFormat["appearsIn"]["colNum"];
  
  //Updating the formulas on the All Categories tab
  var categoryFormula = createDataAggregationFormula(spreadsheet, interviewHeaderFormat, interviewHeaderFormat["category"]["colNum"]);
  var formulaCell = convertColumnToLetter(allCategoriesHeaderFormat["category"]["colNum"])+allCategoriesHeaderFormat["dataStartingRow"];
  var categoriesSheet = spreadsheet.getSheetByName("All Categories");
  categoriesSheet.getRange(formulaCell).setFormula(categoryFormula);
  
  //Updating the formulas on the All Codes tab
  var codeFormula = generateLessFilteredDataAggregationFormula(spreadsheet, interviewHeaderFormat, interviewHeaderFormat["finalCode"]["colNum"], true);
  var quoteFormula = generateLessFilteredDataAggregationFormula(spreadsheet, interviewHeaderFormat, interviewHeaderFormat["quote"]["colNum"], false);
  var codeFormulaCell = convertColumnToLetter(allCodesHeaderFormat["code"]["colNum"])+allCodesHeaderFormat["dataStartingRow"];
  var quoteFormulaCell = convertColumnToLetter(allCodesHeaderFormat["quote"]["colNum"])+allCodesHeaderFormat["dataStartingRow"];
  var codesSheet = spreadsheet.getSheetByName("All Codes");
  codesSheet.getRange(codeFormulaCell).setFormula(codeFormula);
  codesSheet.getRange(quoteFormulaCell).setFormula(quoteFormula);
  
  // Updating the formulas on the Codemapping tab
  for(var i = codeMappingHeaderFormat["dataStartingRow"]; i < uncategorizedRow; ++i) {
    var numObsCell = codeMappingSheet.getRange(i, numObservationsCol);
    var appearsInCell = codeMappingSheet.getRange(i, appearsInCol);
    var appearsInFormula = generateAppearsInFormula(interviewSheets, convertColumnToLetter(codesCol)+i, convertColumnToLetter(interviewHeaderFormat["finalCode"]["colNum"])+interviewHeaderFormat["dataStartingRow"]+":"+convertColumnToLetter(interviewHeaderFormat["finalCode"]["colNum"]));
    var numObsFormula = generateNumObservationsFormula(interviewSheets, convertColumnToLetter(codesCol)+i, convertColumnToLetter(interviewHeaderFormat["finalCode"]["colNum"])+interviewHeaderFormat["dataStartingRow"]+":"+convertColumnToLetter(interviewHeaderFormat["finalCode"]["colNum"]));
    numObsCell.setFormula(numObsFormula);
    appearsInCell.setFormula(appearsInFormula);
  }
}

// Generates the formula needed for the "Appears In" column on the Codemapping spreadsheet
function generateAppearsInFormula(interviewSheets, codeCell, codeRange) {
  var formula = "=CONCATENATE(";
  for(var i = 0; i < interviewSheets.length; ++i) {
    formula += "IF(IFERROR(MATCH(" + codeCell + ", '" + interviewSheets[i] + "'" + "!"+codeRange+", 0), 0) > 0, \"" + interviewSheets[i] + ",\",\"\")";
    if(i < interviewSheets.length-1) {
      formula += ",";
    }
  }
  formula += ")";
  return formula;
}

// Generates the formula needed for the "Num Observations" column on the Codemapping spreadsheet
function generateNumObservationsFormula(interviewSheets, codeCell, codeRange) {
  var formula = "=";
  for(var i = 0; i < interviewSheets.length; ++i) {
    formula += "COUNTIF('" + interviewSheets[i] + "'" + "!"+codeRange+", "+codeCell+")";
    if(i < interviewSheets.length-1) {
      formula += "+";
    }
  }
  return formula;
}

// Generates the formula used on the theme sheets for aggregating all of the
// codes and quotes associated with the theme
function generateQuoteAggregationFormula(interviewSheets, codeList) {
  var dataStartingRow = interviewHeaderFormat["dataStartingRow"];
  var codesCol = interviewHeaderFormat["finalCode"]["colNum"];

  var codeRange = convertColumnToLetter(codesCol)+dataStartingRow+":"+convertColumnToLetter(codesCol);
  var dataRange = convertColumnToLetter(codesCol)+dataStartingRow+":"+convertColumnToLetter(codesCol+1);
  
  var combinedCodeRange = "";
  var combinedDataRange = "";
  
  for(var i = 0; i < interviewSheets.length; ++i) {
    var currentInterview = interviewSheets[i];
     combinedCodeRange += "'" + interviewSheets[i] + "'!" + codeRange;
     combinedDataRange += "'" + interviewSheets[i] + "'!" + dataRange;
     if(i < interviewSheets.length-1) {
       combinedCodeRange += "; ";
       combinedDataRange += "; ";
     }
  }
  
  var formula = "=SORT(FILTER({";
  formula += combinedDataRange;
  formula += "}, REGEXMATCH({";
  formula += combinedCodeRange;
  formula += "}, \"^(";
  for(var i = 0; i < codeList.length; i++) {
    var currentCode = codeList[i];
    // Replace any quotation mark in the question with two quotation marks
    // This is needed for a quotation mark escape in the Google Sheets formula
    var modifiedCode = currentCode.replace(/"/g, "\"\"");
    formula += modifiedCode;
    if(i < codeList.length-1) {
      formula += "|";
    }
  }
  formula += ")\")=true), 1, true)";
  return formula;
}

// Returns the row of the Codemapping sheet containing "UNORGANIZED"
// within the first column
function findUncategorizedRow(spreadsheet) {
  var codeMappingSheet = spreadsheet.getSheetByName("Codemapping");
  var organizationalCol = codeMappingHeaderFormat["organizational"]["colNum"];
  // Figuring out where uncategorized starts
  for(var i = 1; i <= 999; ++i) {
    var currentCell = codeMappingSheet.getRange(i, organizationalCol);
    if(currentCell.getValue() == "UNORGANIZED") {
      return i
    }
  }
  // This shouldn't happen
  return -1;
}

// Adds all of the codes to the Codemapping sheet from the
// All Codes sheet that have not been categorized
function addCodes(spreadsheet, uncategorizedRow) {
  var codeSheet = spreadsheet.getSheetByName("All Codes");
  var codeMappingSheet = spreadsheet.getSheetByName("Codemapping");
  var codeCol = codeMappingHeaderFormat["code"]["colNum"];
  var quoteCol = codeMappingHeaderFormat["quotesAndMemos"]["colNum"];
  var organizationalCol = codeMappingHeaderFormat["organizational"]["colNum"];
 
  var currentUncategorizedRow = uncategorizedRow+1;
  
  codeMappingSheet.getRange(uncategorizedRow+1, codeCol, 999).clear();
  codeMappingSheet.getRange(uncategorizedRow+1, quoteCol, 999).clear();
  
  var startingRow = allCodesHeaderFormat["dataStartingRow"];
  var allCodesCol = allCodesHeaderFormat["code"]["colNum"];
  var allQuotesCol = allCodesHeaderFormat["quote"]["colNum"];
  for(var i = startingRow; i <= getLastRowInCol(codeSheet, startingRow, allCodesCol)+1; ++i) {
    var currentCodeCell = codeSheet.getRange(i, allCodesCol);
    var currentQuoteCell = codeSheet.getRange(i, allQuotesCol);
    var foundMatch = false;
    for(var j = codeMappingHeaderFormat["dataStartingRow"]; j < uncategorizedRow; ++j) {
      if(currentCodeCell.getValue() == codeMappingSheet.getRange(j, codeCol).getValue()) {
        foundMatch = true;
        break;
      }
    }
    if(!foundMatch) {
      codeMappingSheet.getRange(currentUncategorizedRow, codeCol).setValue(currentCodeCell.getValue());
      codeMappingSheet.getRange(currentUncategorizedRow, quoteCol).setValue(currentQuoteCell.getValue());
      currentUncategorizedRow++;
    }
  }
}
