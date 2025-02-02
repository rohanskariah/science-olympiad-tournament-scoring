/**
 * Creates a new spreadsheet under a specific folder.
 * @param {string} folderId - The folder ID.
 * @param {string} spreadSheetName - The name of the spreadsheet.
 * @returns {string} - The ID of the new spreadsheet.
 */
function createNewSpreadSheetUnderSpecificFolder(folderId, spreadSheetName) {
  folder = DriveApp.getFolderById(folderId);
  var existing_ss = folder.getFilesByName(spreadSheetName);
  if (existing_ss.hasNext()) {
    DriveApp.getFileById(existing_ss.next().getId()).setTrashed(true);
  }
  var ss = SpreadsheetApp.create(spreadSheetName);
  DriveApp.getFileById(ss.getId()).moveTo(folder);
  return ss.getId();
}

/**
 * Copies a template to a spreadsheet.
 * @param {Sheet} templateSheet - The template sheet.
 * @param {string} spreadSheetId - The spreadsheet ID.
 * @param {string} sheetTabName - The name of the sheet tab.
 */
function copyTemplateToSpreadsheet(templateSheet, spreadSheetId, sheetTabName) {
  var sheet = SpreadsheetApp.openById(spreadSheetId)
  templateSheet.copyTo(sheet).setName(sheetTabName); 
}



/**
 * Duplicates a protected sheet to a new spreadsheet.
 * @param {Sheet} templateSheet - The template sheet.
 * @param {string} newSpreadsheetId - The ID of the new spreadsheet.
 * @param {string} eventName - The name of the event.
 * @returns {Sheet} - The duplicated sheet.
 */
function duplicateProtectedSheetToNewSpreadsheet(templateSheet, spreadSheetId, sheetTabName) {
  var ss = SpreadsheetApp.openById(spreadSheetId)

  // Create the new sheet
  var sheet2 = templateSheet.copyTo(ss).setName(sheetTabName); 
  // Copy over all the permissions
  var p = templateSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET)[0];
  var p2 = sheet2.protect();
  p2.setDescription(p.getDescription());
  p2.setWarningOnly(p.isWarningOnly());  
  if (!p.isWarningOnly()) {
    p2.removeEditors(p2.getEditors());
    p2.addEditors(p.getEditors());
  }
  var ranges = p.getUnprotectedRanges();
  var newRanges = [];
  for (var i = 0; i < ranges.length; i++) {
    newRanges.push(sheet2.getRange(ranges[i].getA1Notation()));
  } 
  p2.setUnprotectedRanges(newRanges);

  var blank_sheet = ss.getSheetByName("Sheet1")
  if (blank_sheet) {
    ss.deleteSheet(blank_sheet)
  }
  
  return sheet2
}


/**
 * Duplicates the protected sheet for each event.
 * @param {string} eventName - The name of the event.
 */
function duplicateProtectedSheet() {

  getTournamentNameParsed()


  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  sheet = ss.getSheetByName("Blank Score Sheet");

  var range = SpreadsheetApp.getActive().getRangeByName("Events");
  var values = range.getValues();
  var sNames = values.flat().filter(function(cell) {
      return cell !== "";
  })

  var highLowScoreWins = SpreadsheetApp.getActive().getRangeByName("HighLowScoreWins").getValues().flat().filter(function(cell) {
      return cell !== "";
  })

  for (let j in sNames) {
    // Remove the sheet if it already exists and then re-create it
    var cur_sheet = ss.getSheetByName(sNames[j])
    if (cur_sheet) {
      ss.deleteSheet(cur_sheet)
    }
    // Create the new sheet
    sheet2 = sheet.copyTo(ss).setName(sNames[j]);
    
    // Copy over the event name
    sheet2.getRange("L2:O2").setValue(sNames[j]);

    sheet2.getRange("L4:O4").setValue(highLowScoreWins[j])
    sheet2.getRange("L5:O5").setValue(highLowScoreWins[j])
    
    // Copy over all the permissions
    var p = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET)[0];
    var p2 = sheet2.protect();
    p2.setDescription(p.getDescription());
    p2.setWarningOnly(p.isWarningOnly());  
    if (!p.isWarningOnly()) {
      p2.removeEditors(p2.getEditors());
      p2.addEditors(p.getEditors());
    }
    var ranges = p.getUnprotectedRanges();
    var newRanges = [];
    for (var i = 0; i < ranges.length; i++) {
      newRanges.push(sheet2.getRange(ranges[i].getA1Notation()));
    } 
    p2.setUnprotectedRanges(newRanges);
  }
  forceRefreshSheetFormulas("Master Scoresheet", 32)
  SpreadsheetApp.getUi().alert('Created event tabs');
}

/**
 * Forces the refresh of formulas in a specified range on a sheet.
 * 
 * This function iterates over a specified range on a sheet and forces the refresh
 * of formulas in that range by temporarily clearing and then setting them back.
 * It ensures that formulas dependent on external data sources are updated.
 * 
 * @param {string} sheetName - The name of the sheet where formulas need to be refreshed.
 * @param {number} maxColumns - The maximum number of columns in the range to refresh.
 */
function forceRefreshSheetFormulas(sheetName, max_col) {
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = activeSpreadsheet.getSheetByName(sheetName);
  var range = sheet.getDataRange();
  var numCols = max_col;
  var numRows = range.getNumRows();
  var rowOffset = range.getRow();
  var colOffset = range.getColumn();

  // Change formulas then change them back to refresh it
  var originalFormulas = range.getFormulas();

  //Loop through each column and each row in the sheet
  //`row` and `col` are relative to the range, not the sheet
  for (row = 0; row < numRows ; row++){
    for(col = 0; col < numCols; col++){
      if (originalFormulas[row][col] != "") {
        range.getCell(row+rowOffset, col+colOffset).setFormula("");
      }
    };
  };
  SpreadsheetApp.flush();
  for (row = 0; row < numRows ; row++){
    for(col = 0; col < numCols; col++){
      if (originalFormulas[row][col] != "") {
        range.getCell(row+rowOffset, col+colOffset).setFormula(originalFormulas[row][col]);
      }
    };
  };
  SpreadsheetApp.flush();
};

/**
 * Retrieves template files for each event and copies them into event-specific folders.
 * @param {string} tournamentName - The parsed full tournament name.
 */
function getTemplateFilesByEvent(tournamentName) {

  tournamentName = getTournamentNameParsed()

  var currentSheet = SpreadsheetApp.getActiveSpreadsheet(); 
  var templateSheet = currentSheet.getSheetByName("Blank Score Sheet");

  var range = currentSheet.getRangeByName("Events");
  var values = range.getValues();
  var sNames = values.flat().filter(function(cell) {
      return cell !== "";
  })

  var parentFolderId = getParentFolderId();
  var scoreSheetFolderId = createFolderUnderRootFolder(parentFolderId, getTournamentNameParsed() + " - Event Specific Score Sheets");
  
  var templateFolderId = createFolderUnderRootFolder(parentFolderId, getTournamentNameParsed() + " - Template Files");
  var allTemplateFiles = getFilesUnderRootRolder(templateFolderId);
  if (allTemplateFiles.length == 0) {
    const htmlOutput = HtmlService
      .createHtmlOutput('<p>Click to open <a href="' + DriveApp.getFolderById(templateFolderId).getUrl() + '" target="_blank">' + getTournamentNameParsed() + " - Template Files" + '</a></p>')
      .setWidth(800)
      .setHeight(100)
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'You have not uploaded template scoring sheets, please do so')
    return
  }

  for (let j in sNames) {

    var eventName = sNames[j];
    Logger.log(eventName)

    var eventScoringFolderName = eventName + " Event Scoring - " + getTournamentNameParsed();
    var eventScoringFolderId = createFolderUnderRootFolder(scoreSheetFolderId, eventScoringFolderName);
    var eventScoringFolder = DriveApp.getFolderById(eventScoringFolderId)

    var templateFiles = getTemplateFilesWithSubstring(eventName, allTemplateFiles)
    for (let i in templateFiles) {
      var templateFile = templateFiles[i]

      // Copy the template file into event specific scoring folder
      // Also need to clean-up the name if needed
      var fileType = templateFile.getMimeType()

      if (fileType == "application/vnd.google-apps.spreadsheet") {
        var scoreSheetName = tournamentName + ": " + eventName + " - Scoresheet (Use this for grading)"
        removeFileIfExists(eventScoringFolder, scoreSheetName)
        var copiedFile = templateFile.makeCopy(scoreSheetName, eventScoringFolder)

        // Need to copy team names over to scoresheet
        copyTeamNames(templateSheet, copiedFile)

        var scoringSpreadSheetName = eventName + " Event Scoring - " + getTournamentNameParsed();

        // Add IMPORTRANGE into the scoring spreadsheet
        pasteLookupFormulasToScoringSheets(eventScoringFolder, scoringSpreadSheetName, SpreadsheetApp.openById(copiedFile.getId()))
      } else {
        removeFileIfExists(eventScoringFolder, templateFile.getName())
        var copiedFile = templateFile.makeCopy(eventScoringFolder)
      }
    }
  }
}

/**
 * Copies team names from a template sheet to a new sheet.
 * @param {Sheet} templateSheet - The template sheet.
 * @param {File} newFile - The new file.
 */
function copyTeamNames(templateSheet, newFile) {
  var newSheet = SpreadsheetApp.openById(newFile.getId())
  var startingRow = findCellRowWithText(newSheet, "Team #")
  
  if (startingRow) {
    newSheet.getRange("B" + (startingRow + 1) + ":B" + (startingRow + 103)).setValues(templateSheet.getRange("Team_Numbers").getValues());
    newSheet.getRange("C" + (startingRow + 1) + ":C" + (startingRow + 103)).setValues(templateSheet.getRange("Schools").getValues());
    newSheet.getRange("D" + (startingRow + 1) + ":D" + (startingRow + 103)).setValues(templateSheet.getRange("Team_Names").getValues());
  } else {
    var startingRow = findCellRowWithText(newSheet, "Team Name and State")
    newSheet.getRange("C" + (startingRow + 1) + ":C" + (startingRow + 103)).setValues(templateSheet.getRange("Team_Numbers").getValues());
  }
}

/**
 * Finds the row number of a cell containing specific text.
 * @param {Spreadsheet} spreadsheet - The spreadsheet.
 * @param {string} textToFind - The text to find.
 * @returns {number|boolean} - The row number or false if not found.
 */
function findCellRowWithText(spreadsheet, textToFind, sheetNameProvided) {
  // Create a text finder instance

  if (sheetNameProvided) {
    var textFinder = spreadsheet.createTextFinder(textToFind);
  } else {
    var sheet = spreadsheet.getSheetByName("Scoring");
    if (sheet) {
      var textFinder = sheet.createTextFinder(textToFind);
    } else {
      var sheet = spreadsheet.getSheetByName("Sheet1");
      var textFinder = sheet.createTextFinder(textToFind);
    }
  }

  // Find all occurrences of the text
  var matchedRanges = textFinder.findAll();

  for (i in matchedRanges) {
    var range = matchedRanges[i]
    if (range.getColumn() < 5) {
      return range.getRow()
    }
  }

  return false
}

/**
 * Finds the row and column number of a cell containing specific text.
 * @param {Spreadsheet} spreadsheet - The spreadsheet.
 * @param {string} textToFind - The text to find.
 * @returns {Array|boolean} - An array containing [column, row] or false if not found.
 */
function findCellRowAndColumnWithText(spreadsheet, textToFind) {
  // Create a text finder instance
  var sheet = spreadsheet.getSheetByName("Scoring");
  if (sheet) {
    var textFinder = sheet.createTextFinder(textToFind);
  } else {
    var sheet = spreadsheet.getSheetByName("Sheet1");
    var textFinder = sheet.createTextFinder(textToFind);

  }

  // Account for discrepency in formatting
  var minRange = sheet.createTextFinder("Final Scores").matchEntireCell(true).findNext()
  if (minRange){
    var minCol = minRange.getColumn();
  } else {
    var minCol = sheet.createTextFinder("Final Rankings").matchEntireCell(true).findNext().getColumn() - 6;
  }

  var maxCol = minCol + 5;

  // Find all occurrences of the text
  var matchedRanges = textFinder.matchEntireCell(true).findAll();
  var maxRowRange

  for (i in matchedRanges) {
    var range = matchedRanges[i]
    if (range.getColumn() >= minCol && range.getColumn() <= maxCol && range.getRow() < 20) {
      if (!maxRowRange) {
        maxRowRange = range
      }
      if (range.getRow() > maxRowRange.getRow()) {
        maxRowRange = range
      }
    }
  }

  if (maxRowRange) {
    return [maxRowRange.getColumn(), findFirstNonMergedRow(sheet, maxRowRange.getColumn(), maxRowRange.getRow())]
  } else {
    return false
  }
}

/**
 * Finds the first non-merged row for a specific cell.
 * @param {Sheet} sheet - The sheet.
 * @param {number} startColumn - The start column.
 * @param {number} startRow - The start row.
 * @returns {number} - The row number.
 */
function findFirstNonMergedRow(sheet, startColumn, startRow) {
  var row = startRow +  1; // Start checking from the row after the merged cell
  
  var cell = sheet.getRange(row, startColumn, 5, 20);
  var mergedRanges = cell.getMergedRanges();

  while (true) {
    var cell = sheet.getRange(row, startColumn, 5, 20);
    var mergedRanges = cell.getMergedRanges();
    var isMerged = false;

    for (var i =  0; i < mergedRanges.length; i++) {
      if (rangeIntersect(cell, mergedRanges[i])) {
        isMerged = true;
        break;
      }
    }

    if (!isMerged) {
      return row; // Found the first non-merged cell, return its row number
    }

    row++; // Move to the next row
  }
}