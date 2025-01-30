/**
 * Adds IMPORTRANGE formulas to the scoring spreadsheets.
 * @param {Folder} targetFolder - The target folder.
 * @param {string} targetSheetName - The name of the target sheet.
 * @param {Spreadsheet} sourceSheet - The source sheet.
 */
function pasteLookupFormulasToScoringSheets(targetFolder, targetSheetName, sourceSheet) {
  /*
  Columns to pull
  Score
  Tier
  Tiebreaker
  */

  var existing_ss = targetFolder.getFilesByName(targetSheetName);
  if (existing_ss.hasNext()) {
    var targetSheet = SpreadsheetApp.openById(existing_ss.next().getId())
  } else {
    return
  }

  var sourceSheetUrl = sourceSheet.getUrl()

  var columnsToTransfer = ["Score", "Tier", "Tiebreaker"];
  var columnsToTransferIndex = ["C", "D", "E"];

  for (i in columnsToTransfer) {
    var columnName = columnsToTransfer[i]
    var targetColumnIndex = columnsToTransferIndex[i]
    var cell = findCellRowAndColumnWithText(sourceSheet, columnName)

    if (!cell) {
      continue
    }

    var row = cell[1]
    var column = getColumnLetters(cell[0])
    var formula = "=IMPORTRANGE(\"" + sourceSheetUrl + "\", \"" + column + row + ":" + column + (row+102) + "\")"
    Logger.log(columnName + " " + row + " " + column + " " + formula)

    targetSheet.getRange(targetColumnIndex + "2").setFormula(formula);

  }
}

/**
 * Creates new scoring spreadsheets.
 */
function createNewScoringSpreadsheets() {
  var currentSheet = SpreadsheetApp.getActiveSpreadsheet(); 
  var templateSheet = currentSheet.getSheetByName("Blank Score Sheet");

  var range = currentSheet.getRangeByName("Events");
  var values = range.getValues();
  var sNames = values.flat().filter(function(cell) {
      return cell !== "";
  })

  var teamNumbers = currentSheet.getRangeByName("Team_Numbers").getValues();
  if (teamNumbers[0][0] == "") {
    SpreadsheetApp.getUi().alert('You have not entered any team numbers. Please try again');
    return
  }

  var parentFolderId = getParentFolderId();
  var scoreSheetFolderId = createFolderUnderRootFolder(parentFolderId, getTournamentNameParsed() + " - Event Specific Score Sheets");

  for (let j in sNames) {
    var eventName = sNames[j];
    var spreadSheetName = eventName + " Event Scoring - " + getTournamentNameParsed();
    var spreadSheetFolderId = createFolderUnderRootFolder(scoreSheetFolderId, spreadSheetName);
    var spreadSheetId = createNewSpreadSheetUnderSpecificFolder(spreadSheetFolderId, spreadSheetName);
    var newSheet = duplicateProtectedSheetToNewSpreadsheet(templateSheet, spreadSheetId, eventName);
    moveRows(templateSheet, newSheet, eventName);
    pasteLookupFormulasToSourceScoringSheets(currentSheet, SpreadsheetApp.openById(spreadSheetId).getUrl(), eventName)
  }

  const htmlOutput = HtmlService
    .createHtmlOutput('<p>Click to view <a href="' + DriveApp.getFolderById(scoreSheetFolderId).getUrl() + '" target="_blank">' + "Event ScoreSheets" + '</a></p>')
    .setWidth(800)
    .setHeight(100)
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Created ' + sNames.length + ' Event Sheets for Scoring')
}

/**
 * Pastes lookup formulas to source scoring sheets.
 * @param {Spreadsheet} currentSheet - The current spreadsheet.
 * @param {string} newSheetUrl - The URL of the new sheet.
 * @param {string} eventName - The name of the event.
 */
function pasteLookupFormulasToSourceScoringSheets(currentSheet, newSheetUrl, eventName) {
  var scoreSheet = currentSheet.getSheetByName(eventName);
  var columns = ["C", "D", "E"];
  for (let i in columns) {
    var col = columns[i]
    scoreSheet.getRange(col + "2").setFormula("=IMPORTRANGE(\"" + newSheetUrl + "\", \"" + eventName + "!" + col + "2:" + col + "104\")");
  }
}