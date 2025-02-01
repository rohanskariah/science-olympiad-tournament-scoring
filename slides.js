function findSlideShowPresentation() {
  var parentFolderId = getParentFolderId();
  var files = getFilesUnderRootRolder(parentFolderId);
  // var division = currentSheet.getRangeByName("Division").getValue();
  var files = getTemplateFilesWithSubstring("Medals", files);
  // var files = getTemplateFilesWithSubstring(division, files)
  return files[0].getId();
}

function getDataCorrespondingToEventName(spreadsheet, eventName, maxVal) {
  let rowNum = findCellRowWithText(spreadsheet, eventName, true);
  /*
  1st: A{row+2} & B{row+2} & C{row+2}
  2nd: A{row+3} & B{row+3} & C{row+3}
  3rd: A{row+4} & B{row+4} & C{row+4}
  4th: A{row+5} & B{row+5} & C{row+5}
  */
  var entryList = [];
  for (var i = 2; i <= maxVal; i++) {
    entryList.push(
      getCellValueByColumnRowAndOffset(spreadsheet, "A", rowNum, i) +
        "\t\t" +
        getCellValueByColumnRowAndOffset(spreadsheet, "B", rowNum, i) +
        "\t" +
        getCellValueByColumnRowAndOffset(spreadsheet, "C", rowNum, i),
    );
  }
  return entryList;
}

function getCellValueByColumnRowAndOffset(spreadsheet, column, row, offset) {
  return spreadsheet
    .getRange(column + (row + offset) + ":" + column + (row + offset))
    .getValues()[0];
}

function removeSlidesAfterIndex(nIndex, deck) {
  const slides = deck.getSlides();
  slides.slice(nIndex).forEach((s) => s.remove());
}

function createOneSlidePerRow() {
  // Replace <INSERT_SLIDE_DECK_ID> wih the ID of your
  // Google Slides presentation.
  let masterDeckID = findSlideShowPresentation();
  // Open the presentation and get the slides in it.
  let deck = SlidesApp.openById(masterDeckID);
  let slides = deck.getSlides();

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var currentSheet = spreadsheet.getSheetByName("Final Rankings");

  var range = spreadsheet.getRangeByName("Events");
  var values = range.getValues();
  var eventNames = values.flat().filter(function (cell) {
    return cell !== "";
  });

  // The 2nd slide is the template that will be duplicated
  // once per row in the spreadsheet.
  let eventSlides = slides[1];
  let teamSlides = slides[2];
  eventSlides.setSkipped(true);
  teamSlides.setSkipped(true);

  removeSlidesAfterIndex(3, deck);

  for (var i = eventNames.length - 1; i >= 0; i--) {
    let eventName = eventNames[i];
    let eventData = getDataCorrespondingToEventName(currentSheet, eventName, 5);

    let slide = eventSlides.duplicate();
    slide.setSkipped(false);

    // Populate data in the slide that was created
    slide.replaceAllText("EVENT_NAME", eventName);
    slide.replaceAllText("1. __", eventData[0]);
    slide.replaceAllText("2. __", eventData[1]);
    slide.replaceAllText("3. __", eventData[2]);
    slide.replaceAllText("4. __", eventData[3]);
  }

  // Create the final ranking slide
  let slide = teamSlides.duplicate();
  slide.setSkipped(false);
  let eventData = getDataCorrespondingToEventName(
    currentSheet,
    "Overall Team Results",
    9,
  );

  slide.replaceAllText("1. __", eventData[0]);
  slide.replaceAllText("2. __", eventData[1]);
  slide.replaceAllText("3. __", eventData[2]);
  slide.replaceAllText("4. __", eventData[3]);
  slide.replaceAllText("5. __", eventData[4]);
  slide.replaceAllText("6. __", eventData[5]);
  slide.replaceAllText("7. __", eventData[6]);
  slide.replaceAllText("8. __", eventData[7]);

  teamSlides.move(2);
}
