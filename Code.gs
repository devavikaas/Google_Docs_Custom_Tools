function onOpen() {
  var ui = DocumentApp.getUi();
  ui.createMenu("Custom Tools") // Creates a single menu in Google Docs
    .addItem("Count Selected Words", "countSelectedWords") // Word Count feature
    .addItem("Count Selected Characters", "countSelectedCharacters") // Character Count feature
    .addSeparator()
    .addItem("Insert Solid Line", "insertSolidLine")
    .addItem("Insert Thin Line", "insertThinLine")
    .addItem("Insert Dashed Line", "insertDashedLine")
    .addItem("Insert Dotted Line", "insertDottedLine")
    .addItem("Insert Thick Block Line", "insertThickBlockLine")
    .addToUi();
}

// Function to count words in selected text
function countSelectedWords() {
  var doc = DocumentApp.getActiveDocument();
  var selection = doc.getSelection();

  if (!selection) {
    DocumentApp.getUi().alert("No text selected! Please select some text.");
    return;
  }

  var totalWords = 0;

  selection.getRangeElements().forEach(function(element) {
    if (element.getElement().asText()) {
      var text = element.getElement().asText().getText();
      var selectedText = text.substring(element.getStartOffset(), element.getEndOffsetInclusive() + 1);
      totalWords += selectedText.trim().split(/\s+/).filter(word => word.length > 0).length;
    }
  });

  DocumentApp.getUi().alert("Word Count: " + totalWords);
}

// Function to count characters in selected text
function countSelectedCharacters() {
  var doc = DocumentApp.getActiveDocument();
  var selection = doc.getSelection();

  if (!selection) {
    DocumentApp.getUi().alert("No text selected! Please select some text.");
    return;
  }

  var totalCharacters = 0;

  selection.getRangeElements().forEach(function(element) {
    if (element.getElement().asText()) {
      var text = element.getElement().asText().getText();
      var selectedText = text.substring(element.getStartOffset(), element.getEndOffsetInclusive() + 1);
      totalCharacters += selectedText.length;
    }
  });

  DocumentApp.getUi().alert("Character Count: " + totalCharacters);
}

// Function to insert lines at cursor
function insertLineAtCursor(lineChar, lineWidth, makeBold) {
  var doc = DocumentApp.getActiveDocument();
  var cursor = doc.getCursor();

  if (!cursor) {
    DocumentApp.getUi().alert("Place the cursor where you want to insert a line.");
    return;
  }

  var fullLine = lineChar.repeat(lineWidth).slice(0, lineWidth);
  var element = cursor.getSurroundingText();
  var textOffset = cursor.getSurroundingTextOffset();

  if (element && element.editAsText) {
    var textElement = element.editAsText();
    textElement.insertText(textOffset, fullLine);

    var maxBoldIndex = Math.min(textOffset + lineWidth, textElement.getText().length);
    textElement.setBold(textOffset, maxBoldIndex - 1, makeBold);
  } else {
    var body = doc.getBody();
    var paragraph = body.insertParagraph(0, fullLine);
    paragraph.setBold(makeBold);
  }
}

// Line Variations
function insertSolidLine() { insertLineAtCursor("_", 76, true); }
function insertThinLine() { insertLineAtCursor("_", 76, false); }
function insertDashedLine() { insertLineAtCursor("-", 127, false); }
function insertDottedLine() { insertLineAtCursor(".", 153, false); }
function insertThickBlockLine() { insertLineAtCursor("█", 60, true); }
