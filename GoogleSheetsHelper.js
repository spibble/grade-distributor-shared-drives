//// Google Sheets helper functions

// Get [colLeft, colRight, colRight + 1], using letters for columns (e.g.,['A', 'C', 'D']).
function getSelectionCols() {
  const range = SpreadsheetApp.getActiveRange();
  if (!range) return [0, 0];
  const regex = /[A-Z]+/g;
  const matches1 = range.getA1Notation().match(regex);
  const colLeft = matches1[0];
  const colRight = matches1[1];
  const colRightNumeric = range.getColumn() + range.getNumColumns();
  const colFirstStudent = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(1, colRightNumeric, 1, 1).getA1Notation().match(regex)[0];
  return [colLeft, colRight, colFirstStudent];
}

function toColumnNumber(colName) {
  const thisSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  return thisSheet.getRange(colName + '1').getColumn();
}

function getSelectionHeight() {
  var range = SpreadsheetApp.getActiveRange();
  return range ? range.getHeight() : 0;
}

// Throw an error if the header of the specified column doesn't match the name.
function checkColumnHeader(sheet, colNumber, name, error) {
  if (sheet.getRange(1, colNumber).getValue().toString().toLowerCase() !== name) {
    throw (error);
  }
}

// Return the first row in the specified column with a blank entry.
function findBlankInColumn(sheet, colNumber) {
  for (var row = 1; ; row++) {
    if (sheet.getRange(row, colNumber).getValue().toString().trim() == '') {
      return row;
    }
  }
}

function getColumnValuesWithoutHeading(sheet, colNumber, maxRow) {
  return sheet.getRange(2, colNumber, maxRow - 1, 1).getValues();
}

function isColumnBlank(sheet, numRows, colNumber) {
  const values = sheet.getRange(1, colNumber, numRows, 1).getValues();
  for (var i = 0; i < numRows; i++) {
    if (values[i].toString().trim() !== '') {
      return false;
    }
  }
  return true;
}

// I observed that sometimes number format was empty string for formulas,
// which caused the value not to be displayed. This fixes that.
function getNumberFormats(range) {
  const REPLACEMENT_NUMBER_FORMAT = '0.###############';
  const formats = range.getNumberFormats();
  for (var i = 0; i < formats.length; i++) {
    for (var j = 0; j < formats[i].length; j++) {
      if (formats[i][j].toString() === '') {
        formats[i][j] = REPLACEMENT_NUMBER_FORMAT;
      }
    }
  }
  return formats;
}

// Copy contents, formatting, and optionally notes.
// This does not copy the font formats of numeric cells because getRichTextValues() doesn't.
function copyContents(sourceRange, destRange, copyNotes) {
  destRange.setNumberFormats(getNumberFormats(sourceRange));
  destRange.setRichTextValues(sourceRange.getRichTextValues());
  destRange.setValues(sourceRange.getValues());

  // additional formatting stuff
  destRange.setBackgrounds(sourceRange.getBackgrounds());
  destRange.setFontColors(sourceRange.getFontColors());
  destRange.setFontWeights(sourceRange.getFontWeights());
  destRange.setFontStyles(sourceRange.getFontStyles());
  destRange.setFontFamilies(sourceRange.getFontFamilies());
  destRange.setFontSizes(sourceRange.getFontSizes());
  destRange.setHorizontalAlignments(sourceRange.getHorizontalAlignments());
  destRange.setVerticalAlignments(sourceRange.getVerticalAlignments());

  if (copyNotes) {
    destRange.setNotes(sourceRange.getNotes());
  }
}
