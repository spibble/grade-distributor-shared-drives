//// Business logic for Grade Distribution

// Top-level function called from sidebar.
function shareGrades(firstPublicColumnLetter, lastPublicColumnLetter, maxRow,
  firstStudentColumnLetter, copyPublicNotes, copyStudentNotes) {
  const thisSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const thisSheet = thisSpreadsheet.getActiveSheet();
  const name = thisSpreadsheet.getName();
  const configuration = getConfiguration(thisSpreadsheet);
  const studentsFolder = DriveApp.getFolderById(configuration[0]);
  const prefix = configuration[1];
  const suffix = configuration[2];

  const firstPublicColumnNumber = toColumnNumber(firstPublicColumnLetter);
  const lastPublicColumnNumber = toColumnNumber(lastPublicColumnLetter);

  // Process students until an empty column heading is reached.
  var numSuccesses = 0
  var numFailures = 0
  var colNum = toColumnNumber(firstStudentColumnLetter);
  while (thisSheet.getRange(1, colNum).getValue() !== '') {
    try {
      makeNewSheet(name, thisSheet, colNum, maxRow, studentsFolder, prefix, suffix, firstPublicColumnNumber, lastPublicColumnNumber, copyPublicNotes, copyStudentNotes);
      numSuccesses++;
    } catch (e) {
      // The typical reason for failure is that it couldn't find a folder for a given student.
      // getStudentFolder() will display an appropriate message.
      console.log(e);
      numFailures++;
    }
    colNum++;
  }
  toast('Done. Successes: ' + numSuccesses + ' Failures: ' + numFailures);
  return numFailures == 0;
}

function makeNewSheet(name, parentSheet, studentColNumber, maxRow, studentsFolder, prefix, suffix, firstPublicColNumber, lastPublicColNumber, copyPublicNotes, copyStudentNotes) {
  const numPublicCols = lastPublicColNumber - firstPublicColNumber + 1;
  const parentStudentColRange = parentSheet.getRange(1, studentColNumber, maxRow, 1);
  const studentName = parentStudentColRange.getCell(1, 1).getValue();
  const studentFolderName = prefix + studentName + suffix;
  const fileName = studentName + ': ' + name;
  const studentFolder = getStudentFolder(studentsFolder, studentFolderName);
  const childSheet = createNewSheet(studentFolder, fileName, studentName);

  copyPublicColumns(parentSheet, childSheet, maxRow, firstPublicColNumber, lastPublicColNumber, copyPublicNotes);
  createStudentColumn(childSheet, parentStudentColRange, maxRow, numPublicCols, copyStudentNotes);
}

// Copy the public columns (values, formats, and optionally notes) of the current spreadsheet.
function copyPublicColumns(parentSheet, newSheet, maxRow, firstPublicColNumber, lastPublicColNumber, copyPublicNotes) {
  const numPublicCols = lastPublicColNumber - firstPublicColNumber + 1;
  const publicColsRange = parentSheet.getRange(1, firstPublicColNumber, maxRow, numPublicCols)
  const newPublicColsRange = newSheet.getRange(1, 1, maxRow, numPublicCols);

  copyContents(publicColsRange, newPublicColsRange, copyPublicNotes);
}

// Copy the student column and highlight numeric cells that differ from rightmost public column.
function createStudentColumn(childSheet, parentStudentColRange, maxRow, numPublicCols, copyStudentNotes) {
  const childStudentColRange = childSheet.getRange(1, numPublicCols + 1, maxRow, 1);
  copyContents(parentStudentColRange, childStudentColRange, copyStudentNotes);

  // Highlight scores that differ from the maximum (except for Total and Percent).
  // This assumes that the maximum score is in the last public column, and it 
  // excludes the bottom two rows, which are assumed to be the Total and Percent.
  const maxValues = childSheet.getRange(1, numPublicCols, maxRow, 1).getValues();
  for (var row = 1; row <= maxRow - 2; row++) {
    var max = maxValues[row - 1];
    if (isNumber(max)) {
      var actualRange = childStudentColRange.getCell(row, 1);
      var value = actualRange.getValue();
      if (value != max) {
        actualRange.setBackground('yellow');
      }
    }
  }
}
