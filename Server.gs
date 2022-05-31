function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent()
}

/*
function onOpenSidebar() {

  var ui = HtmlService.createTemplateFromFile('Sidebar')
    .evaluate()
    .setTitle('Materialize CSS Sidebar Example')
  
  SpreadsheetApp.getUi().showSidebar(ui)
}
*/

/*
 * Enables teachers to share with students only the grading spreadsheet columns relevant to each of them.
 * This add-on enables a teacher to grade all of their students in a single spreadsheet, with one row per 
 * problem and one column per student, and to make child spreadsheets showing each student only the header 
 * columns and column with their grade.
 */


const CONFIGURATION_FILE_NAME = 'Grade Distributor Configuration';

// The name of the folder holding students' folders.
var STUDENTS_FOLDER_NAME = 'Students';

// The name of the file holding the prefix to student folders, such as 'CS115-'.
var PREFIX_FILE_NAME = 'PREFIX';

// The name of the file holding the suffix to student folders, such as '-CS115'.
var SUFFIX_FILE_NAME = 'SUFFIX';


//// Google add-on methods

function onOpen(e) {
  SpreadsheetApp.getUi().createAddonMenu()
    .addItem('Distribute grades', 'showDistributeGradesSidebar')
    .addItem('Set up class', 'showSetupClassSidebar')
    .addToUi();
}

function onInstall(e) {
  onOpen(e);
}

function showSidebar(title, filename) {
  var ui = HtmlService.createTemplateFromFile(filename)
    .evaluate()
    .setTitle(title);
  SpreadsheetApp.getUi().showSidebar(ui);
}

function showDistributeGradesSidebar() {
  showSidebar('Distribute Grades', 'DistributeGradesSidebar');
}

function showSetupClassSidebar() {
  showSidebar('Set Up Class', 'setup_class');
}


//// Error and alert handling

function error(s) {
  SpreadsheetApp.getUi().alert(s);
  throw s;
}

function toast(msg) {
  SpreadsheetApp.getActiveSpreadsheet().toast(msg);
  //console.log(msg);
}


//// Google Drive helper methods

function getEnclosingFolder(folder) {
  var enclosingFolders = folder.getParents();
  var enclosingFolder;
  if (enclosingFolders.hasNext()) {
    enclosingFolder = enclosingFolders.next();
    if (enclosingFolders.hasNext()) {
      error('Folder has more than one enclosing folder: ' + folder);
    }
  } else {
    error('Folder has no enclosing folder: ' + folder);
  }
  return enclosingFolder;
}

// Gets the folder that holds the student folders
// by searching for a folder with the appropriate name.
function getStudentsFolder(folder) {
  for (; ;) {
    candidates = folder.getFoldersByName(STUDENTS_FOLDER_NAME);
    if (candidates.hasNext()) {
      var studentsFolder = candidates.next();
      if (candidates.hasNext()) {
        error('Too many folders in ' + folder + ' named ' + STUDENTS_FOLDER_NAME);
      }
      return studentsFolder;
    }
    folder = getEnclosingFolder(folder);
  }
}

function getFileByName(folder, fileName) {
  const files = folder.getFilesByName(fileName);
  if (files.hasNext()) {
    var file = files.next();
    if (files.hasNext()) {
      throw ('There are too many files named "' + fileName + '".');
    }
    return file;
  }
  return null;
}

function getAffix(fileName, studentFolder) {
  var files = studentFolder.getFilesByName(fileName);
  if (files.hasNext()) {
    var file = files.next();
    var body = DocumentApp.openById(file.getId()).getBody().getText();
    return body.trim();
  } else {
    return '';
  }
}

function getStudentFolder(studentsFolder, name) {
  var folders = studentsFolder.getFoldersByName(name);
  if (folders.hasNext()) {
    folder = folders.next();
    if (folders.hasNext()) {
      error('Too many folders found named ' + name);
    }
    return folder;
  }
  error('Folder not found: ' + name);
}

function deleteOldSpreadsheet(studentFolder, fileName) {
  const files = studentFolder.getFilesByName(fileName);
  if (files.hasNext()) {
    Drive.Files.remove(files.next().getId());
  }
}

// This creates a new spreadsheet and returns the sheet it contains.
function createNewSheet(studentFolder, fileName, studentName) {
  deleteOldSpreadsheet(studentFolder, fileName);
  const newSpreadsheet = createNewSpreadsheet(fileName, studentFolder.getId());
  const newSheet = newSpreadsheet.getActiveSheet();
  toast('Made sheet for ' + studentName);
  return newSheet;
}

function createNewSpreadsheet(spreadsheetName, folderId) {
  // http://stackoverflow.com/a/41509877/631051
  var resource = {
    title: spreadsheetName,
    mimeType: MimeType.GOOGLE_SHEETS,
    parents: [{ id: folderId }]
  };
  var newSpreadsheetFile = Drive.Files.insert(resource);
  return SpreadsheetApp.openById(newSpreadsheetFile.id);
}

function createFolderIfNotPresent(parentFolder, newFolderName) {
  const oldFoldersIterator = parentFolder.getFoldersByName(newFolderName);
  while (oldFoldersIterator.hasNext()) {
    if (!oldFoldersIterator.next().isTrashed()) {
      throw ('There is already a folder named "' + newFolderName + '"');
    }
  }
  return parentFolder.createFolder(newFolderName);
}

// https://stackoverflow.com/a/29647047/631051
// https://developers.google.com/drive/api/v2/reference/permissions?hl=en
function addUserPermission(fileId, email, writePermission, notify) {
  var request = Drive.Permissions.insert({
    'value': email,
    'type': 'user',
    'role': (writePermission ? 'writer' : 'reader'),
    'withLink': false
  },
    fileId,
    {
      'sendNotificationEmails': notify
    });
}



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


/// JavaScript helper functions

function isNumber(n) {
  // https://stackoverflow.com/a/1421988/631051
  return !isNaN(parseFloat(n)) && !isNaN(n - 0)
}

function ensureValuesUnique(values) {
  const valuesSet = new Set();  // requires V8 engine
  for (var i = 0; i < values.length; i++) {
    var value = values[i].toString().trim();
    if (valuesSet.has(value)) {
      throw ('Duplicate value: ' + value);
    }
    valuesSet.add(value);
  }
}


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

function getConfiguration(spreadsheet) {
  const configurationFile = getConfigurationFile(spreadsheet);
  const configSpreadsheet = SpreadsheetApp.openById(configurationFile.getId());
  const sheet = configSpreadsheet.getActiveSheet();
  // Skip top row to get column with folder id, prefix, and suffix
  const values = sheet.getRange(2, 1, 4, 1).getValues();
  const studentsFolderId = values[0][0].toString().trim();
  const prefix = values[1][0].toString().trim();
  const suffix = values[2][0].toString().trim();
  if (studentsFolderId === '') {
    throw ('File named "' + '" is corrupted. Grade distribution cannot proceed.');
  }
  return [studentsFolderId, prefix, suffix];
}

function getConfigurationFile(spreadsheet) {
  const parentFolders = DriveApp.getFileById(spreadsheet.getId()).getParents();
  var configurationFile = null;
  if (parentFolders.hasNext()) {
    var candidate = parentFolders.next();
    var newConfigFile = getFileByName(candidate, CONFIGURATION_FILE_NAME);
    if (newConfigFile !== null) {
      if (configurationFile === null) {
        configurationFile = newConfigFile;
      } else {
        throw ('This spreadsheet is in two different folders that contain files named "' + CONFIGURATION_FILE_NAME + '". Grade distribution cannot proceed.');
      }
    }
  }
  if (configurationFile === null) {
    throw ('Unable to find file in this folder named "' + CONFIGURATION_FILE_NAME + '". Grade distribution cannot proceed.');
  }
  return configurationFile;
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

// Copy contents, formatting, and optionally notes.
// This does not copy the font formats of numeric cells because getRichTextValues() doesn't.
function copyContents(sourceRange, destRange, copyNotes) {
  destRange.setNumberFormats(getNumberFormats(sourceRange));
  destRange.setRichTextValues(sourceRange.getRichTextValues());
  destRange.setValues(sourceRange.getValues());

  if (copyNotes) {
    destRange.setNotes(sourceRange.getNotes());
  }
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


//// Business logic for Setup

// Not const because of Apps Script issue. https://stackoverflow.com/a/54413892/631051
const NAME_COLUMN_NUMBER = 1;
const NAME_COLUMN_NAME = 'name';
const NAME_ERROR_MESSAGE = 'The top cell in the first column must be "Name".';
const EMAIL_COLUMN_NUMBER = 2;
const EMAIL_COLUMN_NAME = 'email';
const EMAIL_ERROR_MESSAGE = 'The top cell in the second column must be "Email".';
const URL_COLUMN_NUMBER = 3;
const URL_ERROR_MESSAGE = 'The third column should be empty.';
const CONFIGURATION_HEADER_TEXT = 'This file contains information needed by the Grade Distributor add-on. Do not edit it.';

// returns [names, emails] if good; otherwise throws exception
function validateStudentSetupSheet() {
  const thisSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  checkColumnHeader(thisSheet, NAME_COLUMN_NUMBER, NAME_COLUMN_NAME, NAME_ERROR_MESSAGE);
  checkColumnHeader(thisSheet, EMAIL_COLUMN_NUMBER, EMAIL_COLUMN_NAME, EMAIL_ERROR_MESSAGE);

  // Ensure same number of names and email addresses.
  const maxNameRow = findBlankInColumn(thisSheet, NAME_COLUMN_NUMBER) - 1;
  const maxEmailRow = findBlankInColumn(thisSheet, EMAIL_COLUMN_NUMBER) - 1;
  if (maxNameRow > maxEmailRow) {
    throw ('You have more names (' + (maxNameRow - 1) + ') than email addresses (' + (maxEmailRow - 1) + ').');
  }
  if (maxNameRow < maxEmailRow) {
    throw ('You have more email addresses (' + (maxEmailRow - 1) + ') than email addresses (' + (maxNameRow - 1) + ').');
  }

  const names = getNames(thisSheet, maxNameRow);
  const emails = getEmails(thisSheet, maxEmailRow);
  ensureValuesUnique(names);
  ensureValuesUnique(emails);
  if (!isColumnBlank(thisSheet, maxNameRow, URL_COLUMN_NUMBER)) {
    throw (URL_ERROR_MESSAGE);
  }

  return [names, emails];
}

function getNames(sheet, maxRow) {
  return getColumnValuesWithoutHeading(sheet, NAME_COLUMN_NUMBER, maxRow);
}

function getEmails(sheet, maxRow) {
  return getColumnValuesWithoutHeading(sheet, EMAIL_COLUMN_NUMBER, maxRow);
}

function createStudentSheets(topFolderName, prefix, suffix, giveEditAccess, emailStudent) {
  const thisSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const thisSheet = thisSpreadsheet.getActiveSheet();
  const currentFolder = DriveApp.getFileById(thisSpreadsheet.getId()).getParents().next();

  // Check if configuration file already exists.
  if (currentFolder.getFilesByName(CONFIGURATION_FILE_NAME).hasNext()) {
    throw ('There is already a file named "' + CONFIGURATION_FILE_NAME + '", which means Grade Distributor has already done the setup process in this directory. If the setup was unsuccessful, delete that file and try again.');
  }

  // Double-check that data is valid and get names and emails.
  const arr = validateStudentSetupSheet();
  const names = arr[0];
  const emails = arr[1];

  // Create folder for all students.
  const studentsFolder = createFolderIfNotPresent(currentFolder, topFolderName);

  // Create student folders, adding the URLs to the spreadsheet.
  const numStudents = names.length;
  const urlRange = thisSheet.getRange(1, URL_COLUMN_NUMBER, names.length + 1, 1);
  urlRange.getCell(1, 1).setValue('Folder URL');
  for (var i = 0; i < numStudents; i++) {
    var name = prefix + names[i].toString().trim() + suffix;
    var childFolder = studentsFolder.createFolder(name);
    addUserPermission(childFolder.getId(), emails[i].toString(), giveEditAccess, emailStudent);
    urlRange.getCell(i + 2, 1).setValue(childFolder.getUrl());
  }

  createConfigurationSheet(currentFolder, studentsFolder.getId(), prefix, suffix);

  return numStudents;
}

function createConfigurationSheet(folder, studentsFolderId, prefix, suffix) {
  const spreadsheet = createNewSpreadsheet(CONFIGURATION_FILE_NAME, folder.getId());
  const sheet = spreadsheet.getActiveSheet();
  const vals = [
    [CONFIGURATION_HEADER_TEXT],
    [studentsFolderId.toString()],
    [prefix.toString().trim()],
    [suffix.toString().trim()]
  ];
  const range = sheet.getRange(1, 1, vals.length, 1);
  const textFormats = [['@'], ['@'], ['@'], ['@']];
  range.setValues(vals);
  range.setNumberFormats(textFormats);
}
