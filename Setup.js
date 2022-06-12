//// Business logic for Setup

const NAME_COLUMN_NUMBER = 1;
const NAME_COLUMN_NAME = 'name';
const NAME_ERROR_MESSAGE = 'The top cell in the first column must be "Name".';
const EMAIL_COLUMN_NUMBER = 2;
const EMAIL_COLUMN_NAME = 'email';
const EMAIL_ERROR_MESSAGE = 'The top cell in the second column must be "Email".';
const URL_COLUMN_NUMBER = 3;
const URL_ERROR_MESSAGE = 'The third column should be empty.';
const CONFIGURATION_HEADER_TEXT = 'This file contains information needed by the Grade Distributor add-on. Do not edit it.';
const IS_VALIDATED_KEY = 'validated';

// returns [names, emails] if good; otherwise throws exception
function validateStudentSetupSheet() {
  setIsValidated(false);
  deleteTriggers();
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

  setIsValidated(true);
  addTrigger();
  return [names, emails];
}

function getIsValidated() {
  return PropertiesService.getScriptProperties().getProperty(IS_VALIDATED_KEY) === 'true';
}

function setIsValidated(value) {
  PropertiesService.getScriptProperties().setProperty(IS_VALIDATED_KEY, value);
}

function deleteTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
}

function addTrigger() {
  deleteTriggers();
  if (ScriptApp.getProjectTriggers().length == 0) {
    ScriptApp.newTrigger("respondToEdit")
      .forSpreadsheet(SpreadsheetApp.getActive())
      .onEdit()
      .create()
  }
}

function respondToEdit(e) {
  setIsValidated(false);
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
    var name = names[i].toString().trim();
    toast('Creating folder for ' + name);
    var childFolderName = prefix + name + suffix;
    var childFolder = studentsFolder.createFolder(childFolderName);
    addUserPermission(childFolder.getId(), emails[i].toString(), giveEditAccess, emailStudent);
    urlRange.getCell(i + 2, 1).setValue(childFolder.getUrl());
  }

  createConfigurationSheet(currentFolder, studentsFolder.getId(), prefix, suffix);

  return numStudents;
}

function createConfigurationSheet(folder, studentsFolderId, prefix, suffix) {
  toast('Creating configuration file');
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
