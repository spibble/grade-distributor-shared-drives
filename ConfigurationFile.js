//// Creating and accessing configuration file.

const CONFIGURATION_FILE_NAME = 'Grade Distributor Configuration';
const CONFIGURATION_HEADER_TEXT = 'This file contains information needed by the Grade Distributor add-on. Do not edit it.';

function createConfigurationSheet(thisSpreadsheet, studentsFolderId, prefix, suffix) {
  const earlierConfigurationFile = getConfigurationFile(thisSpreadsheet)
  if (earlierConfigurationFile != null) {
    throw ('There is already a file named "' + CONFIGURATION_FILE_NAME + '" , which means Grade Distributor has already done the setup process in this or a parent directory. If the setup was unsuccessful, delete that file and try again.');
  }
  const folder = DriveApp.getFileById(thisSpreadsheet.getId()).getParents().next();

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

function getConfiguration(spreadsheet) {
  const configurationFile = getConfigurationFile(spreadsheet);
  if (configurationFile === null) {
    throw ('Unable to find file named "' + CONFIGURATION_FILE_NAME + '". Grade distribution cannot proceed.');
  }
  return getConfigurationFromFile(configurationFile);
}

function getConfigurationFromFile(configurationFile) {
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
    var newConfigFile = getFileByName(candidate, CONFIGURATION_FILE_NAME, true);
    if (newConfigFile !== null) {
      if (configurationFile === null) {
        configurationFile = newConfigFile;
      } else {
        throw ('This spreadsheet is in two different folders that contain files named "' + CONFIGURATION_FILE_NAME + '". Grade distribution cannot proceed.');
      }
    }
  }
  return configurationFile;
}