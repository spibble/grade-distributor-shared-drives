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

// throws error if folder already exists
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