function onOpen() {
  SpreadsheetApp
    .getUi()
    .createMenu('Sidebar')
    .addItem('Open', 'onOpenSidebar')
    .addToUi()
}

function onOpenSidebar() {

  var ui = HtmlService.createTemplateFromFile('Sidebar')
    .evaluate()
    .setTitle('Materialize CSS Sidebar Example')
  
  SpreadsheetApp.getUi().showSidebar(ui)
}

// Google Sheets helper functions
// ------------------------------

function getSelectionHeight() {
  var range = SpreadsheetApp.getActiveRange();
  return range ? range.getHeight() : 0;
}

// Get [colLeft, colRight, colRight + 1], using letters for columns (e.g.,['A', 'C', 'D']).
function getSelectionCols() {
  var range = SpreadsheetApp.getActiveRange();
  if (!range) return ['', '', ''];
  var regex = /[A-Z]+/g;
  var matches1 = range.getA1Notation().match(regex);
  var colLeft = matches1[0];
  var colRight = matches1[1];
  var colRightNumeric = range.getColumn() + range.getNumColumns();
  var colFirstStudent = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(1, colRightNumeric, 1, 1).getA1Notation().match(regex)[0];
  return [colLeft, colRight, colFirstStudent];
}

function onFormSubmit(form) {

  var spreadsheet = SpreadsheetApp.getActive()
  var sheet = getSheet(spreadsheet) 
  var photoUrl = (form.filepath) ? DriveApp.createFile(form.photo).getUrl() : ''
  
  sheet.appendRow([
    new Date(),
    form.first_name,
    form.last_name,
    form.password,
    form.email,
    photoUrl
  ])

  
  // Private Functions
  // -----------------
  
  function getSheet(spreadsheet) {
    
    var sheet = spreadsheet.getSheetByName('Results')  
    
    if (sheet === null) {
      sheet = spreadsheet.insertSheet().setName('Results')
    }
    
    if (sheet.getRange('A1').getValue() === '') {
    
      sheet
        .getRange('A1:F1')
        .setValues([['Timestamp', 'First Name', 'Last Name', 'Password', 'Email', 'Photo']])
        
      sheet.setFrozenRows(1)
    }
    
    return sheet
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent()
}