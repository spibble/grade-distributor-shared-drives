//// Google add-on methods

function onOpen(e) {
  SpreadsheetApp.getUi().createAddonMenu()
    .addItem('Distribute grades', 'showDistributeGradesSidebar')
    .addItem('Create class', 'showCreateClassSidebar')
    .addItem('Add students to class', 'showAddStudentsSidebar')
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

function showCreateClassSidebar() {
  showSidebar('Create Class', 'CreateClassSidebar');
}

function showAddStudentsSidebar() {
  showSidebar('Add Students', 'AddStudentsSidebar');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent()
}


//// Error and alert handling

function error(s) {
  SpreadsheetApp.getUi().alert(s);
  throw s;
}

function toast(msg) {
  SpreadsheetApp.getActiveSpreadsheet().toast(msg);
  console.log(msg);
}

