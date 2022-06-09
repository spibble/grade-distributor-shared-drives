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
  showSidebar('Set Up Class', 'SetupSidebar');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent()
}

const CONFIGURATION_FILE_NAME = 'Grade Distributor Configuration';


//// Error and alert handling

function error(s) {
  SpreadsheetApp.getUi().alert(s);
  throw s;
}

function toast(msg) {
  SpreadsheetApp.getActiveSpreadsheet().toast(msg);
  console.log(msg);
}

