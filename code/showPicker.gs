// code.gs

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Select an Excel File with the Google Drive Picker')
      .addItem('Choose Excel File', 'showPicker')
      .addToUi();
}

/**
 * Displays an HTML-service dialog in Google Sheets that contains client-side
 * JavaScript code for the Google Picker API.
 *
 * https://www.labnol.org/code/20039-google-picker-with-apps-script
 */
function showPicker() {
  var html = HtmlService.createHtmlOutputFromFile('code/Picker.html')
      .setWidth(600)
      .setHeight(425)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showModalDialog(html, 'Select Folder');
}

function getOAuthToken() {
  DriveApp.getRootFolder();
  return ScriptApp.getOAuthToken();
}