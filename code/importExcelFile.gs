/** 
* 
* Import an Excel spreadsheet into your Google Sheet with Drive API.
* Activate before using: Resources -> Advanced Google Services... -> Turn Drive API on.
*
* id {string} The ID of a file saved in your Google Drive (can be grabbed from the share URL of the file or manually with the Picker)
*
* References
* https://developers.google.com/apps-script/reference/document/document-app#getui
*/

function importExcelFile(id){
  
  var id = id;
  var currentDate = new Date();
  var file = DriveApp.getFileById(id);  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet();
  
  // Is the attachment an Excel file?  
  if (file.getMimeType() == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"){
        
    //convert to Google Sheet
    var convertedFile = {
      title: file.getName()
      //    parents: [{ id: "Insert the file ID of your preferred parent folder if you have one" }]
    };
    convertedFile = Drive.Files.insert(convertedFile,file, {
      convert:true
    });
    
    // Grab the converted Google Sheet
    var SSSheets = SpreadsheetApp.openById(convertedFile.id);
    // Get full range of data
    var sheetRange = SSSheets.getDataRange();
    
    // Grab the current formats from therange
    var fontFamilies = sheetRange.getFontFamilies();
    var fontColors = sheetRange.getFontColors();
    var fontLines = sheetRange.getFontLines();
    var fontSizes = sheetRange.getFontSizes();  
    var fontStyles = sheetRange.getFontStyles();
    var fontWeights = sheetRange.getFontWeights();
    var hAlignments = sheetRange.getHorizontalAlignments();
    var vAlignments = sheetRange.getVerticalAlignments();
    var backgroundColors = sheetRange.getBackgrounds();
    var numberFormats = sheetRange.getNumberFormats();  
    
    // Set the data values and formatting in range
    var sheetData = sheetRange.getValues();
    sheet.clear();
    sheet.getRange(1, 1, SSSheets.getLastRow(), SSSheets.getLastColumn())
    .setValues(sheetData)
    .setFontFamilies(fontFamilies)
    .setFontColors(fontColors)
    .setFontLines(fontLines)
    .setFontSizes(fontSizes)  
    .setFontStyles(fontStyles)
    .setFontWeights(fontWeights)
    .setHorizontalAlignments(hAlignments)
    .setVerticalAlignments(vAlignments)
    .setBackgrounds(backgroundColors)
    .setNumberFormats(numberFormats);     
    
    // Update A1 with a note with some pertinent data, optional but I like it
    sheet.getRange("A1")
    .setNote("Original Excel file location: " + file.getUrl() + " | Converted Google Sheet location: " + SSSheets.getUrl() + " | Imported " + currentDate);
        
    // Update sheet name
    sheet.setName(SSSheets.getName());
    
  } else {
    Logger.log("Not an Excel file");
  }
}