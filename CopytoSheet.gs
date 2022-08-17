function CopytoSheet() {
  
  // 1- Navigate the file to be converted 
  let spreadsheet = SpreadsheetApp.openById("spreadsheet_id");
  let currentSpreadsheet = SpreadsheetApp.getActive();
  let sheet = currentSpreadsheet.getSheetByName("Working_Sheet"); // Working_sheet is where to enter input, e.g. file name and run function to retrieve an output
  let fileName = sheet.getRange('A1').getValue(); // cell A1 contains the file name
  
  // 2- Create a Copy gs of the excel file
  let spreadsheetId = convertExcelToGoogleSheets(fileName);
  
  // 3- Import data from the copy gs into Working_Sheet
  let importedSheetName = importDataFromSpreadsheet(spreadsheetId);
}

// Sub-procedures as follows

function convertExcelToGoogleSheets(fileName) {
  let files = DriveApp.getFilesByName(fileName);
  let excelFile = null;
  if(files.hasNext())
    excelFile = files.next();
  else
    return null;
  let blob = excelFile.getBlob();
  let config = {
    title: "[Google Sheets] " + excelFile.getName(),
    parents: [{id: excelFile.getParents().next().getId()}],
    mimeType: MimeType.GOOGLE_SHEETS 
  };
  let spreadsheet = Drive.Files.insert(config, blob);
  return spreadsheet.id; // return the copy gs id
}

function importDataFromSpreadsheet(spreadsheetId) {
  let spreadsheet = SpreadsheetApp.openById(spreadsheetId); // open the copy gs
  let currentSpreadsheet = SpreadsheetApp.getActive(); // get active for the Working_Sheet
  let des_sheet = currentSpreadsheet.getSheetByName("Des_Sheet"); // Des_Sheet is where to import the data into
  des_sheet.clearContents(); // make sure there are no unnecessary rows and columns
  let dataToImport = spreadsheet.getDataRange(); // return the range to get values from the copy Spreadsheet
  let range = des_sheet.getRange(1,1,dataToImport.getNumRows(), dataToImport.getNumColumns()); // return the range to paste the data
  range.setValues(dataToImport.getValues()); // paste
  DriveApp.getFileById(spreadsheet.getId()).setTrashed(true); // move the copy gs into trash
}
