function GmailtoDrive() {
  
  // navigate the attachment
  var threads = GmailApp.search('from:abc.company@gmail.com'); // return all threads from a specific email address
  var message = threads[0].getMessages()[0]; // return the first message from the first thread
  var attachment = message.getAttachments()[0]; // return the first attachment
  
  // create a copy of the attachment
  var xlsxBlob = attachment.setContentType(MimeType.MICROSOFT_EXCEL); // create a blob to store the content
  var target_file = DriveApp.createFile(xlsxBlob); // create a copy from the blob
  var date = Utilities.formatDate(new Date(), "GMT+7", "dd-MM-yyyy"); 
  target_file.setName("ABC Company " + date + ".xlsx"); // set name by the current date

  // get file name from a Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Sheet 1");
  var folder_name = sheet.getRange("A1").getValue(); // cell A1 contains the file name
  
  // get an iteration of folders namely folder_name in folder by id
  var folders = DriveApp.getFolderById("folder_id").getFoldersByName(folder_name);

  if (folders.hasNext()) {
    var target_folder = folders.next();
  }
  else {
    var target_folder = DriveApp.getFolderById("1NxgZF260akFIOJAeXLngiXY_0jW_1WXu").createFolder(folder_name);
  }
  // ---------------------------------------------------------------------------------------------------------
  target_file.moveTo(target_folder); 
  
  var count = 0;
  var files = target_folder.getFilesByName(target_file.getName());
  while (files.hasNext()) {
    count++;
    file = files.next();
  }
  if (count != 1) {
    target_file.setTrashed(true);
    toast("The lastest file was updated.");
  }
  else {
    toast("The update has been done.");
  }
}
