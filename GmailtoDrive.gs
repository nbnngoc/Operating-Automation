// Store daily .xlsx files sent from your partner/customer in specific Drive folder(s)
// Automate with time-driven trigger, e.g. daily

function GmailtoDrive() {
  
  // 1- Navigate the attachment
  var threads = GmailApp.search('from:abc.company@gmail.com'); // return all threads from a specific email address
  var message = threads[0].getMessages()[0]; // return the first message from the first thread
  var attachment = message.getAttachments()[0]; // return the first attachment in that first message
  
  // 2- Create a copy of the attachment
  var xlsxBlob = attachment.setContentType(MimeType.MICROSOFT_EXCEL); // create a blob to store the content
  var target_file = DriveApp.createFile(xlsxBlob); // create a copy file from the blob
  var date = Utilities.formatDate(new Date(), "GMT+7", "dd-MM-yyyy"); 
  target_file.setName(date); // set name for the copy file, e.g. containing the current date

  // 3- Decide which folder is to stored the attachment
  ///// 3a- Get name of the destination folder, e.g. from a spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var sheet = ss.getSheetByName("Sheet");
  var folder_name = sheet.getRange("A1").getValue(); // cell A1 contains the file name (might contain a formula so it's changeable, e.g. the current month)
  
  ///// 3b- Check if the destination folder has existed or not
  var folders = DriveApp.getFolderById("parent_folder_id").getFoldersByName(folder_name); // e.g. there are many destination folders differentiated by month, so that the parent folder is to store all of them
  if (folders.hasNext()) { 
    var des_folder = folders.next(); // if true, destination folder is the lastest one
  }
  else {
    var des_folder = DriveApp.getFolderById("parent_folder_id").createFolder(folder_name); // if false, create new destination folder named after folder_name
  }
  
  // 4- Move the attachment to the destination folder
  target_file.moveTo(des_folder);
  
  // 5- Remove the duplicate
  var count = 0;
  var files = des_folder.getFilesByName(target_file.getName()); // return the attachment name
  while (files.hasNext()) {
    count++; // count no of files with the name returned in line 37
    file = files.next();
  }
  if (count != 1) { // is there any duplicate?
    target_file.setTrashed(true); // if true, remove the attachment
    toast("The lastest file was updated.");
  }
  else {
    toast("The update has been done."); // if false, do nothing
  }
}
