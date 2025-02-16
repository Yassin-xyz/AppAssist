/**
 * Function to duplicate a Google Sheets tab into a new Google Sheets file.
 * @param {string} sheetName - The name of the sheet to duplicate.
 * @param {string} folderId - The ID of the Google Drive folder where the new file will be stored.
 * @param {string} newSheetName - The name of the new Google Sheets file.
 */
function duplicateGSheet(sheetName, folderId, newSheetName) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
  
    if (!sheet) {
        Logger.log("❌ The sheet '" + sheetName + "' does not exist.");
        return;
    }
  
    // 1️⃣ Create a temporary sheet in the current file
    var tempSheet = ss.insertSheet(sheetName + "_temp");
    Logger.log("📌 Temporary sheet created: " + tempSheet.getName());
  
    // 2️⃣ Copy values **as text** and formats
    var range = sheet.getDataRange();
    var values = range.getValues(); // Retrieve values
    var formats = range.getNumberFormats(); // Retrieve numeric formats
    
    // Convert all values to text (prevents 004 → 4 conversion)
    for (var i = 0; i < values.length; i++) {
        for (var j = 0; j < values[i].length; j++) {
            values[i][j] = values[i][j].toString(); // Convert each value to text
        }
    }
  
    var targetRange = tempSheet.getRange(1, 1, values.length, values[0].length);
    targetRange.setValues(values); // Copy converted text values
    targetRange.setNumberFormats(formats); // Copy numeric formats
    
    // Copy formatting (colors, borders, etc.)
    range.copyFormatToRange(tempSheet, 1, values[0].length, 1, values.length);
  
    // Copy column widths and row heights
    for (var i = 1; i <= sheet.getLastColumn(); i++) {
        tempSheet.setColumnWidth(i, sheet.getColumnWidth(i));
    }
    for (var j = 1; j <= sheet.getLastRow(); j++) {
        tempSheet.setRowHeight(j, sheet.getRowHeight(j));
    }
  
    Logger.log("✅ Data and formatting copied to the temporary sheet.");
  
    // 3️⃣ Create a new Google Sheets file
    var newFile = SpreadsheetApp.create(newSheetName);
    var newSs = SpreadsheetApp.openById(newFile.getId());
  
    // 4️⃣ Copy the temporary sheet to the new file
    var copiedSheet = tempSheet.copyTo(newSs);
    copiedSheet.setName(sheetName); // Rename the copied sheet
    Logger.log("📌 The sheet has been copied to the new file.");
  
    // Delete the default sheet in the new file
    var defaultSheet = newSs.getSheets()[0];
    if (defaultSheet.getSheetId() !== copiedSheet.getSheetId()) {
        newSs.deleteSheet(defaultSheet);
    }
  
    // ✅ Remove the temporary sheet after copying
    try {
        ss.deleteSheet(tempSheet);
        Logger.log("🗑️ Temporary sheet '" + tempSheet.getName() + "' successfully deleted.");
    } catch (e) {
        Logger.log("⚠️ Error while deleting the temporary sheet: " + e.message);
    }
  
    // 5️⃣ Move the new file to the specified folder
    var folder = DriveApp.getFolderById(folderId);
    var file = DriveApp.getFileById(newFile.getId());
    folder.addFile(file);
    DriveApp.getRootFolder().removeFile(file); // Remove from "My Drive"
  
    Logger.log("✅ Final file created: " + newSheetName + " and moved to folder ID " + folderId);
}