/**
 * Exports a Google Sheets tab to a CSV file in Google Drive.
 * @param {string} sheetName - The name of the sheet to export.
 * @param {string} folderId - The ID of the Google Drive folder where the CSV will be saved.
 * @param {string} fileName - The name of the output CSV file.
 */
function exportSheetToCSV(sheetName, folderId, fileName) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
  
    if (!sheet) {
        Logger.log("❌ The sheet '" + sheetName + "' does not exist.");
        return;
    }
  
    Logger.log("📌 Starting export of sheet '" + sheetName + "' to CSV.");
  
    // 1️⃣ Retrieve the values as an array
    var range = sheet.getDataRange();
    var values = range.getValues(); // Get raw values
  
    // 2️⃣ Convert values to text (prevents loss of leading zeros and formats dates)
    var csvContent = values.map(row => 
        row.map(cell => {
            if (cell instanceof Date) {
                return `"'${Utilities.formatDate(cell, Session.getScriptTimeZone(), "dd/MM/yyyy")}"`; // Format DD/MM/YYYY as text
            }
            return `"'${cell.toString().replace(/"/g, '""')}"`; // Force text format by adding an apostrophe
        }).join(";")
    ).join("\n");
  
    Logger.log("✅ Data successfully converted to CSV format.");
  
    // 3️⃣ Create the CSV file in Google Drive
    var folder = DriveApp.getFolderById(folderId);
    var file = folder.createFile(fileName + ".csv", csvContent, MimeType.CSV);
    
    Logger.log("✅ CSV file successfully created: " + file.getName());
    Browser.msgBox("✅ CSV file successfully created in the specified folder!");
}
