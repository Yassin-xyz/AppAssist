/**
* Finds the index of a column in a given sheet.
* @param {string} sheetName - Name of the sheet to search in.
* @param {number} [headerRow=1] - Row where the header is located (default: row 1).
* @param {string} value - Value to search for in the header.
* @return {number|null} - Column index (0-based) if found, otherwise null.
*/
function findColumnIndex(sheetName, headerRow, value) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      Browser.msgBox("Error: Sheet '" + sheetName + "' does not exist.");
      return null;
    }
    
    headerRow = headerRow || 1;
    var range = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn());
    var headers = range.getValues()[0];
    
    for (var i = 0; i < headers.length; i++) {
      if (headers[i] === value) {
        return i;
      }
    }
    
    Browser.msgBox("Error: Value '" + value + "' not found in sheet '" + sheetName + "'.");
    return null;
    }