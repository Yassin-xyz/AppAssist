/**
 * Main function to transcode and copy data
 * @param {string} gsheet_in - Name of the source sheet.
 * @param {string} gsheet_out - Name of the destination sheet.
 * @param {string} transco - Name of the sheet containing the transcoding table.
 * @param {number} row - Row from gsheet_in to copy.
 */
function transcodeAndCopy(gsheet_in, gsheet_out, transco, row) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var sheet_in = ss.getSheetByName(gsheet_in);
  var sheet_out = ss.getSheetByName(gsheet_out);
  var sheet_transco = ss.getSheetByName(transco);
  
  if (!sheet_in || !sheet_out || !sheet_transco) {
    Browser.msgBox("Error: Please ensure all sheets exist!");
    return;
  }

  Logger.log("📌 Starting transcoding from '" + gsheet_in + "' to '" + gsheet_out + "' using '" + transco + "'");

  // 🔍 Step 1: Retrieve the transcoding table
  var transcoData = retrieveTranscoTable(sheet_transco);
  Logger.log("📋 Transcoding table loaded: " + JSON.stringify(transcoData));

  // 🔍 Step 2: Apply transcoding rules
  var transcoMap = applyTranscoRules(transcoData);
  Logger.log("🔍 TranscoMap after processing: " + JSON.stringify(transcoMap));

  // 🔍 Verification: If transcoMap is empty, error
  if (Object.keys(transcoMap).length === 0) {
    Browser.msgBox("❌ ERROR: transcoMap is empty. Check the in/out columns in the transcoding table.");
    return;
  }

  // 🔍 Step 3: Retrieve data
  var rowData = retrieveData(sheet_in, transcoMap, row);
  Logger.log("📥 Data retrieved for row " + row + ": " + JSON.stringify(rowData));

  // 🔍 Step 4: Write data
  Logger.log("✍️ Starting to write data to '" + gsheet_out + "'");
  writeData(sheet_out, rowData, transcoMap, gsheet_out);

  Browser.msgBox("✅ Transcoding and copying completed successfully!");
}

/**
* Retrieves the transcoding table as an array of objects.
* @param {Sheet} sheet - Sheet containing the transcoding table.
* @return {Array} - Array of objects [{in: "ColumnName", out: "Range", split: "Separator"}]
*/
function retrieveTranscoTable(sheet) {
  var data = sheet.getDataRange().getValues();
  var transcoTable = [];

  for (var i = 1; i < data.length; i++) { // Skip header
    transcoTable.push({
      in: data[i][0],  
      out: data[i][1], 
      split: data[i][2] || null 
    });
  }
  return transcoTable;
}

/**
* Processes the transcoding table and handles duplicates (split/concat).
* @param {Array} transcoTable - Array of transcoding rules.
* @return {Object} - Object map with duplication management.
*/
function processTranscoding(transcoTable) {
  var transcoMap = {};

  transcoTable.forEach(function(item) {
    if (!transcoMap[item.in]) {
      transcoMap[item.in] = { out: item.out, split: item.split, values: [] };
    } else {
      if (transcoMap[item.in].out === item.out) {
        transcoMap[item.in].split = item.split; 
      }
    }
  });

  return transcoMap;
}

function applyTranscoRules(transcoTable) {
  var transcoMap = {};

  transcoTable.forEach(function(item) {
    if (!item.in || !item.out) {
      Logger.log("⚠️ Skipping transco row due to missing information: " + JSON.stringify(item));
      return;
    }

    if (!transcoMap[item.out]) {
      transcoMap[item.out] = { sources: [], split: null };
    }
    transcoMap[item.out].sources.push(item.in);

    if (item.split) {
      transcoMap[item.out].split = item.split;
    }
  });

  Logger.log("✅ transcoMap built: " + JSON.stringify(transcoMap));
  return transcoMap;
}

/**
* Retrieves data from gsheet_in based on the specified row.
* @param {Sheet} sheet - Source sheet.
* @param {Object} transcoMap - Object containing column mappings.
* @param {number} row - Row to retrieve.
* @return {Object} - Retrieved data { "ColumnName": "Value" }
*/
function retrieveData(sheet, transcoMap, row) {
  var results = {};
  
  Object.keys(transcoMap).forEach(function(outCol) {
    if (!transcoMap[outCol] || !Array.isArray(transcoMap[outCol].sources)) {
      Logger.log("⚠️ Issue detected: transcoMap[" + outCol + "] is invalid or lacks sources.");
      return;
    }

    var values = [];

    transcoMap[outCol].sources.forEach(function(column) {
      var colIndex = findColumnIndex(sheet.getName(), 1, column);
      if (colIndex !== null) {
        var value = sheet.getRange(row, colIndex + 1).getValue();

        if (value instanceof Date) {
          value = Utilities.formatDate(value, Session.getScriptTimeZone(), "dd/MM/yyyy");
        }

        if (value !== "") {
          values.push(String(value));
        }
      }
    });

    if (transcoMap[outCol].split && values.length > 0) {
      var splitValues = [];
      values.forEach(function(v) {
        splitValues = splitValues.concat(v.split(transcoMap[outCol].split));
      });
      results[outCol] = splitValues;
    } else {
      results[outCol] = values;
    }
  });

  return results;
}

/**
* Writes the retrieved data to gsheet_out.
* @param {Sheet} sheet - Destination sheet.
* @param {Object} data - Data to write { "ColumnName": "Value" }
* @param {Object} transcoMap - Column mapping.
*/
function writeData(sheet, data, transcoMap, gsheet_out) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log("📌 Starting data writing to '" + gsheet_out + "'");

  Object.keys(data).forEach(function(outCol) {
    var values = data[outCol].filter(v => v !== ""); 
    var finalValue = values.join(transcoMap[outCol].split || ", ");

    var range = findNamedRangeOrCell(ss, sheet, outCol, gsheet_out);

    if (!range) {
      Logger.log("❌ ERROR: Unable to write to '" + outCol + "'");
      return;
    }

    try {
      range.setValue(finalValue);
      Logger.log("✅ Successfully wrote to '" + outCol + "': " + finalValue);
    } catch (e) {
      Logger.log("❌ ERROR writing to '" + outCol + "': " + e.message);
    }
  });

  Logger.log("✅ Writing process completed!");
}

/**
* Finds a named range or a standard cell.
*/
function findNamedRangeOrCell(ss, sheet, rangeName, gsheet_out) {
  var range = null;

  try {
    range = ss.getRangeByName(rangeName);
    if (range) return range;
  } catch (e) {}

  try {
    range = ss.getRangeByName(gsheet_out + "!" + rangeName);
    if (range) return range;
  } catch (e) {}

  if (/^[A-Z]+[0-9]+$/.test(rangeName)) {
    try {
      return sheet.getRange(rangeName);
    } catch (e) {}
  }

  return null;
}

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
