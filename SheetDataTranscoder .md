Overview
This script is designed to transcode and copy data from one Google Sheets tab to another using a transcoding table. It retrieves specific data from a source sheet, applies predefined transcoding rules, and writes the transformed data into a destination sheet.

How It Works
The script follows these four key steps:

Retrieve the transcoding table – Loads the mapping of source columns to destination columns.
Apply transcoding rules – Structures the mappings, handling concatenation or splitting if necessary.
Retrieve data from the source sheet – Extracts data based on the mappings.
Write data into the destination sheet – Inserts transformed data in the specified format.
Main Function
transcodeAndCopy(gsheet_in, gsheet_out, transco, row)
Parameters
gsheet_in (string) – Name of the source sheet.
gsheet_out (string) – Name of the destination sheet.
transco (string) – Name of the transcoding table sheet.
row (number) – Row number from the source sheet to copy.
Description
Checks if all sheets exist; otherwise, it displays an error.
Loads the transcoding table from the transco sheet.
Processes the transcoding rules to build a mapping object.
Retrieves and formats data from the specified row in gsheet_in.
Writes the transformed data into gsheet_out.
Example Usage
javascript
Copy
transcodeAndCopy("InputSheet", "OutputSheet", "TranscoSheet", 5);
Copies and transcodes data from row 5 in "InputSheet", using mappings from "TranscoSheet", and writes it to "OutputSheet".

Supporting Functions
1. Retrieve Transcoding Table
retrieveTranscoTable(sheet)
Purpose: Loads the transcoding table and converts it into an array of objects.

Returns
Array – A list of objects structured as:
json
Copy
[
  { "in": "SourceColumn", "out": "TargetColumn", "split": "Separator" }
]
2. Apply Transcoding Rules
applyTranscoRules(transcoTable)
Purpose: Converts the array of transcoding rules into a structured mapping object.

Returns
Object – A dictionary with destination columns as keys:
json
Copy
{
  "TargetColumn": { "sources": ["SourceColumn1", "SourceColumn2"], "split": "," }
}
3. Retrieve Data from the Source Sheet
retrieveData(sheet, transcoMap, row)
Purpose: Retrieves the data from the specified row based on the transcoding map.

Returns
Object – A dictionary with destination columns as keys:
json
Copy
{
  "TargetColumn": ["Value1", "Value2"]
}
4. Write Data to the Destination Sheet
writeData(sheet, data, transcoMap, gsheet_out)
Purpose: Writes the transformed data into the correct locations in the destination sheet.

5. Helper Functions
Find Named Range or Cell
findNamedRangeOrCell(ss, sheet, rangeName, gsheet_out)

Finds a named range or a cell reference in the destination sheet.
Find Column Index
findColumnIndex(sheetName, headerRow, value)

Finds the column index of a given header name.
Example Workflow
Scenario: Mapping and Copying Data
Transcoding Table (TranscoSheet)
Source Column	Target Column	Split Separator
Name	Full Name	
DOB	Date of Birth	
Address	Location	
Tags	Categories	,
Source Sheet (InputSheet)
Name	DOB	Address	Tags
John	01/01/2000	123 Street	Tag1, Tag2
Expected Output (OutputSheet)
Full Name	Date of Birth	Location	Categories
John	01/01/2000	123 Street	Tag1, Tag2
Execution
javascript
Copy
transcodeAndCopy("InputSheet", "OutputSheet", "TranscoSheet", 2);