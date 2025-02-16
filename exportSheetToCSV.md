Function Overview
The exportSheetToCSV function exports a specified sheet from a Google Spreadsheet as a CSV file and saves it in a specified Google Drive folder.

Key Features
✔ Ensures the sheet exists before processing.
✔ Converts all cell values to text format to prevent data loss (e.g., leading zeros).
✔ Formats dates properly (DD/MM/YYYY format) for consistency.
✔ Saves the file directly in a specified Google Drive folder.

How It Works
Retrieves the sheet: The function first checks if the sheet exists. If not, it logs an error.
Extracts data: The entire data range of the sheet is retrieved.
Formats data:
Converts dates to "DD/MM/YYYY" format.
Ensures numbers are stored as text to prevent issues with leading zeros.
Escapes quotation marks (") to prevent CSV formatting errors.
Saves the CSV file: The processed data is stored as a CSV file in the specified Google Drive folder.
Parameters
Parameter	Type	Description
sheetName	string	Name of the sheet to export.
folderId	string	Google Drive folder ID where the file will be saved.
fileName	string	Name of the output CSV file (without .csv extension).
Return Value
No return value.
Creates a CSV file in Google Drive and displays a success message.
Logs important details for debugging.
Example Usage
javascript
Copy
exportSheetToCSV("SalesData", "1a2b3c4d5e6f7g8h9i0j", "Sales_Report_Feb");
📌 This will:

Export the "SalesData" sheet.
Save it in the Google Drive folder with ID "1a2b3c4d5e6f7g8h9i0j".
Name the output file "Sales_Report_Feb.csv".
Error Handling
Error Scenario	Handling Mechanism
The sheet does not exist	Logs an error (❌ The sheet does not exist) and stops execution.
Folder ID is incorrect	Google Drive will throw an error when trying to save the file.
Cell data contains special characters	Quotation marks (") are escaped automatically.
Use Cases
✔ Data Backup: Save spreadsheet data as CSV for external storage.
✔ Integration with Other Systems: Many external systems accept CSV as input.
✔ Automated Data Extraction: Schedule exports to update CSV files in Google Drive.