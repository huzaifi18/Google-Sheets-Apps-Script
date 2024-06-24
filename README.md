# Google Sheets Apps Script Repository

This repository contains a collection of Google Sheets Apps Script functions designed to automate various tasks related to spreadsheet management and data processing. Below is an overview of the scripts available:

## Scripts Overview

### 1. `copyMultipleSheets.js`

This script copies specific sheets ("Q2 - 1ON1" and "P GRUP") from the active spreadsheet to all spreadsheets within the same folder. It renames existing sheets as "_old" versions if necessary.

### 2. `copySpecificSheet.js`

Copies the third sheet from the active spreadsheet to all other spreadsheets in the same folder. If a sheet named "Dashboard" already exists, it renames it to "Dashboard_old" before copying the new sheet.

### 3. `createCustomMenu.js`

Adds a custom menu to the Google Sheets UI ("Slip Fee"), allowing users to download the sheet as a PDF.

### 4. `duplicateAndModifySheet_2.js`

Adds or updates a "Q3 - 1ON1" sheet in spreadsheets within the same folder, performs various modifications, and protects specific ranges in the sheet.

### 5. `duplicateAndModifySheet.js`

Adds a "SLIP GAJI" sheet to spreadsheets within the same folder, renames it to "SLIP FEE," modifies formulas, hides the original sheet, and protects specific ranges in the new sheet.

### 6. `generateSpreadsheetFile.js`

Creates employee-specific Google Sheets files within department folders, copies necessary sheets, sets formulas, and protects specific ranges.

### 7. `PDFCreationAndDownload.js`

Allows users to download the "SLIP GAJI" sheet as a PDF with customized settings such as margins and paper size.

### 8. `protectRangeOrSheet.js`

Protects specific ranges and sheets in all spreadsheets within the same folder as the active spreadsheet, with incremental file processing and time-based triggers.

## Usage

Each script serves a specific purpose related to managing Google Sheets within a folder. To use these scripts:

1. Open the desired Google Sheet where you want to execute the script.
2. Navigate to "Extensions" > "Apps Script" and paste the script code.
3. Save and run the script from the Apps Script editor.
4. Adjust script parameters and configurations as needed for your use case.

## Contributing

Feel free to contribute to this repository by submitting pull requests or issues if you encounter any bugs or have suggestions for improvements.

## License

This repository is licensed under the MIT License. See the [LICENSE](./LICENSE) file for more details.
