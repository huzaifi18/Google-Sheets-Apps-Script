/**
 * Copies specific sheets from the active spreadsheet to all spreadsheets within the same folder,
 * renaming existing sheets as "_old" versions if necessary.
 * Sheets to be copied include "Q2 - 1ON1" and "P GRUP".
 */
function copySheet() {
    // Get the active spreadsheet and its parent folder
    var source = SpreadsheetApp.getActiveSpreadsheet();
    var sourceSheets = source.getSheets(); // Get all sheets in the active spreadsheet
    var sourceFile = DriveApp.getFileById(source.getId());
    var sourceFolder = sourceFile.getParents().next();
    var folderFiles = sourceFolder.getFiles();
  
    // Define sheet names to be copied
    var targetSheetNames = ['Q2 - 1ON1', 'P GRUP'];
  
    // Iterate through each file in the folder
    while (folderFiles.hasNext()) {
      var thisFile = folderFiles.next();
      Logger.log(thisFile);
  
      // Skip the current active spreadsheet file
      if (thisFile.getName() !== sourceFile.getName()) {
        var currentSS = SpreadsheetApp.openById(thisFile.getId());
        var existingSheetNames = currentSS.getSheets().map(function(s) { return s.getName(); });
  
        // Iterate over each target sheet name to copy
        for (var i = 0; i < targetSheetNames.length; i++) {
          var targetSheetName = targetSheetNames[i];
          var targetOldSheetName = targetSheetName + '_old';
  
          // Check if target sheet or its "_old" version already exists in the current spreadsheet
          var targetSheetIndex = existingSheetNames.indexOf(targetSheetName);
          var targetOldSheetIndex = existingSheetNames.indexOf(targetOldSheetName);
  
          // If "_old" version exists, skip copying to this file
          if (targetOldSheetIndex !== -1) {
            continue;
          }
  
          // If original target sheet exists, rename it to "_old" version
          if (targetSheetIndex !== -1) {
            currentSS.getSheets()[targetSheetIndex].setName(targetOldSheetName);
          }
  
          // Copy the sheet from the source spreadsheet to the current spreadsheet
          var sheetToCopy = sourceSheets[i]; // Assuming Q2 - 1ON1 is at index 0 and P GRUP is at index 1
          var newSheet = sheetToCopy.copyTo(currentSS);
          Logger.log("Copying sheet: " + targetSheetName);
          newSheet.setName(targetSheetName);
        }
      }
    }
  }
  