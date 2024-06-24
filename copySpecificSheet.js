/**
 * Copies the third sheet from the active spreadsheet to all other spreadsheets
 * in the same folder. If a sheet named "Dashboard" already exists, it renames 
 * it to "Dashboard_old" before copying the new sheet. Skips files that already 
 * contain a "Dashboard_old" sheet.
 */
function copyDashboardSheet() {
    // Get the active spreadsheet and the third sheet in it
    var source = SpreadsheetApp.getActiveSpreadsheet();
    var sourceSheets = source.getSheets()[2];
    var sourceFile = DriveApp.getFileById(source.getId());
    var sourceFolder = sourceFile.getParents().next();
    var folderFiles = sourceFolder.getFiles();
  
    // Define target sheet names
    var targetSheetName = 'Dashboard';
    var targetOldSheetName = targetSheetName + '_old';
  
    // Iterate over all files in the folder
    while (folderFiles.hasNext()) {
      var thisFile = folderFiles.next();
      Logger.log(thisFile); // Log the file for debugging purposes
  
      // Skip the active file itself
      if (thisFile.getName() !== sourceFile.getName()) {
        var currentSS = SpreadsheetApp.openById(thisFile.getId());
        var existingSheetNames = currentSS.getSheets().map(function (s) { return s.getName(); });
        
        // Check if target sheet or its "_old" version already exists in the current spreadsheet
        var targetSheetIndex = existingSheetNames.indexOf(targetSheetName);
        var targetOldSheetIndex = existingSheetNames.indexOf(targetOldSheetName);
  
        // If "_old" version exists, skip copying to this file
        if (targetOldSheetIndex !== -1) {
          continue;
        }
  
        // If original target sheet exists, rename it to "_old" version
        if (targetSheetIndex !== -1) {
          var oldSheet = currentSS.getSheets()[targetSheetIndex];
          Logger.log("Rename Old Sheet"); // Log the rename action
          oldSheet.setName(targetOldSheetName);
          Logger.log("Hiding Old Sheet"); // Log the hide action
          oldSheet.hideSheet();
        }
  
        // Copy the sheet to the current spreadsheet
        var newSheet = sourceSheets.copyTo(currentSS);
        Logger.log("Copy " + targetSheetName); // Log the copy action
        newSheet.setName(targetSheetName);
  
        // Move the new sheet to index 2 (third position)
        currentSS.setActiveSheet(newSheet);
        Logger.log("Moving New Sheet"); // Log the move action
        currentSS.moveActiveSheet(2);
      }
    }
  }
  