/**
 * Recalculates formulas in the "Sheet1" sheet for a specified number of spreadsheets
 * in the same folder as the active spreadsheet. It processes files incrementally
 * and sets a time-based trigger if there are more files to process.
 */
function recalculateVRegSheet() {
    // Get the active spreadsheet and its parent folder
    var source = SpreadsheetApp.getActiveSpreadsheet();
    var sourceFile = DriveApp.getFileById(source.getId());
    var sourceFolder = sourceFile.getParents().next();
    var folderFiles = sourceFolder.getFiles();
  
    // Get script properties to track processed files
    var scriptProperties = PropertiesService.getScriptProperties();
    var processedFiles = JSON.parse(scriptProperties.getProperty('processedFiles') || '[]');
    
    // Define the maximum number of files to process in one run
    var maxFilesToProcess = 5; 
    var filesProcessed = 0;
  
    // Process files until the maximum is reached or no more files are found
    while (folderFiles.hasNext() && filesProcessed < maxFilesToProcess) {
      var file = folderFiles.next();
      Logger.log(file); // Log the file for debugging purposes
  
      // Check if the file has already been processed
      if (processedFiles.indexOf(file.getId()) === -1) {
        var spreadsheet = SpreadsheetApp.open(file);
        var sheet = spreadsheet.getSheetByName("Sheet1");
  
        // If the "Sheet1" sheet is found, recalculate its formulas
        if (sheet) {
          var range = sheet.getDataRange();
          var formulas = range.getFormulas();
  
          // Iterate through all cells and set their formulas again to recalculate
          for (var rowIndex = 0; rowIndex < formulas.length; rowIndex++) {
            for (var colIndex = 0; colIndex < formulas[rowIndex].length; colIndex++) {
              if (formulas[rowIndex][colIndex]) {
                sheet.getRange(rowIndex + 1, colIndex + 1).setFormula(formulas[rowIndex][colIndex]);
              }
            }
          }
  
          SpreadsheetApp.flush(); // Apply all pending changes
          processedFiles.push(file.getId()); // Mark the file as processed
          filesProcessed++;
        } else {
          Logger.log("Sheet named 'Sheet1' not found in file: " + file.getName()); // Log if the sheet is not found
        }
      }
    }
  
    // Save the updated list of processed files
    scriptProperties.setProperty('processedFiles', JSON.stringify(processedFiles));
  
    // If there are more files to process, set a time-based trigger to continue processing
    if (folderFiles.hasNext()) {
      ScriptApp.newTrigger('recalculateVRegSheet')
        .timeBased()
        .after(1 * 60 * 1000) // 1 minute
        .create();
    } else {
      // Reset the processed files property when all files are processed
      scriptProperties.deleteProperty('processedFiles');
    }
  }
  