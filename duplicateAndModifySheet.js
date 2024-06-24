/**
 * Adds a "SLIP GAJI" sheet to spreadsheets within the same folder, renaming it to "SLIP FEE",
 * modifies formulas, hides the original sheet, and protects specific ranges in the new sheet.
 * Sets a time-based trigger to continue processing if more files are available.
 */
function addSlipGajiSheet() {
    // Define source and target sheet names
    var sourceSheetName = 'SLIP GAJI';
    var targetSheetName = 'SLIP FEE';
  
    // Get the active spreadsheet, its file ID, parent folder, and list of files in the folder
    var source = SpreadsheetApp.getActiveSpreadsheet();
    var sourceFile = DriveApp.getFileById(source.getId());
    var sourceFolder = sourceFile.getParents().next();
    var folderFiles = sourceFolder.getFiles();
    var scriptProperties = PropertiesService.getScriptProperties();
  
    // Retrieve processed files list or initialize an empty array
    var processedFiles = JSON.parse(scriptProperties.getProperty('processedFiles') || '[]');
  
    // Set maximum number of files to process in one execution
    var maxFilesToProcess = 6;  // Change this number based on your needs
    var filesProcessed = 0;
  
    // Get current user's email for protection
    var email = Session.getActiveUser().getEmail();
  
    // Loop through each file in the folder and process up to maxFilesToProcess files
    while (folderFiles.hasNext() && filesProcessed < maxFilesToProcess) {
      var file = folderFiles.next();
      Logger.log(file);
  
      // Check if the file has already been processed
      if (processedFiles.indexOf(file.getId()) === -1) {
        var spreadsheet = SpreadsheetApp.open(file);
  
        // Check if "SLIP GAJI" sheet already exists in the file
        var sheetSlipGaji = spreadsheet.getSheetByName("SLIP GAJI");
        if (sheetSlipGaji) {
          Logger.log('Sheet "SLIP GAJI" already exists. Skipping this file.');
          continue;
        }
  
        // Copy the source sheet to the target spreadsheet
        var sourceSheet = source.getSheetByName(sourceSheetName);
        if (sourceSheet) {
          sourceSheet.copyTo(spreadsheet).setName(sourceSheetName);
        } else {
          Logger.log('Source sheet ' + sourceSheetName + ' not found.');
          continue;
        }
  
        // Modify formulas in the target sheet
        var targetSheet = spreadsheet.getSheetByName(targetSheetName);
        if (targetSheet) {
          var range = targetSheet.getRange('C25:L54');
          var formulas = range.getFormulas();
          Logger.log("Replace formula");
          for (var i = 0; i < formulas.length; i++) {
            for (var j = 0; j < formulas[i].length; j++) {
              if (formulas[i][j]) {
                formulas[i][j] = formulas[i][j].replace('" "', '""');
              }
            }
          }
          range.setFormulas(formulas);
  
          // Set formula in cell D1 to refer to 'SLIP GAJI'!D1
          targetSheet.getRange('D1').setFormula('=\'SLIP GAJI\'!D1');
  
          // Hide the original "SLIP FEE" sheet
          targetSheet.hideSheet();
  
        } else {
          Logger.log('Target sheet ' + targetSheetName + ' not found in ' + file.getName());
        }
  
        // Protect specific ranges in the new "SLIP GAJI" sheet
        var sheetSlipGajiToProtect = spreadsheet.getSheetByName("SLIP GAJI");
        if (sheetSlipGajiToProtect) {
          Logger.log("Protect ranges");
          protectRange(sheetSlipGajiToProtect, 'A2:K72', email);
          protectRange(sheetSlipGajiToProtect, 'L:AF', email);
          protectRange(sheetSlipGajiToProtect, 'H1', email);
        }
  
        // Record this file as processed
        processedFiles.push(file.getId());
        filesProcessed++;
      }
    }
  
    // Update processed files list in script properties
    scriptProperties.setProperty('processedFiles', JSON.stringify(processedFiles));
  
    // If there are more files to process, set a time-based trigger to continue processing
    if (folderFiles.hasNext()) {
      ScriptApp.newTrigger('addSlipGajiSheet')
        .timeBased()
        .after(1 * 60 * 1000)  // 1 minute
        .create();
    } else {
      // Reset the processed files property when all files are processed
      scriptProperties.deleteProperty('processedFiles');
    }
  }


  /**
 * Protects a specified range in a given sheet by adding the user's email as the sole editor.
 * @param {Sheet} sheet - The sheet object where the range is located.
 * @param {string} rangeA1Notation - The A1 notation of the range to protect.
 * @param {string} email - The email address of the user to add as an editor.
 */
function protectRange(sheet, rangeA1Notation, email) {
    var range = sheet.getRange(rangeA1Notation);
    var protection = range.protect();
    protection.removeEditors(protection.getEditors());
    protection.addEditor(email);
    protection.setWarningOnly(false);
  }