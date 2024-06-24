/**
 * Protects specific ranges and sheets in all spreadsheets within the same folder
 * as the active spreadsheet. It processes files incrementally and sets a time-based
 * trigger if there are more files to process.
 */
function protectRanges() {
    // Get the active spreadsheet and its parent folder
    var source = SpreadsheetApp.getActiveSpreadsheet();
    var sourceFile = DriveApp.getFileById(source.getId());
    var sourceFolder = sourceFile.getParents().next();
    var folderFiles = sourceFolder.getFiles();
  
    // Get script properties to track processed files
    var scriptProperties = PropertiesService.getScriptProperties();
    var processedFiles = JSON.parse(scriptProperties.getProperty('processedFiles') || '[]');
    
    // Define the maximum number of files to process in one run
    var maxFilesToProcess = 1;
    var filesProcessed = 0;
  
    // Get the email of the active user to set as the editor
    var email = Session.getActiveUser().getEmail();
  
    // Process files until the maximum is reached or no more files are found
    while (folderFiles.hasNext() && filesProcessed < maxFilesToProcess) {
      var file = folderFiles.next();
      Logger.log(file); // Log the file for debugging purposes
  
      // Check if the file has already been processed
      if (processedFiles.indexOf(file.getId()) === -1) {    
        var spreadsheet = SpreadsheetApp.open(file);
  
        // Protect the "DASHBOARD" sheet
        Logger.log("Protect Dashboard Sheet");
        protectSheet(spreadsheet, "DASHBOARD", email);
  
        // Sheet names to be protected
        var sheetsToProtect = ["Q2 - 1ON1", "P GRUP"];
        
        sheetsToProtect.forEach(function(sheetName) {
          var sheet = spreadsheet.getSheetByName(sheetName);
          Logger.log(sheetName); // Log the sheet name
  
          if (sheet) {        
            // Protect range B10:I
            Logger.log("Protect B10:I");
            protectRange(sheet, 'B10:I', email);
            
            // Protect specific ranges in the sheet
            // Protect ranges J14:V15, J19:V120, J24:V25, J29:V30, etc.
            Logger.log("Protect Fee dan Nominal");
            for (var row = 14; row <= 60; row += 5) {
              var range1 = `J${row}:V${row+1}`;
              var range2 = `J${row+5}:V${row+6}`;
              protectRange(sheet, range1, email);
              protectRange(sheet, range2, email);
            }
          }
        });
  
        // Mark the file as processed
        processedFiles.push(file.getId());
        filesProcessed++;
      }
    }
  
    // Save the updated list of processed files
    scriptProperties.setProperty('processedFiles', JSON.stringify(processedFiles));
  
    // If there are more files to process, set a time-based trigger to continue processing
    if (folderFiles.hasNext()) {
      ScriptApp.newTrigger('protectRanges')
        .timeBased()
        .after(1 * 60 * 1000) // 1 minute
        .create();
    } else {
      // Reset the processed files property when all files are processed
      scriptProperties.deleteProperty('processedFiles');
    }
  }
  
  /**
   * Protects a specified range in a sheet, allowing only the given email to edit it.
   *
   * @param {Sheet} sheet - The sheet containing the range to be protected.
   * @param {string} rangeA1Notation - The A1 notation of the range to be protected.
   * @param {string} email - The email address of the user to be allowed to edit the range.
   */
  function protectRange(sheet, rangeA1Notation, email) {
    var range = sheet.getRange(rangeA1Notation);
    var protection = range.protect();
    protection.removeEditors(protection.getEditors());
    protection.addEditor(email);
    protection.setWarningOnly(false);
  }
  
  /**
   * Protects an entire sheet, allowing only the given email to edit it.
   *
   * @param {Spreadsheet} spreadsheet - The spreadsheet containing the sheet to be protected.
   * @param {string} sheetName - The name of the sheet to be protected.
   * @param {string} email - The email address of the user to be allowed to edit the sheet.
   */
  function protectSheet(spreadsheet, sheetName, email) {
    var sheet = spreadsheet.getSheetByName(sheetName);
    if (sheet) {
      var protection = sheet.protect();
      protection.removeEditors(protection.getEditors());
      protection.addEditor(email);
      protection.setWarningOnly(false);
    }
  }
  