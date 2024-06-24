/**
 * Adds or updates a "Q3 - 1ON1" sheet in spreadsheets within the same folder:
 * 1. Renames existing "Q3 - 1ON1" to "Q3 - 1ON1_old" if it exists.
 * 2. Duplicates "Q2 - 1ON1" sheet and renames it to "Q3 - 1ON1".
 * 3. Activates a filter with specific criteria on column 10.
 * 4. Clears values in active cell but skips filtered rows.
 * 5. Removes filters.
 * 6. Finds and replaces "Q2 - 1ON1" within formulas with "Q3 - 1ON1".
 * 7. Moves the new "Q3 - 1ON1" sheet to a specified position.
 * 8. Protects specified ranges in the "Q3 - 1ON1" sheet.
 * Sets a time-based trigger to continue processing if more files are available.
 */
function addQ3Sheets() {
    // Get the active spreadsheet, its file ID, parent folder, and list of files in the folder
    var source = SpreadsheetApp.getActiveSpreadsheet();
    var sourceFile = DriveApp.getFileById(source.getId());
    var sourceFolder = sourceFile.getParents().next();
    var folderFiles = sourceFolder.getFiles();
    var scriptProperties = PropertiesService.getScriptProperties();
  
    // Retrieve processed files list or initialize an empty array
    var processedFiles = JSON.parse(scriptProperties.getProperty('processedFiles') || '[]');
  
    // Set maximum number of files to process in one execution
    var maxFilesToProcess = 2; // Change this number based on your needs
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
  
        // Step 1: Check if "Q3 - 1ON1_old" sheet exists, skip processing this file if it does
        var sheetQ3Old = spreadsheet.getSheetByName("Q3 - 1ON1_old");
        if (sheetQ3Old) {
          Logger.log('Sheet "Q3 - 1ON1_old" already exists. Skipping this file.');
          continue;
        }
  
        // Step 2: Rename existing "Q3 - 1ON1" sheet to "Q3 - 1ON1_old"
        var sheetQ3 = spreadsheet.getSheetByName("Q3 - 1ON1");
        if (sheetQ3) {
          Logger.log("Rename old sheet");
          sheetQ3.setName("Q3 - 1ON1_old");
          Logger.log("Hide old sheet");
          sheetQ3.hideSheet();
        }
  
        // Step 3: Duplicate "Q2 - 1ON1" sheet and rename duplicated sheet to "Q3 - 1ON1"
        var sheetQ2 = spreadsheet.getSheetByName("Q2 - 1ON1");
        if (sheetQ2) {
          Logger.log("Duplicate Q2 Sheet");
          var newSheet = sheetQ2.copyTo(spreadsheet);
          Logger.log("Rename Q2 to Q3");
          newSheet.setName("Q3 - 1ON1");
        } else {
          throw new Error('Sheet "Q2 - 1ON1" does not exist.');
        }
  
        // Step 4: Activate filter with specific criteria on column 10
        Logger.log("Applying column filter");
        var filter = newSheet.getFilter();
        if (filter) {
          // If filter exists, modify the existing filter criteria for column 10
          var criteria = SpreadsheetApp.newFilterCriteria()
            .setHiddenValues(['', 'FEE', 'NOMINAL'])
            .build();
          filter.setColumnFilterCriteria(10, criteria);
        } else {
          // If filter does not exist, create a new filter
          newSheet.getRange('J10:AN1260').createFilter();
          filter = newSheet.getFilter();
          var criteria = SpreadsheetApp.newFilterCriteria()
            .setHiddenValues(['', 'FEE', 'NOMINAL'])
            .build();
          filter.setColumnFilterCriteria(10, criteria);
        }
  
        // Step 5: Clear value in active cell but skip filtered rows
        Logger.log("Clear range");
        newSheet.getRange('K11:AN1258').activate();
        newSheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  
        // Step 6: Unfilter column
        Logger.log("Unfilter");
        var filter = newSheet.getFilter();
        if (filter) {
          filter.remove();
        }
  
        // Step 7: Find and replace "Q2 - 1ON1" within formulas then replace with "Q3 - 1ON1"
        Logger.log("Replace Formula in C11");
        var cellC11 = newSheet.getRange('C11');
        var cellC11Formula = cellC11.getFormula();
        if (cellC11Formula.includes("Q2 1ON1")) {
          cellC11.setFormula(cellC11Formula.replace(/Q2 1ON1/g, "Q3 1ON1"));
        }
  
        // Replace in column K
        Logger.log("Replace formula in column K");
        var rangeColumnK = newSheet.getRange('K:K');
        var formulasColumnK = rangeColumnK.getFormulas();
        for (var i = 0; i < formulasColumnK.length; i++) {
          if (formulasColumnK[i][0] && formulasColumnK[i][0].includes("Q2 1ON1")) {
            formulasColumnK[i][0] = formulasColumnK[i][0].replace(/Q2 1ON1/g, "Q3 1ON1");
          }
        }
        rangeColumnK.setFormulas(formulasColumnK);
  
        // Step 8: Move the new "Q3 - 1ON1" sheet to a specified position
        Logger.log("Move Sheet");
        spreadsheet.setActiveSheet(newSheet);
        spreadsheet.moveActiveSheet(20);
  
        // Sheet names to be protected
        var sheetsToProtect = ["Q3 - 1ON1"];
  
        // Protect specified ranges in the "Q3 - 1ON1" sheet
        sheetsToProtect.forEach(function(sheetName) {
          var sheet = spreadsheet.getSheetByName(sheetName);
          if (sheet) {
            // Protect range B10:I
            Logger.log("Protect B10:I");
            protectRange(sheet, 'B10:I', email);
  
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
  
        // Record this file as processed
        processedFiles.push(file.getId());
        filesProcessed++;
      }
    }
  
    // Update processed files list in script properties
    scriptProperties.setProperty('processedFiles', JSON.stringify(processedFiles));
  
    // If there are more files to process, set a time-based trigger to continue processing
    if (folderFiles.hasNext()) {
      ScriptApp.newTrigger('addQ3Sheets')
        .timeBased()
        .after(1 * 60 * 1000) // 1 minute
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
  