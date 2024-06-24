function createEmployeeSheets() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var dataSheet = ss.getSheetByName("Data Employee");
    var recapSheet = ss.getSheetByName("Recap");
    var slipGajiSheet = ss.getSheetByName("Slip Gaji");
    var dataRange = dataSheet.getDataRange();
    var dataValues = dataRange.getValues();
  
    var sourceFile = DriveApp.getFileById(ss.getId());
    var sourceFolder = sourceFile.getParents().next();
  
    for (var i = 1; i < dataValues.length; i++) {
      var row = dataValues[i];
      var nip = row[0];
      var nama = row[1];
      var dept = row[3];
      var email = row[4];
  
      // Check if the department folder already exists
      var folder;
      var existingFolders = sourceFolder.getFoldersByName(dept);
      if (existingFolders.hasNext()) {
        folder = existingFolders.next();
      } else {
        folder = sourceFolder.createFolder(dept);
      }
  
      // Check if the employee sheet already exists in the department folder
      var existingFiles = folder.getFilesByName(nama + " Slip Gaji");
      if (!existingFiles.hasNext()) {
        // Create a new Google Sheets file for the employee
        var employeeSheet = SpreadsheetApp.create(nama + " Slip Gaji");
        var copyFile = DriveApp.getFileById(employeeSheet.getId());
        folder.addFile(copyFile);
        DriveApp.getRootFolder().removeFile(copyFile);
  
        // Copy Slip Gaji sheet to the new Google Sheets file
        var slipGajiCopy = slipGajiSheet.copyTo(employeeSheet);
        slipGajiCopy.setName("Slip Gaji");
        employeeSheet.deleteSheet(employeeSheet.getSheetByName("Sheet1"));
  
        // Write link to the new Google Sheets file in the Link column
        var sheetUrl = employeeSheet.getUrl();
        dataSheet.getRange(i + 1, 6).setValue(sheetUrl);
  
        // Give edit access to the employee's email
        var targetSheet = SpreadsheetApp.openByUrl(sheetUrl);
        targetSheet.addEditor(email);
  
        // Copy corresponding NIP value to cell D10 in Slip Gaji sheet
        targetSheet.getSheetByName("Slip Gaji").getRange("D10").setValue(nip);
  
        // Protect specific cells
        protectRange(slipGajiCopy, "D10", null); // Protect D10 without additional editor
        protectRange(slipGajiCopy, "B12:J", null); // Protect B12:J with no additional editor
        protectRange(slipGajiCopy, "G11:J11", null); // Protect G11:J11 with no additional editor
        protectRange(slipGajiCopy, "B1:J10", null); // Protect B1:J10 with no additional editor
      }
    }
  }

  
function protectRange(sheet, rangeA1Notation, email) {
    var range = sheet.getRange(rangeA1Notation);
    var protection = range.protect().setDescription('Cell Protection');
    var me = Session.getEffectiveUser();

    // Ensure the current user is an editor before removing others. Otherwise, they may lose access altogether.
    protection.addEditor(me);
    protection.removeEditors(protection.getEditors());
    if (protection.canDomainEdit()) {
        protection.setDomainEdit(false);
    }

    if (email) {
        protection.addEditor(email);
    }
}
  