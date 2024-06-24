function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Slip Fee')
        .addItem('Download Slip Fee', 'downloadSheetAsPDF')  // Adds menu item to download the sheet as PDF
        .addToUi();  // Adds the menu to the spreadsheet UI
  }
  