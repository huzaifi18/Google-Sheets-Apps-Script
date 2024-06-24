function downloadSheetAsPDF() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName('SLIP GAJI');  // Gets the sheet named "SLIP GAJI" from the active spreadsheet
    
    if (!sheet) {
      SpreadsheetApp.getUi().alert('Sheet "SLIP GAJI" not found!');  // Displays an alert if the sheet is not found
      return;
    }
    
    // Define the range to export to PDF
    var range = sheet.getRange('A2:K72');
    var sheetId = sheet.getSheetId();  // Gets the ID of the sheet
    var rangeParameters = 'A2:K72';  // Defines the range for export
  
    // PDF export settings
    var topMargin = 0.32;  // Top margin in inches (0.811 cm)
    var rightMargin = 0.263;  // Right margin in inches (1.408 cm)
    var bottomMargin = 0;  // Bottom margin in inches (0 cm)
    var leftMargin = 0.263;  // Left margin in inches (0.668 cm)
  
    // Construct the export URL for PDF generation
    var url = 'https://docs.google.com/spreadsheets/d/' + spreadsheet.getId() +
              '/export?format=pdf' +
              '&size=A4' +            // Paper size
              '&portrait=true' +      // Orientation, true for portrait, false for landscape
              '&fitw=true' +          // Fit to width, false for actual size
              '&sheetnames=false&printtitle=false&pagenumbers=false' + // Optional settings
              '&gridlines=false' +    // Whether to show gridlines
              '&fzr=false' +          // Repeat row headers (frozen rows) on each page
              '&gid=' + sheetId +     // Sheet ID
              '&range=' + rangeParameters + // Range
              '&top_margin=' + topMargin +
              '&right_margin=' + rightMargin +
              '&bottom_margin=' + bottomMargin +
              '&left_margin=' + leftMargin;
  
    // Get OAuth token for authentication
    var token = ScriptApp.getOAuthToken();
    
    // Fetch the PDF file using URL Fetch service with OAuth token
    var response = UrlFetchApp.fetch(url, {
      headers: {
        'Authorization': 'Bearer ' + token
      }
    });
  
    // Convert response to Blob
    var blob = response.getBlob().setName(sheet.getName() + '.pdf');
  
    // Provide a link to download the PDF
    var ui = SpreadsheetApp.getUi();
    var htmlOutput = HtmlService.createHtmlOutput('<a href="' + DriveApp.createFile(blob).getUrl() + '" target="_blank">Download PDF</a>');
    ui.showModalDialog(htmlOutput, 'Download PDF');
  }
  