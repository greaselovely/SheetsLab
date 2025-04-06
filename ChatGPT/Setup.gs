// Setup.gs
// SheetsLab - Sheet creation and initialization

function setupSheets() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetNames = [CONFIG.SHEETS.MAIN, CONFIG.SHEETS.DATA, CONFIG.SHEETS.DASHBOARD];
    var existingSheets = ss.getSheets().map(function(sheet) {
      return sheet.getName();
    });
    
    sheetNames.forEach(function(name) {
      if (existingSheets.indexOf(name) === -1) {
        ss.insertSheet(name);
        Logger.log("Created sheet: " + name);
      } else {
        Logger.log("Sheet already exists: " + name);
      }
    });
    
    // Optionally remove default "Sheet1" if it's extra
    var defaultSheet = ss.getSheetByName("Sheet1");
    if (defaultSheet && existingSheets.indexOf("Sheet1") !== -1 && ss.getSheets().length > sheetNames.length) {
      ss.deleteSheet(defaultSheet);
      Logger.log("Deleted default sheet: Sheet1");
    }
  }
  