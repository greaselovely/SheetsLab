// DataTools.gs
// SheetsLab - Data handling utilities

/**
 * Validates a range by ensuring no empty cells exist.
 * @param {string} sheetName - The name of the sheet.
 * @param {string} rangeA1Notation - The A1 notation of the range to validate.
 * @returns {boolean} - Returns true if all cells are non-empty, otherwise false.
 */
function validateDataRange(sheetName, rangeA1Notation) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) {
      throw new Error("Sheet '" + sheetName + "' not found.");
    }
    var range = sheet.getRange(rangeA1Notation);
    var values = range.getValues();
    for (var i = 0; i < values.length; i++) {
      for (var j = 0; j < values[i].length; j++) {
        if (values[i][j] === "") {
          return false;
        }
      }
    }
    return true;
  }
  
  /**
   * Filters data in a sheet based on a specific column value.
   * Assumes the first row is the header.
   * @param {string} sheetName - The name of the sheet.
   * @param {number} columnNumber - The column index (1-based) to filter by.
   * @param {*} filterValue - The value to filter on.
   * @returns {Array} - Returns the filtered data, including headers.
   */
  function filterDataByColumn(sheetName, columnNumber, filterValue) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) {
      throw new Error("Sheet '" + sheetName + "' not found.");
    }
    var data = sheet.getDataRange().getValues();
    // Keep header and rows matching the filter
    var filtered = data.filter(function(row, index) {
      if (index === 0) return true; // always keep header
      return row[columnNumber - 1] == filterValue;
    });
    return filtered;
  }
  
  /**
   * Transforms all string data in a given range to uppercase.
   * @param {string} sheetName - The name of the sheet.
   * @param {string} rangeA1Notation - The A1 notation of the range to transform.
   */
  function transformDataToUpperCase(sheetName, rangeA1Notation) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) {
      throw new Error("Sheet '" + sheetName + "' not found.");
    }
    var range = sheet.getRange(rangeA1Notation);
    var values = range.getValues();
    var transformed = values.map(function(row) {
      return row.map(function(cell) {
        return (typeof cell === "string") ? cell.toUpperCase() : cell;
      });
    });
    range.setValues(transformed);
  }
  
  /**
   * Appends a row of data to the specified sheet.
   * @param {string} sheetName - The name of the sheet.
   * @param {Array} rowData - An array representing a single row of data.
   */
  function appendRowData(sheetName, rowData) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) {
      throw new Error("Sheet '" + sheetName + "' not found.");
    }
    sheet.appendRow(rowData);
  }
  
  /**
   * Clears content in a specified range.
   * @param {string} sheetName - The name of the sheet.
   * @param {string} rangeA1Notation - The A1 notation of the range to clear.
   */
  function clearDataRange(sheetName, rangeA1Notation) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) {
      throw new Error("Sheet '" + sheetName + "' not found.");
    }
    var range = sheet.getRange(rangeA1Notation);
    range.clearContent();
  }
  