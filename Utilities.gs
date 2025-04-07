/**
 * Utilities.gs
 * General utility functions for SheetsLab
 * 
 * This file contains utility functions that are used across
 * the SheetsLab project.
 * 
 * @version 1.0.0
 */

/**
 * Activates a sheet by its GID (sheet ID)
 * @param {number} gid - The GID of the sheet to activate
 */
function activateSheetByGid(gid) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    
    for (const sheet of sheets) {
      if (sheet.getSheetId() === gid) {
        sheet.activate();
        return;
      }
    }
    
    // If we get here, the sheet wasn't found
    SpreadsheetApp.getUi().alert('Sheet with GID ' + gid + ' not found.');
  }
  
  /**
   * Generates a random ID
   * @param {number} length - The length of the ID (default: 8)
   * @return {string} The generated ID
   */
  function generateRandomId(length = 8) {
    const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
    let result = '';
    
    for (let i = 0; i < length; i++) {
      result += chars.charAt(Math.floor(Math.random() * chars.length));
    }
    
    return result;
  }
  
  /**
   * Formats a number as currency
   * @param {number} value - The value to format
   * @param {string} currencySymbol - The currency symbol to use (default: $)
   * @param {number} decimals - The number of decimal places (default: 2)
   * @return {string} The formatted currency string
   */
  function formatCurrency(value, currencySymbol = '$', decimals = 2) {
    return currencySymbol + value.toFixed(decimals).replace(/\d(?=(\d{3})+\.)/g, '$&,');
  }
  
  /**
   * Formats a date in a specified format
   * @param {Date} date - The date to format
   * @param {string} format - The format string (default: 'yyyy-MM-dd')
   * @return {string} The formatted date string
   */
  function formatDate(date, format = 'yyyy-MM-dd') {
    if (!(date instanceof Date)) {
      return 'Invalid date';
    }
    
    const year = date.getFullYear();
    const month = date.getMonth() + 1;
    const day = date.getDate();
    const hours = date.getHours();
    const minutes = date.getMinutes();
    const seconds = date.getSeconds();
    
    // Pad single-digit numbers with leading zeros
    const pad = (num) => (num < 10 ? '0' + num : num);
    
    // Replace format tokens with actual values
    return format
      .replace('yyyy', year)
      .replace('MM', pad(month))
      .replace('dd', pad(day))
      .replace('HH', pad(hours))
      .replace('mm', pad(minutes))
      .replace('ss', pad(seconds));
  }
  
  /**
   * Validates an email address
   * @param {string} email - The email address to validate
   * @return {boolean} True if valid, false otherwise
   */
  function isValidEmail(email) {
    const emailRegex = /^[a-zA-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?(?:\.[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?)*$/;
    return emailRegex.test(email);
  }
  
  /**
   * Escapes special characters in a string for use in HTML
   * @param {string} str - The string to escape
   * @return {string} The escaped string
   */
  function escapeHtml(str) {
    if (typeof str !== 'string') {
      return str;
    }
    
    return str
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }
  
  /**
   * Creates a backup of the current spreadsheet
   * @param {string} suffix - Optional suffix to add to the backup name
   * @return {string} The URL of the new backup spreadsheet
   */
  function createBackup(suffix = '') {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const timestamp = formatDate(new Date(), 'yyyy-MM-dd_HHmm');
    const backupName = ss.getName() + ' - Backup ' + timestamp + (suffix ? ' ' + suffix : '');
    
    // Create a copy of the spreadsheet
    const backup = DriveApp.getFileById(ss.getId()).makeCopy(backupName);
    
    // Return the URL of the new spreadsheet
    return backup.getUrl();
  }
  
  /**
   * Gets information about the current user
   * @return {Object} User information
   */
  function getUserInfo() {
    const email = Session.getEffectiveUser().getEmail();
    const username = email.split('@')[0];
    const domain = email.split('@')[1];
    
    return {
      email: email,
      username: username,
      domain: domain,
      isGoogleDomain: domain === 'gmail.com' || domain === 'googlemail.com',
    };
  }
  
  /**
   * Debounces a function call
   * @param {Function} func - The function to debounce
   * @param {number} wait - The wait time in milliseconds
   * @return {Function} The debounced function
   */
  function debounce(func, wait) {
    let timeout;
    
    return function executedFunction(...args) {
      const later = () => {
        clearTimeout(timeout);
        func(...args);
      };
      
      clearTimeout(timeout);
      timeout = setTimeout(later, wait);
    };
  }
  
  /**
   * Gets the column letter for a column index (1-based)
   * @param {number} columnIndex - The column index (1 for A, 2 for B, etc.)
   * @return {string} The column letter(s)
   */
  function getColumnLetter(columnIndex) {
    let columnLetter = '';
    
    while (columnIndex > 0) {
      const remainder = (columnIndex - 1) % 26;
      columnLetter = String.fromCharCode(65 + remainder) + columnLetter;
      columnIndex = Math.floor((columnIndex - 1) / 26);
    }
    
    return columnLetter;
  }
  
  /**
   * Gets the column index for a column letter (1-based)
   * @param {string} columnLetter - The column letter(s) (A, B, AA, etc.)
   * @return {number} The column index
   */
  function getColumnIndex(columnLetter) {
    let columnIndex = 0;
    
    for (let i = 0; i < columnLetter.length; i++) {
      columnIndex = columnIndex * 26 + (columnLetter.charCodeAt(i) - 64);
    }
    
    return columnIndex;
  }
  
  /**
   * Converts a range address to a range object
   * @param {string} rangeAddress - The range address (e.g., "A1:B10")
   * @param {SpreadsheetApp.Sheet} sheet - The sheet containing the range
   * @return {SpreadsheetApp.Range} The range object
   */
  function getRangeFromAddress(rangeAddress, sheet) {
    sheet = sheet || SpreadsheetApp.getActiveSheet();
    
    // Parse the range address
    const match = rangeAddress.match(/([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?/);
    
    if (!match) {
      throw new Error('Invalid range address: ' + rangeAddress);
    }
    
    const startCol = getColumnIndex(match[1]);
    const startRow = parseInt(match[2]);
    
    // If it's a single cell
    if (!match[3]) {
      return sheet.getRange(startRow, startCol);
    }
    
    // If it's a range
    const endCol = getColumnIndex(match[3]);
    const endRow = parseInt(match[4]);
    
    const numRows = endRow - startRow + 1;
    const numCols = endCol - startCol + 1;
    
    return sheet.getRange(startRow, startCol, numRows, numCols);
  }
  
  /**
   * Extracts unique values from a range of cells
   * Similar to UNIQUE formula but as a utility function
   * @param {Array<Array>} range - The range of values
   * @param {number} columnIndex - The column index to extract unique values from (0-based)
   * @return {Array} The unique values
   */
  function getUniqueValues(range, columnIndex = 0) {
    // Extract the specified column
    const values = range.map(row => row[columnIndex]);
    
    // Filter out duplicates and empty values
    const uniqueValues = values.filter((value, index, self) => 
      value !== '' && 
      value !== null && 
      value !== undefined && 
      self.indexOf(value) === index
    );
    
    return uniqueValues.sort();
  }
  
  /**
   * Gets the data range of a sheet excluding headers
   * @param {SpreadsheetApp.Sheet} sheet - The sheet to get data from
   * @param {number} headerRows - Number of header rows to exclude (default: 1)
   * @return {SpreadsheetApp.Range} The data range
   */
  function getDataRange(sheet, headerRows = 1) {
    sheet = sheet || SpreadsheetApp.getActiveSheet();
    
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();
    
    if (lastRow <= headerRows) {
      return null; // No data rows
    }
    
    return sheet.getRange(headerRows + 1, 1, lastRow - headerRows, lastColumn);
  }
  
  /**
   * Parses a URL and extracts query parameters
   * @param {string} url - The URL to parse
   * @return {Object} Object containing the parsed URL parts
   */
  function parseUrl(url) {
    try {
      const parsed = {};
      
      // Extract the protocol
      const protocolSplit = url.split('://');
      if (protocolSplit.length > 1) {
        parsed.protocol = protocolSplit[0];
        url = protocolSplit[1];
      }
      
      // Extract the domain and path
      const pathSplit = url.split('/');
      parsed.domain = pathSplit[0];
      
      // Extract query parameters
      const queryIndex = url.indexOf('?');
      if (queryIndex !== -1) {
        const queryString = url.substring(queryIndex + 1);
        const queryParams = {};
        
        queryString.split('&').forEach(param => {
          const [key, value] = param.split('=');
          queryParams[key] = decodeURIComponent(value || '');
        });
        
        parsed.queryParams = queryParams;
      }
      
      return parsed;
    } catch (error) {
      console.error('Error parsing URL:', error);
      return { error: 'Invalid URL' };
    }
  }