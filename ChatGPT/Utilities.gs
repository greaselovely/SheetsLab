// Utilities.gs
// SheetsLab - General utility functions

/**
 * Logs informational messages if debugging is enabled.
 * @param {string} message - The message to log.
 */
function logInfo(message) {
    if (CONFIG.DEBUG) {
      Logger.log("[INFO] " + message);
    }
  }
  
  /**
   * Logs error messages.
   * @param {string} message - The error message to log.
   */
  function logError(message) {
    Logger.log("[ERROR] " + message);
  }
  
  /**
   * Safely parses a JSON string and returns an object.
   * @param {string} jsonString - The JSON string to parse.
   * @return {Object|null} The parsed JSON object, or null if parsing fails.
   */
  function safeParseJSON(jsonString) {
    try {
      return JSON.parse(jsonString);
    } catch (e) {
      logError("Failed to parse JSON: " + e);
      return null;
    }
  }
  
  /**
   * Formats a number to a fixed number of decimal places.
   * @param {number} num - The number to format.
   * @param {number} decimals - The number of decimals.
   * @return {string} The formatted number.
   */
  function formatNumber(num, decimals) {
    return Number(num).toFixed(decimals);
  }
  
  /**
   * Checks if the provided value is empty (null, undefined, or an empty string).
   * @param {*} value - The value to check.
   * @return {boolean} True if the value is empty, false otherwise.
   */
  function isEmpty(value) {
    return (value === null || value === undefined || value === '');
  }
  
  /**
   * Debounces a function, limiting how often it can run.
   * Useful for managing rapid UI events.
   * @param {Function} func - The function to debounce.
   * @param {number} wait - The debounce delay in milliseconds.
   * @return {Function} The debounced function.
   */
  function debounce(func, wait) {
    var timeout;
    return function() {
      var context = this, args = arguments;
      clearTimeout(timeout);
      timeout = setTimeout(function() {
        func.apply(context, args);
      }, wait);
    };
  }
  