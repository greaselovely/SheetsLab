// Integration.gs
// SheetsLab - External API connections and processing functions

/**
 * Fetches data from an external API using the provided endpoint.
 * @param {string} endpoint - The API endpoint to call.
 * @returns {Object} - The parsed JSON response from the API.
 */
function fetchDataFromApi(endpoint) {
  try {
    var response = UrlFetchApp.fetch(endpoint);
    var json = JSON.parse(response.getContentText());
    return json;
  } catch (e) {
    Logger.log("Error fetching API data: " + e);
    throw new Error("Failed to fetch API data.");
  }
}

/**
 * Demonstrates API integration by fetching data from the ISS API.
 * The ISS API endpoint is: http://api.open-notify.org/iss-now.json
 */
function demoApiIntegration() {
  var endpoint = "http://api.open-notify.org/iss-now.json";
  var data = fetchDataFromApi(endpoint);
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = ss.getSheetByName(CONFIG.SHEETS.DATA);
  if (!dataSheet) {
    throw new Error("DATA sheet not found.");
  }
  
  // Clear existing data and set up headers
  dataSheet.clear();
  var headers = ["Timestamp", "Latitude", "Longitude"];
  dataSheet.appendRow(headers);
  
  var timestamp = data.timestamp;
  var latitude = data.iss_position.latitude;
  var longitude = data.iss_position.longitude;
  
  dataSheet.appendRow([timestamp, latitude, longitude]);
  
  SpreadsheetApp.getUi().alert("ISS API integration demo completed. Check the DATA sheet.");
}

/**
 * Processes the advanced form submission.
 * @param {string} name - The name entered by the user.
 * @param {string} email - The email address entered by the user.
 * @param {string} message - The message entered by the user.
 * @return {string} Success message.
 */
function processAdvancedForm(name, email, message) {
  Logger.log("Advanced Form Submission - Name: " + name + ", Email: " + email + ", Message: " + message);
  return "Form submitted successfully!";
}

/**
 * Processes email automation by sending an email using Gmail.
 * @param {string} recipient - The recipient's email address.
 * @param {string} subject - The subject of the email.
 * @param {string} body - The body of the email.
 * @return {string} Success message.
 */
function processEmailAutomation(recipient, subject, body) {
  try {
    GmailApp.sendEmail(recipient, subject, body);
    return "Email sent successfully to " + recipient + "!";
  } catch (e) {
    Logger.log("Error sending email: " + e);
    throw new Error("Failed to send email.");
  }
}

/**
 * Processes CSV data by parsing it and appending it to the DATA sheet.
 * @param {string} csvContent - The CSV content as a string.
 * @return {string} Success message.
 */
function processCsvData(csvContent) {
  var csvData = Utilities.parseCsv(csvContent);
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = ss.getSheetByName(CONFIG.SHEETS.DATA);
  if (!dataSheet) {
    throw new Error("DATA sheet not found.");
  }
  
  // Clear the sheet and append the CSV data
  dataSheet.clear();
  csvData.forEach(function(row) {
    dataSheet.appendRow(row);
  });
  
  return "CSV data imported successfully!";
}

/**
 * Demonstrates CSV import by using a sample CSV string.
 */
function demoCsvImport() {
  var sampleCsv = "Name,Age,Email\nAlice,30,alice@example.com\nBob,25,bob@example.com\nCharlie,35,charlie@example.com";
  var message = processCsvData(sampleCsv);
  SpreadsheetApp.getUi().alert(message);
}
