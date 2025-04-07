/**
 * Setup.gs
 * Sheet creation and initialization for SheetsLab
 * 
 * This file contains functions to create and set up all necessary sheets
 * and initial configurations for the SheetsLab project.
 * 
 * @version 1.0.0
 */

/**
 * Initialize the entire SheetsLab project
 * Creates all necessary sheets and sets up initial configurations
 * @param {boolean} showConfirmation - Whether to show confirmation dialogs (default: true)
 */
function initializeSheetsLab(showConfirmation = true) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  // Confirm setup with user if showConfirmation is true
  if (showConfirmation) {
    const response = ui.alert(
      'Initialize SheetsLab',
      'This will create several new sheets in your spreadsheet to demonstrate Google Sheets capabilities. Continue?',
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) {
      return;
    }
  }
  
  // Create each lab sheet
  createHomeSheet(ss);
  createUILabSheet(ss);
  createDataLabSheet(ss);
  createVisualizationLabSheet(ss);
  createIntegrationLabSheet(ss);
  createFormulaLabSheet(ss);
  
  // Set Home as active sheet
  ss.getSheetByName(CONFIG.SHEETS.HOME).activate();
  
  // Show success message if showConfirmation is true
  if (showConfirmation) {
    ui.alert(
      'SheetsLab Initialized',
      'All lab sheets have been created successfully. You can now explore the different capabilities via the SheetsLab menu.',
      ui.ButtonSet.OK
    );
  }
}

/**
 * Creates the Home sheet with navigation and overview
 * @param {SpreadsheetApp.Spreadsheet} ss - The active spreadsheet
 */
function createHomeSheet(ss) {
  // Check if sheet exists, if not create it
  let sheet = ss.getSheetByName(CONFIG.SHEETS.HOME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.HOME);
  } else {
    // Clear existing content if sheet exists
    sheet.clear();
  }
  
  // Set up basic formatting
  sheet.setTabColor(CONFIG.COLORS.PRIMARY);
  
  // Create title and description
  sheet.getRange("A1:F1").merge().setValue(CONFIG.PROJECT_NAME)
    .setFontSize(24)
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
  
  sheet.getRange("A2:F2").merge().setValue("A Comprehensive Google Sheets Capability Showcase")
    .setFontSize(14)
    .setFontStyle("italic")
    .setHorizontalAlignment("center");
  
  // Create navigation section
  sheet.getRange("A4").setValue("Navigation:").setFontWeight("bold");
  
  // Create links to other sheets
  const labSheets = [
    {name: CONFIG.SHEETS.UI_LAB, description: "Interactive UI elements like sidebars, modals, and custom menus"},
    {name: CONFIG.SHEETS.DATA_LAB, description: "Data handling, validation, filtering, and automation"},
    {name: CONFIG.SHEETS.VISUALIZATION_LAB, description: "Charts, dashboards, and data visualization"},
    {name: CONFIG.SHEETS.INTEGRATION_LAB, description: "External API connections and service integrations"},
    {name: CONFIG.SHEETS.FORMULA_LAB, description: "Advanced formula and function demonstrations"}
  ];
  
  // Add sheet links with descriptions
  for (let i = 0; i < labSheets.length; i++) {
    const rowIndex = i + 5;
    sheet.getRange(`A${rowIndex}`).setValue(labSheets[i].name);
    
    // Create hyperlink formula to the sheet
    sheet.getRange(`A${rowIndex}`).setFormula(
      `=HYPERLINK("#gid=${getGidForSheet(ss, labSheets[i].name)}", "${labSheets[i].name}")`
    ).setFontColor(CONFIG.COLORS.PRIMARY);
    
    sheet.getRange(`B${rowIndex}:F${rowIndex}`).merge()
      .setValue(labSheets[i].description);
  }
  
  // Add version and instructions
  sheet.getRange("A12").setValue(`Version: ${CONFIG.VERSION}`);
  sheet.getRange("A14").setValue("Getting Started:").setFontWeight("bold");
  sheet.getRange("A15:F15").merge().setValue("Use the SheetsLab menu at the top to access different demonstrations and features.");

  // Set column widths
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 400);
  
  // Add borders and styling
  sheet.getRange("A4:F" + (labSheets.length + 5)).setBorder(true, true, true, true, true, true);
}

/**
 * Creates the UI Elements Lab sheet
 * @param {SpreadsheetApp.Spreadsheet} ss - The active spreadsheet
 */
function createUILabSheet(ss) {
  createLabSheet(ss, CONFIG.SHEETS.UI_LAB, "Explore interactive UI elements", CONFIG.COLORS.SECONDARY);
}

/**
 * Creates the Data Handling Lab sheet
 * @param {SpreadsheetApp.Spreadsheet} ss - The active spreadsheet
 */
function createDataLabSheet(ss) {
  createLabSheet(ss, CONFIG.SHEETS.DATA_LAB, "Explore data handling capabilities", CONFIG.COLORS.ACCENT);
}

/**
 * Creates the Visualization Lab sheet
 * @param {SpreadsheetApp.Spreadsheet} ss - The active spreadsheet
 */
function createVisualizationLabSheet(ss) {
  createLabSheet(ss, CONFIG.SHEETS.VISUALIZATION_LAB, "Explore data visualization features", "#FB8C00"); // Orange
}

/**
 * Creates the Integration Lab sheet
 * @param {SpreadsheetApp.Spreadsheet} ss - The active spreadsheet
 */
function createIntegrationLabSheet(ss) {
  createLabSheet(ss, CONFIG.SHEETS.INTEGRATION_LAB, "Explore integration capabilities", "#8E24AA"); // Purple
}

/**
 * Creates the Formula Lab sheet
 * @param {SpreadsheetApp.Spreadsheet} ss - The active spreadsheet
 */
function createFormulaLabSheet(ss) {
  createLabSheet(ss, CONFIG.SHEETS.FORMULA_LAB, "Explore advanced formulas and functions", "#0288D1"); // Light Blue
}

/**
 * Helper function to create a generic lab sheet
 * @param {SpreadsheetApp.Spreadsheet} ss - The active spreadsheet
 * @param {string} sheetName - Name of the sheet to create
 * @param {string} description - Description of the sheet
 * @param {string} color - Tab color for the sheet
 */
function createLabSheet(ss, sheetName, description, color) {
  // Check if sheet exists, if not create it
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    // Clear existing content if sheet exists
    sheet.clear();
  }
  
  // Set tab color
  sheet.setTabColor(color);
  
  // Set up title and description
  sheet.getRange("A1:F1").merge().setValue(sheetName)
    .setFontSize(18)
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
  
  sheet.getRange("A2:F2").merge().setValue(description)
    .setFontSize(12)
    .setFontStyle("italic")
    .setHorizontalAlignment("center");
  
  // Add home link
  sheet.getRange("A4").setFormula(
    `=HYPERLINK("#gid=${getGidForSheet(ss, CONFIG.SHEETS.HOME)}", "← Back to Home")`
  ).setFontColor(CONFIG.COLORS.PRIMARY);
  
  // Set standard column widths
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 200);
}

/**
 * Helper function to get the GID of a sheet by name
 * @param {SpreadsheetApp.Spreadsheet} ss - The active spreadsheet
 * @param {string} sheetName - Name of the sheet to find
 * @return {number} The GID of the sheet, or 0 if not found
 */
function getGidForSheet(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    return 0;
  }
  
  try {
    // Get the sheet ID from its URL
    const url = ss.getUrl();
    const gid = sheet.getSheetId();
    return gid;
  } catch (e) {
    console.error('Error getting sheet GID:', e);
    return 0;
  }
}