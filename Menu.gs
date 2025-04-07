/**
 * Menu.gs
 * Custom menu creation and handling for SheetsLab
 * 
 * This file contains functions to create and manage the custom SheetsLab
 * menu in the Google Sheets UI.
 * 
 * @version 1.0.0
 */

/**
 * Creates the SheetsLab custom menu in the Google Sheets UI
 * This function is automatically called when the spreadsheet opens
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  // Create the main SheetsLab menu
  ui.createMenu(CONFIG.PROJECT_NAME)
    .addItem('Initialize SheetsLab', 'initializeSheetsLab')
    .addItem('Initialize SheetsLab with Sample Data', 'initializeWithSampleData')
    .addSeparator()
    .addSubMenu(ui.createMenu('UI Elements Lab')
      .addItem('Show Navigation Sidebar', 'showNavigationSidebar')
      .addItem('Show Simple Modal Dialog', 'showSimpleDialog')
      .addItem('Show Advanced Form Dialog', 'showAdvancedFormDialog')
      .addItem('Show Toast Notification', 'showToastNotification')
      .addItem('Show Progress Indicator', 'showProgressDemo'))
    .addSubMenu(ui.createMenu('Data Handling Lab')
      .addItem('Generate Sample Data', 'generateSampleData')
      .addItem('Apply Data Validation Rules', 'applyDataValidation')
      .addItem('Create Advanced Filter Views', 'createAdvancedFilters')
      .addItem('Run Data Transformation', 'runDataTransformation'))
    .addSubMenu(ui.createMenu('Visualization Lab')
      .addItem('Create Dashboard', 'createDashboard')
      .addItem('Generate Interactive Chart', 'createInteractiveChart')
      .addItem('Create Data-Driven Heatmap', 'createDataHeatmap'))
    .addSubMenu(ui.createMenu('Integration Lab')
      .addItem('Connect to External API', 'showApiConnectionDialog')
      .addItem('Import External Data', 'showDataImportOptions')
      .addItem('Email Automation Demo', 'showEmailAutomationDialog'))
    .addSubMenu(ui.createMenu('Formula Lab')
      .addItem('Show Array Formula Examples', 'navigateToArrayFormulas')
      .addItem('Show Query Function Examples', 'navigateToQueryFunctions')
      .addItem('Show Custom Function Examples', 'navigateToCustomFunctions'))
    .addSeparator()
    .addItem('Show SheetsLab Navigator', 'showNavigationSidebar')
    .addItem('Show Welcome Screen', 'showWelcomeDialog')
    .addItem('About SheetsLab', 'showAboutDialog')
    .addToUi();
    
  // Check if this is the first time opening the spreadsheet
  checkFirstRun();
}


/**
 * Shows a toast notification with a message
 * @param {string} message - The message to display
 * @param {string} title - The title for the toast (optional)
 * @param {number} timeout - Timeout in seconds (optional, default 5)
 */
function showToast(message, title = CONFIG.PROJECT_NAME, timeout = 5) {
  SpreadsheetApp.getActiveSpreadsheet().toast(message, title, timeout);
}

/**
 * Shows the about dialog with information about SheetsLab
 */
function showAboutDialog() {
  const htmlOutput = HtmlService.createHtmlOutput(`
    <div style="font-family: Arial, sans-serif; padding: 20px;">
      <h2 style="color: ${CONFIG.COLORS.PRIMARY};">${CONFIG.PROJECT_NAME}</h2>
      <h3>Version ${CONFIG.VERSION}</h3>
      <p>A comprehensive showcase of Google Sheets capabilities.</p>
      <p>This project demonstrates a wide range of advanced features and techniques
      that can be implemented in Google Sheets, turning it into a powerful application platform.</p>
      <p><strong>GitHub:</strong> <a href="${CONFIG.GITHUB_URL}" target="_blank">${CONFIG.GITHUB_URL}</a></p>
      <hr>
      <p><em>Created as an open-source knowledge base for the Google Sheets community.</em></p>
    </div>
  `)
    .setWidth(CONFIG.UI.DIALOG_WIDTH)
    .setHeight(300);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'About SheetsLab');
}

/**
 * Simple demonstration of a toast notification
 */
function showToastNotification() {
  showToast('This is a toast notification example!', 'UI Demo', 3);
}

/**
 * Shows a simple modal dialog
 */
function showSimpleDialog() {
  const htmlOutput = HtmlService.createHtmlOutput(`
    <div style="font-family: Arial, sans-serif; padding: 20px;">
      <h2 style="color: ${CONFIG.COLORS.PRIMARY};">Simple Modal Dialog</h2>
      <p>This is a basic modal dialog created using HtmlService.</p>
      <p>Modal dialogs can be used to:</p>
      <ul>
        <li>Display information to users</li>
        <li>Collect user input through forms</li>
        <li>Show confirmations before important actions</li>
        <li>Display help content or instructions</li>
      </ul>
      <button onclick="google.script.host.close()">Close</button>
    </div>
  `)
    .setWidth(CONFIG.UI.DIALOG_WIDTH)
    .setHeight(CONFIG.UI.DIALOG_HEIGHT);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Simple Dialog Example');
}

/**
 * Placeholder for functions that will be implemented in other files
 * These functions are referenced in the menu but defined elsewhere
 */
function showNavigationSidebar() {
  // This will be implemented in UI.gs
  showToast('Navigation sidebar function will be implemented in UI.gs');
}

function showAdvancedFormDialog() {
  // This will be implemented in UI.gs
  showToast('Advanced form dialog function will be implemented in UI.gs');
}

function showProgressDemo() {
  // This will be implemented in UI.gs
  showToast('Progress indicator function will be implemented in UI.gs');
}

function generateSampleData() {
  // This will be implemented in DataTools.gs
  showToast('Generate sample data function will be implemented in DataTools.gs');
}

function applyDataValidation() {
  // This will be implemented in DataTools.gs
  showToast('Data validation function will be implemented in DataTools.gs');
}

function createAdvancedFilters() {
  // This will be implemented in DataTools.gs
  showToast('Advanced filters function will be implemented in DataTools.gs');
}

function runDataTransformation() {
  // This will be implemented in DataTools.gs
  showToast('Data transformation function will be implemented in DataTools.gs');
}

function createDashboard() {
  // This will be implemented in Visualization.gs
  showToast('Dashboard creation function will be implemented in Visualization.gs');
}

function createInteractiveChart() {
  // This will be implemented in Visualization.gs
  showToast('Interactive chart function will be implemented in Visualization.gs');
}

function createDataHeatmap() {
  // This will be implemented in Visualization.gs
  showToast('Data heatmap function will be implemented in Visualization.gs');
}

function showApiConnectionDialog() {
  // This will be implemented in Integration.gs
  showToast('API connection dialog function will be implemented in Integration.gs');
}

function showDataImportOptions() {
  // This will be implemented in Integration.gs
  showToast('Data import options function will be implemented in Integration.gs');
}

function showEmailAutomationDialog() {
  // This will be implemented in Integration.gs
  showToast('Email automation dialog function will be implemented in Integration.gs');
}

function navigateToArrayFormulas() {
  // This will be implemented in FormulaExamples.gs
  showToast('Array formulas navigation function will be implemented in FormulaExamples.gs');
}

function navigateToQueryFunctions() {
  // This will be implemented in FormulaExamples.gs
  showToast('Query functions navigation function will be implemented in FormulaExamples.gs');
}

function navigateToCustomFunctions() {
  // This will be implemented in FormulaExamples.gs
  showToast('Custom functions navigation function will be implemented in FormulaExamples.gs');
}

/**
 * Initializes SheetsLab and generates all sample data
 * This provides a one-click setup for users to quickly see all capabilities
 */
function initializeWithSampleData() {
  const ui = SpreadsheetApp.getUi();
  
  // Confirm with the user
  const response = ui.alert(
    'Initialize SheetsLab with Sample Data',
    'This will create all sheets and generate sample data for all labs. It may take a moment to complete. Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  // Show a toast to indicate the process has started
  showToast('Starting complete initialization...', 'SheetsLab Setup', 10);
  
  try {
    // First initialize the basic sheets
    initializeSheetsLab(false); // Pass false to suppress the success message
    
    // Now generate sample data for each lab
    
    // 1. Generate sample data for Data Handling Lab
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.getSheetByName(CONFIG.SHEETS.DATA_LAB).activate();
    generateSampleData();
    applyDataValidation();
    createAdvancedFilters();
    runDataTransformation();
    
    // 2. Generate dashboard for Visualization Lab
    ss.getSheetByName(CONFIG.SHEETS.VISUALIZATION_LAB).activate();
    createDashboard();
    createInteractiveChart();
    createDataHeatmap();
    
    // 3. Fetch sample API data for Integration Lab
    ss.getSheetByName(CONFIG.SHEETS.INTEGRATION_LAB).activate();
    fetchApiData({api: 'iss-location'}); // Use the ISS location API as default
    
    // 4. Create formula examples for Formula Lab
    ss.getSheetByName(CONFIG.SHEETS.FORMULA_LAB).activate();
    createArrayFormulaExamples(ss.getActiveSheet());
    
    // 5. Show some UI examples in UI Lab
    ss.getSheetByName(CONFIG.SHEETS.UI_LAB).activate();
    
    // Return to the Home sheet
    ss.getSheetByName(CONFIG.SHEETS.HOME).activate();
    
    // Show success message
    ui.alert(
      'SheetsLab Fully Initialized',
      'All sheets have been created and populated with sample data. You can now explore all the capabilities through the menu and navigation.',
      ui.ButtonSet.OK
    );
    
    // Suggest opening the sidebar
    const sidebarResponse = ui.alert(
      'Open Navigator',
      'Would you like to open the SheetsLab Navigator sidebar?',
      ui.ButtonSet.YES_NO
    );
    
    if (sidebarResponse === ui.Button.YES) {
      showNavigationSidebar();
    }
  } catch (error) {
    // Show error message if something goes wrong
    ui.alert(
      'Error during initialization',
      'An error occurred while setting up SheetsLab: ' + error.toString(),
      ui.ButtonSet.OK
    );
  }
}

/**
 * Shows the welcome dialog with information and setup options
 */
function showWelcomeDialog() {
  // Create the HTML for the dialog
  const htmlOutput = HtmlService.createTemplateFromFile('WelcomeDialog')
    .evaluate()
    .setWidth(600)
    .setHeight(500)
    .setTitle('Welcome to SheetsLab');
  
  // Display the dialog
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Welcome to SheetsLab');
}

/**
 * Checks if this is the first time running the spreadsheet
 * If it is, show the welcome dialog
 */
function checkFirstRun() {
  const props = PropertiesService.getDocumentProperties();
  const hasRun = props.getProperty('sheetsLabInitialized');
  
  if (!hasRun) {
    // This is the first run, show the welcome dialog
    // Since we can't use setTimeout in Apps Script, we'll show the dialog immediately
    showWelcomeDialog();
    
    // Mark as initialized
    props.setProperty('sheetsLabInitialized', 'true');
  }
}