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
      .addItem('About SheetsLab', 'showAboutDialog')
      .addToUi();
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