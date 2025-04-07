/**
 * UI.gs
 * UI elements like sidebars and modals for SheetsLab
 * 
 * This file contains functions to create and manage various UI elements
 * such as sidebars, modal dialogs, and other interactive elements.
 * 
 * @version 1.0.0
 */

/**
 * Shows the navigation sidebar
 */
function showNavigationSidebar() {
    // Create the HTML for the sidebar
    const htmlOutput = HtmlService.createTemplateFromFile('NavigationSidebar')
      .evaluate()
      .setTitle(CONFIG.UI.SIDEBAR_TITLE)
      .setWidth(CONFIG.UI.SIDEBAR_WIDTH);
    
    // Display the sidebar
    SpreadsheetApp.getUi().showSidebar(htmlOutput);
  }
  
  /**
   * Gets the HTML content for the navigation sidebar
   * This is called from the NavigationSidebar.html file
   * @return {Object} Object containing sidebar data
   */
  function getNavigationData() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Get all sheet names and their GIDs
    const sheets = ss.getSheets();
    const labSheets = [];
    
    // Build the data for each lab sheet
    for (const sheet of sheets) {
      const sheetName = sheet.getName();
      // Only include SheetsLab sheets
      if (Object.values(CONFIG.SHEETS).includes(sheetName)) {
        labSheets.push({
          name: sheetName,
          gid: sheet.getSheetId(),
          isActive: sheet.isActive()
        });
      }
    }
    
    // Return the data for the sidebar
    return {
      projectName: CONFIG.PROJECT_NAME,
      version: CONFIG.VERSION,
      sheets: labSheets,
      colors: CONFIG.COLORS
    };
  }
  
  /**
   * Shows an advanced form dialog for data input
   */
  function showAdvancedFormDialog() {
    // Create the HTML for the dialog
    const htmlOutput = HtmlService.createTemplateFromFile('AdvancedFormDialog')
      .evaluate()
      .setWidth(CONFIG.UI.DIALOG_WIDTH)
      .setHeight(CONFIG.UI.DIALOG_HEIGHT);
    
    // Display the dialog
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Advanced Form Example');
  }
  
  /**
   * Processes form data submitted from the advanced form dialog
   * @param {Object} formData - The data submitted from the form
   * @return {Object} Result object with success status and message
   */
  function processFormData(formData) {
    try {
      const sheet = SpreadsheetApp.getActiveSpreadsheet()
        .getSheetByName(CONFIG.SHEETS.UI_LAB);
      
      // Make sure the sheet exists
      if (!sheet) {
        return {
          success: false,
          message: 'UI Lab sheet not found. Please initialize SheetsLab first.'
        };
      }
      
      // Find the next empty row
      const lastRow = Math.max(sheet.getLastRow(), 10);
      const nextRow = lastRow + 1;
      
      // Write headers if they don't exist
      if (lastRow < 11) {
        sheet.getRange('A11').setValue('Timestamp');
        sheet.getRange('B11').setValue('Name');
        sheet.getRange('C11').setValue('Email');
        sheet.getRange('D11').setValue('Category');
        sheet.getRange('E11').setValue('Priority');
        sheet.getRange('F11').setValue('Description');
        sheet.getRange('A11:F11').setFontWeight('bold');
      }
      
      // Write form data to the sheet
      sheet.getRange(nextRow, 1).setValue(new Date());
      sheet.getRange(nextRow, 2).setValue(formData.name);
      sheet.getRange(nextRow, 3).setValue(formData.email);
      sheet.getRange(nextRow, 4).setValue(formData.category);
      sheet.getRange(nextRow, 5).setValue(formData.priority);
      sheet.getRange(nextRow, 6).setValue(formData.description);
      
      // Format the row
      sheet.getRange(nextRow, 1, 1, 6).setBorder(true, true, true, true, true, true);
      
      // Conditional formatting based on priority
      if (formData.priority === 'High') {
        sheet.getRange(nextRow, 5).setBackground('#F4C7C3'); // Light red
      } else if (formData.priority === 'Medium') {
        sheet.getRange(nextRow, 5).setBackground('#FCE8B2'); // Light yellow
      } else {
        sheet.getRange(nextRow, 5).setBackground('#B7E1CD'); // Light green
      }
      
      return {
        success: true,
        message: 'Form data saved successfully!'
      };
    } catch (error) {
      console.error('Error processing form data:', error);
      return {
        success: false,
        message: 'Error processing form data: ' + error.toString()
      };
    }
  }
  
  /**
   * Shows a progress indicator demo
   */
  function showProgressDemo() {
    // Create the HTML for the dialog
    const htmlOutput = HtmlService.createTemplateFromFile('ProgressDemo')
      .evaluate()
      .setWidth(500)
      .setHeight(350);
    
    // Display the dialog
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Progress Indicator Demo');
  }
  
  /**
   * Simulates a long-running process with progress updates
   * @param {number} steps - Total number of steps in the process
   * @return {Object} Status update with current progress
   */
  function runLongProcess(steps) {
    // Get the cache to store progress
    const cache = CacheService.getUserCache();
    
    // Start the process
    cache.put('progress', '0');
    
    // Create a trigger to continue the process
    ScriptApp.newTrigger('continueProcess')
      .timeBased()
      .after(1000) // 1 second delay
      .create();
    
    // Store the total number of steps and current step
    cache.put('totalSteps', steps.toString());
    cache.put('currentStep', '0');
    
    return {
      status: 'started',
      progress: 0,
      message: 'Process started...'
    };
  }
  
  /**
   * Continues the long-running process
   * Called by the time-based trigger
   */
  function continueProcess() {
    // Get the cache with progress data
    const cache = CacheService.getUserCache();
    
    // Get the current progress
    let currentStep = parseInt(cache.get('currentStep') || '0');
    const totalSteps = parseInt(cache.get('totalSteps') || '10');
    
    // Increment the progress
    currentStep++;
    
    // Calculate the percentage
    const progress = Math.round((currentStep / totalSteps) * 100);
    
    // Update the cache
    cache.put('currentStep', currentStep.toString());
    cache.put('progress', progress.toString());
    
    // If not done, create another trigger
    if (currentStep < totalSteps) {
      ScriptApp.newTrigger('continueProcess')
        .timeBased()
        .after(1000) // 1 second delay
        .create();
    }
    
    // Clean up this trigger
    const triggers = ScriptApp.getProjectTriggers();
    for (const trigger of triggers) {
      if (trigger.getHandlerFunction() === 'continueProcess') {
        ScriptApp.deleteTrigger(trigger);
        break;
      }
    }
  }
  
  /**
   * Gets the current progress of the long-running process
   * @return {Object} Current progress information
   */
  function getProgress() {
    const cache = CacheService.getUserCache();
    const progress = parseInt(cache.get('progress') || '0');
    const currentStep = parseInt(cache.get('currentStep') || '0');
    const totalSteps = parseInt(cache.get('totalSteps') || '10');
    
    // Determine the status message
    let message = 'Processing...';
    let status = 'running';
    
    if (progress >= 100) {
      message = 'Process completed successfully!';
      status = 'completed';
    } else if (currentStep === 0) {
      message = 'Initializing process...';
    } else {
      message = `Processing step ${currentStep} of ${totalSteps}...`;
    }
    
    return {
      status: status,
      progress: progress,
      message: message
    };
  }
  
  /**
   * Includes an HTML file in another HTML file
   * @param {string} filename - Name of the HTML file to include
   * @return {string} The content of the HTML file
   */
  function include(filename) {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  }