/**
 * Config.gs
 * Configuration settings and constants for SheetsLab
 * 
 * This file contains all the configuration variables and settings
 * used throughout the SheetsLab project.
 * 
 * @version 1.0.0
 */

/**
 * Global configuration object for SheetsLab
 */
const CONFIG = {
    // Project information
    PROJECT_NAME: "SheetsLab",
    VERSION: "1.0.0",
    GITHUB_URL: "https://github.com/yourusername/sheetslab", // Update with actual repo
    
    // Sheet names
    SHEETS: {
      HOME: "Home",
      UI_LAB: "UI Elements Lab",
      DATA_LAB: "Data Handling Lab",
      VISUALIZATION_LAB: "Visualization Lab",
      INTEGRATION_LAB: "Integration Lab",
      FORMULA_LAB: "Formula Lab"
    },
    
    // UI Configuration
    UI: {
      SIDEBAR_WIDTH: 300,
      SIDEBAR_TITLE: "SheetsLab Navigator",
      DIALOG_WIDTH: 600,
      DIALOG_HEIGHT: 400
    },
    
    // Colors (using Google Material Design palette)
    COLORS: {
      PRIMARY: "#4285F4", // Google Blue
      SECONDARY: "#0F9D58", // Google Green
      ACCENT: "#DB4437", // Google Red
      LIGHT_BG: "#F5F5F5",
      DARK_TEXT: "#212121",
      LIGHT_TEXT: "#FFFFFF"
    },
    
    // Demo data settings
    DEMO_DATA: {
      ROWS: 100,
      CATEGORIES: ["Products", "Services", "Hardware", "Software", "Support"]
    }
  };
  
  /**
   * Function to get configuration values
   * @param {string} key - The configuration key to retrieve
   * @return {*} - The configuration value
   */
  function getConfig(key) {
    return CONFIG[key];
  }