// Visualization.gs
// SheetsLab - Chart and dashboard helpers

/**
 * Creates a sample dashboard chart based on data in the DATA sheet.
 * The chart is inserted into the DASHBOARD sheet.
 */
function createDashboardChart() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var dataSheet = ss.getSheetByName(CONFIG.SHEETS.DATA);
    var dashboardSheet = ss.getSheetByName(CONFIG.SHEETS.DASHBOARD);
    
    if (!dataSheet || !dashboardSheet) {
      throw new Error("Required sheets are missing. Ensure both DATA and DASHBOARD sheets exist.");
    }
    
    // Clear existing content in the DASHBOARD sheet
    dashboardSheet.clear();
    
    // Assume data starts at A1 with headers in the DATA sheet
    var dataRange = dataSheet.getDataRange();
    
    // Build a column chart using Google Charts
    var chart = dashboardSheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(dataRange)
      .setPosition(2, 2, 0, 0)
      .setOption('title', 'Sample Data Chart')
      .setOption('legend', {position: 'bottom'})
      .build();
      
    dashboardSheet.insertChart(chart);
  }
  
  /**
   * Updates the dashboard chart by re-creating it.
   */
  function updateDashboardChart() {
    createDashboardChart();
  }
  