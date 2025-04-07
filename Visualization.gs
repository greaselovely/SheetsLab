/**
 * Visualization.gs
 * Chart and dashboard helpers for SheetsLab
 * 
 * This file contains functions for data visualization, chart creation,
 * and interactive dashboard building.
 * 
 * @version 1.0.0
 */

/**
 * Creates a comprehensive dashboard in the Visualization Lab sheet
 */
function createDashboard() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.VISUALIZATION_LAB);
    
    if (!sheet) {
      SpreadsheetApp.getUi().alert('Visualization Lab sheet not found. Please initialize SheetsLab first.');
      return;
    }
    
    // Get confirmation from user
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'Create Dashboard',
      'This will create a comprehensive dashboard with sample data and multiple charts. Continue?',
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) {
      return;
    }
    
    // Clear existing content (except the header)
    sheet.getRange("A4:Z100").clear();
    
    // Set up dashboard structure
    sheet.getRange("A4:Z4").merge().setValue("Sales Performance Dashboard")
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setFontSize(14)
      .setBackground("#E6E6E6");
    
    // Generate sample data first
    generateDashboardData(sheet);
    
    // Create dashboard components
    createSummaryCards(sheet);
    createSalesChart(sheet);
    createCategoryBreakdown(sheet);
    createRegionalSales(sheet);
    createSalesTable(sheet);
    createTimeSeries(sheet);
    
    // Add instructions
    sheet.getRange("A53:D53").merge().setValue("Dashboard Instructions:")
      .setFontWeight("bold");
    sheet.getRange("A54:D54").merge().setValue("• The data and charts above are for demonstration purposes.");
    sheet.getRange("A55:D55").merge().setValue("• Try using the filters to update the dashboard in real-time.");
    sheet.getRange("A56:D56").merge().setValue("• Click on chart elements to see detailed information.");
    
    // Format the sheet
    sheet.getRange("A4:Z100").applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
    
    // Auto-resize columns
    for (let i = 1; i <= 10; i++) {
      sheet.autoResizeColumn(i);
    }
    
    // Show a success message
    ui.alert('Dashboard created successfully!');
  }
  
  /**
   * Generates sample data for the dashboard
   * @param {SpreadsheetApp.Sheet} sheet - The sheet to generate data in
   */
  function generateDashboardData(sheet) {
    // Create data table headers
    const headers = ["Date", "Product", "Category", "Region", "Units", "Price", "Revenue", "Costs", "Profit", "Margin"];
    sheet.getRange("N5:W5").setValues([headers]).setFontWeight("bold").setBackground("#D9D9D9");
    
    // Define data variations
    const products = [
      "Product A", "Product B", "Product C", "Product D", "Product E",
      "Service X", "Service Y", "Service Z"
    ];
    const categories = ["Hardware", "Software", "Services", "Accessories"];
    const regions = ["North", "South", "East", "West", "Central"];
    
    // Generate random data rows
    const data = [];
    const startDate = new Date(2023, 0, 1); // Jan 1, 2023
    const endDate = new Date(2023, 11, 31); // Dec 31, 2023
    
    for (let i = 1; i <= 50; i++) {
      // Generate a random date between start and end dates
      const randomDate = new Date(startDate.getTime() + Math.random() * (endDate.getTime() - startDate.getTime()));
      
      // Select random product, category and region
      const product = products[Math.floor(Math.random() * products.length)];
      const category = categories[Math.floor(Math.random() * categories.length)];
      const region = regions[Math.floor(Math.random() * regions.length)];
      
      // Generate random sales data
      const units = Math.floor(Math.random() * 100) + 1;
      const price = Math.round((Math.random() * 200 + 50) * 100) / 100; // Between 50 and 250
      const revenue = units * price;
      const costs = Math.round(revenue * (0.4 + Math.random() * 0.3) * 100) / 100; // 40-70% of revenue
      const profit = revenue - costs;
      const margin = profit / revenue;
      
      // Add the row to the data array
      data.push([
        randomDate, // Date
        product, // Product
        category, // Category
        region, // Region
        units, // Units
        price, // Price
        revenue, // Revenue
        costs, // Costs
        profit, // Profit
        margin // Margin
      ]);
    }
    
    // Write the data to the sheet
    sheet.getRange(6, 14, data.length, data[0].length).setValues(data);
    
    // Format the cells
    sheet.getRange("N6:N55").setNumberFormat("yyyy-mm-dd");
    sheet.getRange("S6:V55").setNumberFormat("$#,##0.00");
    sheet.getRange("W6:W55").setNumberFormat("0.00%");
    
    // Add filters
    sheet.getRange("N5:W55").createFilter();
  }
  
  /**
   * Creates summary cards for the dashboard
   * @param {SpreadsheetApp.Sheet} sheet - The sheet to add the summary cards to
   */
  function createSummaryCards(sheet) {
    // Set up the summary cards area
    sheet.getRange("A6:L9").setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
    
    // Create 4 summary cards with formulas linking to the data
    
    // 1. Total Revenue
    sheet.getRange("A6:C6").merge().setValue("Total Revenue").setFontWeight("bold").setBackground("#D9EAD3");
    sheet.getRange("A7:C8").merge().setFormula('=TEXT(SUM(T6:T55), "$#,##0.00")')
      .setFontSize(18)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
    
    // 2. Total Profit
    sheet.getRange("D6:F6").merge().setValue("Total Profit").setFontWeight("bold").setBackground("#D9EAD3");
    sheet.getRange("D7:F8").merge().setFormula('=TEXT(SUM(V6:V55), "$#,##0.00")')
      .setFontSize(18)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
    
    // 3. Average Margin
    sheet.getRange("G6:I6").merge().setValue("Average Margin").setFontWeight("bold").setBackground("#D9EAD3");
    sheet.getRange("G7:I8").merge().setFormula('=TEXT(AVERAGE(W6:W55), "0.0%")')
      .setFontSize(18)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
    
    // 4. Total Units Sold
    sheet.getRange("J6:L6").merge().setValue("Total Units Sold").setFontWeight("bold").setBackground("#D9EAD3");
    sheet.getRange("J7:L8").merge().setFormula('=SUM(R6:R55)')
      .setFontSize(18)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
  }
  
  /**
   * Creates a sales chart for the dashboard
   * @param {SpreadsheetApp.Sheet} sheet - The sheet to add the chart to
   */
  function createSalesChart(sheet) {
    // Create a column chart for revenue and profit
    const chartDataRange = sheet.getRange("N6:N55, T6:V55"); // Date, Revenue, Costs, Profit
    
    const chart = sheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(chartDataRange)
      .setPosition(10, 1, 0, 0)
      .setOption('title', 'Monthly Revenue, Costs and Profit')
      .setOption('legend', {position: 'top'})
      .setOption('height', 300)
      .setOption('width', 600)
      .setOption('series', {
        0: {targetAxisIndex: 0, type: 'bars', color: '#4285F4'}, // Revenue - Blue
        1: {targetAxisIndex: 0, type: 'bars', color: '#EA4335'}, // Costs - Red
        2: {targetAxisIndex: 0, type: 'bars', color: '#34A853'}  // Profit - Green
      })
      .setOption('hAxis', {
        title: 'Month',
        format: 'MMM yyyy'
      })
      .setOption('vAxis', {
        title: 'Amount ($)',
        format: '$#,##0'
      })
      .build();
    
    sheet.insertChart(chart);
    
    // Add a note about the chart
    sheet.getRange("A22:L22").merge()
      .setValue("⚡ The chart above shows the monthly breakdown of revenue, costs, and profit from the sample data.");
  }
  
  /**
   * Creates a category breakdown chart for the dashboard
   * @param {SpreadsheetApp.Sheet} sheet - The sheet to add the chart to
   */
  function createCategoryBreakdown(sheet) {
    // Create a pie chart for revenue by category
    const categoryData = sheet.getRange("P6:P55, T6:T55"); // Category, Revenue
    
    const chart = sheet.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(categoryData)
      .setPosition(10, 8, 0, 0)
      .setOption('title', 'Revenue by Category')
      .setOption('legend', {position: 'right'})
      .setOption('height', 300)
      .setOption('width', 400)
      .setOption('pieSliceText', 'percentage')
      .setOption('colors', ['#4285F4', '#EA4335', '#FBBC04', '#34A853'])
      .build();
    
    sheet.insertChart(chart);
  }
  
  /**
   * Creates a regional sales chart for the dashboard
   * @param {SpreadsheetApp.Sheet} sheet - The sheet to add the chart to
   */
  function createRegionalSales(sheet) {
    // Create a bar chart for revenue by region
    const regionData = sheet.getRange("Q6:Q55, T6:T55"); // Region, Revenue
    
    const chart = sheet.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(regionData)
      .setPosition(25, 1, 0, 0)
      .setOption('title', 'Revenue by Region')
      .setOption('legend', {position: 'none'})
      .setOption('height', 300)
      .setOption('width', 500)
      .setOption('colors', ['#4285F4'])
      .setOption('hAxis', {
        title: 'Revenue ($)',
        format: '$#,##0'
      })
      .setOption('vAxis', {
        title: 'Region'
      })
      .build();
    
    sheet.insertChart(chart);
  }
  
  /**
   * Creates a sales data table for the dashboard
   * @param {SpreadsheetApp.Sheet} sheet - The sheet to add the table to
   */
  function createSalesTable(sheet) {
    // Create a summary table by product and category
    sheet.getRange("A25:C25").merge().setValue("Top Products by Revenue")
      .setFontWeight("bold")
      .setBackground("#E6E6E6");
    
    // Create a QUERY formula to summarize data
    const queryFormula = '=QUERY(N6:W55, "SELECT O, P, SUM(T) ' +
      'WHERE O IS NOT NULL ' +
      'GROUP BY O, P ' +
      'ORDER BY SUM(T) DESC ' +
      'LABEL O \'Product\', P \'Category\', SUM(T) \'Total Revenue\'", 1)';
    
    sheet.getRange("A26").setFormula(queryFormula);
    
    // Format the results table
    sheet.getRange("C26:C40").setNumberFormat("$#,##0.00");
  }
  
  /**
   * Creates a time series chart for the dashboard
   * @param {SpreadsheetApp.Sheet} sheet - The sheet to add the chart to
   */
  function createTimeSeries(sheet) {
    // Create a line chart for profit margin over time
    const timeData = sheet.getRange("N6:N55, W6:W55"); // Date, Margin
    
    const chart = sheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(timeData)
      .setPosition(25, 8, 0, 0)
      .setOption('title', 'Profit Margin Trend')
      .setOption('legend', {position: 'none'})
      .setOption('height', 300)
      .setOption('width', 400)
      .setOption('colors', ['#34A853'])
      .setOption('lineWidth', 3)
      .setOption('hAxis', {
        title: 'Date',
        format: 'MMM yyyy'
      })
      .setOption('vAxis', {
        title: 'Margin',
        format: '0.0%'
      })
      .setOption('trendlines', {
        0: {
          type: 'linear',
          color: '#EA4335',
          lineWidth: 2,
          opacity: 0.5,
          showR2: true,
          visibleInLegend: true
        }
      })
      .build();
    
    sheet.insertChart(chart);
  }
  
  /**
   * Creates an interactive chart that responds to user input
   */
  function createInteractiveChart() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.VISUALIZATION_LAB);
    
    if (!sheet) {
      SpreadsheetApp.getUi().alert('Visualization Lab sheet not found. Please initialize SheetsLab first.');
      return;
    }
    
    // Get confirmation from user
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'Create Interactive Chart',
      'This will create an interactive chart with controls. Continue?',
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) {
      return;
    }
    
    // Clear a section of the sheet for the interactive chart
    sheet.getRange("A60:L90").clear();
    
    // Add title
    sheet.getRange("A60:L60").merge().setValue("Interactive Chart Example")
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setFontSize(14)
      .setBackground("#E6E6E6");
    
    // Add controls section
    sheet.getRange("A62:C62").merge().setValue("Chart Controls")
      .setFontWeight("bold")
      .setBackground("#D9D9D9");
    
    // Add a data range control
    sheet.getRange("A63").setValue("Data Range:");
    sheet.getRange("B63:C63").merge().setValue("All Data");
    
    // Add chart type selector with data validation
    sheet.getRange("A64").setValue("Chart Type:");
    const chartTypes = ["Column", "Line", "Area", "Scatter", "Combo"];
    const chartTypeRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(chartTypes, true)
      .build();
    sheet.getRange("B64").setValue("Column").setDataValidation(chartTypeRule);
    
    // Add dimension selectors
    sheet.getRange("A65").setValue("Dimension:");
    const dimensions = ["Date", "Product", "Category", "Region"];
    const dimensionRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(dimensions, true)
      .build();
    sheet.getRange("B65").setValue("Date").setDataValidation(dimensionRule);
    
    // Add measure selectors
    sheet.getRange("A66").setValue("Measure 1:");
    const measures = ["Units", "Revenue", "Costs", "Profit", "Margin"];
    const measureRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(measures, true)
      .build();
    sheet.getRange("B66").setValue("Revenue").setDataValidation(measureRule);
    
    sheet.getRange("A67").setValue("Measure 2:");
    sheet.getRange("B67").setValue("Profit").setDataValidation(measureRule);
    
    // Add filter for date range
    sheet.getRange("A69").setValue("Date Filter:");
    sheet.getRange("B69").setValue("From:");
    sheet.getRange("C69").setValue(new Date(2023, 0, 1)).setNumberFormat("yyyy-mm-dd");
    sheet.getRange("B70").setValue("To:");
    sheet.getRange("C70").setValue(new Date(2023, 11, 31)).setNumberFormat("yyyy-mm-dd");
    
    // Add a button to refresh the chart (we'll fake this with a checkbox)
    sheet.getRange("B72:C72").merge().setValue("➤ Refresh Chart")
      .setFontWeight("bold")
      .setBackground("#4285F4")
      .setFontColor("white")
      .setHorizontalAlignment("center");
    
    // Add instructions
    sheet.getRange("A74:C78").merge().setValue(
      "Instructions:\n" +
      "1. Select your preferred chart type\n" +
      "2. Choose the dimension (X-axis)\n" +
      "3. Select up to two measures to display\n" +
      "4. Set date filters if needed\n" +
      "5. Click 'Refresh Chart' to update"
    ).setFontStyle("italic");
    
    // Create the chart with QUERY formula data
    // This formula will read the control values and generate an appropriate dataset
    const queryFormula = '=QUERY(N6:W55, ' +
      '"SELECT "&' +
      'IF(B65="Date", "N", IF(B65="Product", "O", IF(B65="Category", "P", "Q")))&' +
      '", "&' +
      'IF(B66="Units", "SUM(R)", IF(B66="Revenue", "SUM(T)", IF(B66="Costs", "SUM(U)", IF(B66="Profit", "SUM(V)", "AVG(W)"))))&' +
      'IF(B67<>"", ", "&IF(B67="Units", "SUM(R)", IF(B67="Revenue", "SUM(T)", IF(B67="Costs", "SUM(U)", IF(B67="Profit", "SUM(V)", "AVG(W)")))), "")&' +
      '" WHERE N >= DATE \'"&TEXT(C69, "yyyy-mm-dd")&"\' AND N <= DATE \'"&TEXT(C70, "yyyy-mm-dd")&"\' "&' +
      '"GROUP BY "&' +
      'IF(B65="Date", "N", IF(B65="Product", "O", IF(B65="Category", "P", "Q")))&' +
      '" ORDER BY "&' +
      'IF(B65="Date", "N", IF(B65="Product", "O", IF(B65="Category", "P", "Q")))&' +
      '" LABEL "&' +
      'IF(B65="Date", "N \'Date\'", IF(B65="Product", "O \'Product\'", IF(B65="Category", "P \'Category\'", "Q \'Region\'")))&' +
      '", "&' +
      'IF(B66="Units", "SUM(R) \'Units\'", IF(B66="Revenue", "SUM(T) \'Revenue\'", IF(B66="Costs", "SUM(U) \'Costs\'", IF(B66="Profit", "SUM(V) \'Profit\'", "AVG(W) \'Margin\'"))))&' +
      'IF(B67<>"", ", "&IF(B67="Units", "SUM(R) \'Units\'", IF(B67="Revenue", "SUM(T) \'Revenue\'", IF(B67="Costs", "SUM(U) \'Costs\'", IF(B67="Profit", "SUM(V) \'Profit\'", "AVG(W) \'Margin\'")))), "")' +
      '", 1)';
    
    // Place the query result starting at E62
    sheet.getRange("E62").setFormula(queryFormula);
    
    // Create a dynamic chart based on the query results
    const dataRange = sheet.getRange("E62:G72"); // Approximate range for the query results
    
    const chartType = Charts.ChartType.COLUMN; // Default chart type
    
    const chart = sheet.newChart()
      .setChartType(chartType)
      .addRange(dataRange)
      .setPosition(62, 8, 0, 0)
      .setOption('title', 'Interactive Chart')
      .setOption('legend', {position: 'top'})
      .setOption('height', 300)
      .setOption('width', 500)
      .build();
    
    sheet.insertChart(chart);
    
    // Add a note about the dynamic nature
    sheet.getRange("A80:L80").merge()
      .setValue("Note: This chart updates based on your selections in the controls above. Click the 'Refresh Chart' button after making changes.");
  
    // Add a note about limitations
    sheet.getRange("A82:L82").merge()
      .setValue("⚠️ In a real app, the 'Refresh Chart' button would be a proper button with a script trigger to rebuild the chart. This example simulates the concept.");
    
    // Show a success message
    ui.alert('Interactive chart created successfully! Try changing the controls to see different chart variations.');
  }
  
  /**
   * Creates a data-driven heatmap visualization
   */
  function createDataHeatmap() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.VISUALIZATION_LAB);
    
    if (!sheet) {
      SpreadsheetApp.getUi().alert('Visualization Lab sheet not found. Please initialize SheetsLab first.');
      return;
    }
    
    // Get confirmation from user
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'Create Data Heatmap',
      'This will create a data-driven heatmap visualization. Continue?',
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) {
      return;
    }
    
    // Clear a section of the sheet for the heatmap
    sheet.getRange("A95:Z120").clear();
    
    // Add title
    sheet.getRange("A95:Z95").merge().setValue("Data-Driven Heatmap Visualization")
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setFontSize(14)
      .setBackground("#E6E6E6");
    
    // Create a heatmap showing Region (rows) vs Category (columns) with Profit values
    // First create the structure
    sheet.getRange("A97:F97").merge().setValue("Region vs Category Profit Heatmap ($)")
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setBackground("#D9D9D9");
    
    // Get unique categories and regions from the data
    const dataRange = sheet.getRange("N6:W55");
    const values = dataRange.getValues();
    
    const categories = [];
    const regions = [];
    
    // Extract unique categories and regions
    for (const row of values) {
      const category = row[2]; // Category is in column P (index 2 in the range)
      const region = row[3];   // Region is in column Q (index 3 in the range)
      
      if (category && !categories.includes(category)) {
        categories.push(category);
      }
      
      if (region && !regions.includes(region)) {
        regions.push(region);
      }
    }
    
    // Sort the arrays
    categories.sort();
    regions.sort();
    
    // Create the heatmap header (categories)
    const headerRow = [""];
    for (const category of categories) {
      headerRow.push(category);
    }
    headerRow.push("Total"); // Add Total column
    
    sheet.getRange(98, 1, 1, headerRow.length).setValues([headerRow])
      .setFontWeight("bold")
      .setBackground("#F3F3F3");
    
    // Create the heatmap rows
    const heatmapData = [];
    const profitByRegionCategory = {};
    const profitByRegion = {};
    const profitByCategory = {};
    let totalProfit = 0;
    
    // Initialize the data structures
    for (const region of regions) {
      profitByRegion[region] = 0;
      profitByRegionCategory[region] = {};
      
      for (const category of categories) {
        profitByRegionCategory[region][category] = 0;
      }
    }
    
    for (const category of categories) {
      profitByCategory[category] = 0;
    }
    
    // Calculate the profit sums
    for (const row of values) {
      const category = row[2]; // Category is in column P (index 2 in the range)
      const region = row[3];   // Region is in column Q (index 3 in the range)
      const profit = row[8];   // Profit is in column V (index 8 in the range)
      
      if (category && region && !isNaN(profit)) {
        profitByRegionCategory[region][category] += profit;
        profitByRegion[region] += profit;
        profitByCategory[category] += profit;
        totalProfit += profit;
      }
    }
    
    // Build the heatmap data
    for (const region of regions) {
      const row = [region]; // First cell is the region name
      
      for (const category of categories) {
        row.push(profitByRegionCategory[region][category]);
      }
      
      row.push(profitByRegion[region]); // Add the total for the region
      heatmapData.push(row);
    }
    
    // Add a total row
    const totalRow = ["Total"];
    for (const category of categories) {
      totalRow.push(profitByCategory[category]);
    }
    totalRow.push(totalProfit);
    heatmapData.push(totalRow);
    
    // Write the data to the sheet
    sheet.getRange(99, 1, heatmapData.length, heatmapData[0].length).setValues(heatmapData);
    
    // Format the data cells
    sheet.getRange(99, 2, heatmapData.length, heatmapData[0].length - 1).setNumberFormat("$#,##0.00");
    
    // Format the total row and column
    sheet.getRange(99 + regions.length, 1, 1, heatmapData[0].length)
      .setFontWeight("bold")
      .setBackground("#F3F3F3");
    
    sheet.getRange(99, heatmapData[0].length, regions.length, 1)
      .setFontWeight("bold")
      .setBackground("#F3F3F3");
    
    // Apply conditional formatting (heatmap colors)
    const range = sheet.getRange(99, 2, regions.length, categories.length);
    
    // Find the min and max values for scaling
    let minValue = Number.MAX_VALUE;
    let maxValue = Number.MIN_VALUE;
    
    for (let i = 0; i < regions.length; i++) {
      for (let j = 0; j < categories.length; j++) {
        const value = heatmapData[i][j + 1];
        minValue = Math.min(minValue, value);
        maxValue = Math.max(maxValue, value);
      }
    }
    
    // Create the gradient rule for positive values (green)
    const positiveRule = SpreadsheetApp.newConditionalFormatRule()
      .setGradientMaxpoint("#57BB8A") // Dark green
      .setGradientMidpoint("#E0F3DB") // Light green
      .setGradientMinpoint("#FFFFFF") // White
      .setRanges([range])
      .build();
      
    // If there are negative values, we need a different approach with two rules
    if (minValue < 0) {
      // Remove the positive rule and create two rules for positive and negative values
      const negativeRange = sheet.getRange(99, 2, regions.length, categories.length);
      
      const negativeRule = SpreadsheetApp.newConditionalFormatRule()
        .whenNumberLessThan(0)
        .setBackground("#F4C7C3") // Light red for negative
        .setRanges([negativeRange])
        .build();
      
      const positiveRule2 = SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThan(0)
        .setGradientMaxpoint("#57BB8A") // Dark green
        .setGradientMinpoint("#FFFFFF") // White
        .setRanges([range])
        .build();
      
      // Apply both rules
      const rules = sheet.getConditionalFormatRules();
      rules.push(negativeRule);
      rules.push(positiveRule2);
      sheet.setConditionalFormatRules(rules);
    } else {
      // Apply the single gradient rule for all positive values
      const rules = sheet.getConditionalFormatRules();
      rules.push(positiveRule);
      sheet.setConditionalFormatRules(rules);
    }
    
    // Add a legend explaining the heatmap
    sheet.getRange("A" + (101 + regions.length) + ":F" + (101 + regions.length)).merge()
      .setValue("Heatmap Legend:");
    
    sheet.getRange("A" + (102 + regions.length) + ":F" + (102 + regions.length)).merge()
      .setValue("• Darker green = Higher profit");
    
    if (minValue < 0) {
      sheet.getRange("A" + (103 + regions.length) + ":F" + (103 + regions.length)).merge()
        .setValue("• Red = Negative profit (loss)");
    }
    
    // Add borders to the heatmap
    sheet.getRange(98, 1, regions.length + 1, categories.length + 2).setBorder(true, true, true, true, true, true);
    
    // Add an explanation of how it works
    sheet.getRange("H97:L97").merge().setValue("About This Visualization:")
      .setFontWeight("bold");
    
    sheet.getRange("H98:L103").merge().setValue(
      "This heatmap shows the profit distribution across different regions and categories. " +
      "The color intensity indicates the relative profit amount, with darker green representing higher profit. " +
      "The totals are shown in the rightmost column and bottom row.\n\n" +
      "This visualization allows you to quickly identify the most and least profitable combinations " +
      "of regions and categories."
    ).setWrap(true);
    
    // Show a success message
    ui.alert('Data heatmap created successfully!');
  }