/**
 * DataTools.gs
 * Data handling utilities for SheetsLab
 * 
 * This file contains functions for data handling, validation,
 * filtering, and transformation operations.
 * 
 * @version 1.0.0
 */

/**
 * Generates sample data in the Data Handling Lab sheet
 */
function generateSampleData() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.DATA_LAB);
    
    if (!sheet) {
      SpreadsheetApp.getUi().alert('Data Handling Lab sheet not found. Please initialize SheetsLab first.');
      return;
    }
    
    // Clear any existing data in the sample data area
    sheet.getRange("A6:J106").clear();
    
    // Set headers
    const headers = ["ID", "Date", "Category", "Product", "Quantity", "Price", "Total", "Customer", "Region", "Status"];
    sheet.getRange("A6:J6").setValues([headers]).setFontWeight("bold");
    
    // Define data variations for random generation
    const categories = CONFIG.DEMO_DATA.CATEGORIES;
    const products = {
      "Products": ["Widget A", "Widget B", "Widget C", "Premium Widget", "Economy Widget"],
      "Services": ["Consultation", "Installation", "Maintenance", "Training", "Support"],
      "Hardware": ["Server", "Workstation", "Laptop", "Tablet", "Accessories"],
      "Software": ["Operating System", "Office Suite", "Database", "Security", "Utilities"],
      "Support": ["Basic Support", "Priority Support", "24/7 Support", "On-site Support", "Remote Support"]
    };
    const customers = ["Acme Corp", "XYZ Inc", "123 Industries", "Best Company", "Super Enterprises", 
                        "Global Solutions", "Local Business", "Tech Innovators", "Creative Agency", "Data Systems"];
    const regions = ["North", "South", "East", "West", "Central"];
    const statuses = ["Completed", "Pending", "Processing", "Cancelled", "On Hold"];
    
    // Generate random data rows
    const data = [];
    const startDate = new Date(2023, 0, 1); // Jan 1, 2023
    const endDate = new Date(2023, 11, 31); // Dec 31, 2023
    
    for (let i = 1; i <= 100; i++) {
      // Generate a random date between start and end dates
      const randomDate = new Date(startDate.getTime() + Math.random() * (endDate.getTime() - startDate.getTime()));
      
      // Select random category
      const category = categories[Math.floor(Math.random() * categories.length)];
      
      // Select random product based on category
      const product = products[category][Math.floor(Math.random() * products[category].length)];
      
      // Generate random quantity and price
      const quantity = Math.floor(Math.random() * 20) + 1;
      const price = Math.round((Math.random() * 500 + 10) * 100) / 100; // Between 10 and 510
      const total = quantity * price;
      
      // Select random customer, region, and status
      const customer = customers[Math.floor(Math.random() * customers.length)];
      const region = regions[Math.floor(Math.random() * regions.length)];
      const status = statuses[Math.floor(Math.random() * statuses.length)];
      
      // Add the row to the data array
      data.push([
        i, // ID
        randomDate, // Date
        category, // Category
        product, // Product
        quantity, // Quantity
        price, // Price
        total, // Total
        customer, // Customer
        region, // Region
        status // Status
      ]);
    }
    
    // Write the data to the sheet
    sheet.getRange(7, 1, data.length, data[0].length).setValues(data);
    
    // Format the cells
    sheet.getRange("B7:B106").setNumberFormat("yyyy-mm-dd");
    sheet.getRange("F7:G106").setNumberFormat("$0.00");
    
    // Add a description above the data
    sheet.getRange("A4:J4").merge().setValue("Sample Sales Data").setFontWeight("bold");
    sheet.getRange("A5:J5").merge().setValue("This data is randomly generated for demonstration purposes.");
    
    // Auto-resize columns
    for (let i = 1; i <= 10; i++) {
      sheet.autoResizeColumn(i);
    }
    
    // Add borders
    sheet.getRange("A6:J106").setBorder(true, true, true, true, true, true);
    
    // Show a success message
    SpreadsheetApp.getUi().alert('Sample data generated successfully!');
  }
  
  /**
   * Applies data validation rules to the sample data
   */
  function applyDataValidation() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.DATA_LAB);
    
    if (!sheet) {
      SpreadsheetApp.getUi().alert('Data Handling Lab sheet not found. Please initialize SheetsLab first.');
      return;
    }
    
    // Check if sample data exists
    if (sheet.getRange("A6").getValue() !== "ID") {
      SpreadsheetApp.getUi().alert('Please generate sample data first.');
      return;
    }
    
    // Get confirmation from user
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'Apply Data Validation',
      'This will apply various data validation rules to the sample data. Continue?',
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) {
      return;
    }
    
    // 1. Category validation - dropdown list
    const categoryRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(CONFIG.DEMO_DATA.CATEGORIES, true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange("C7:C106").setDataValidation(categoryRule);
    
    // 2. Quantity validation - positive integers only
    const quantityRule = SpreadsheetApp.newDataValidation()
      .requireNumberGreaterThan(0)
      .requireWholeNumber()
      .setAllowInvalid(false)
      .build();
    sheet.getRange("E7:E106").setDataValidation(quantityRule);
    
    // 3. Price validation - positive numbers only
    const priceRule = SpreadsheetApp.newDataValidation()
      .requireNumberGreaterThan(0)
      .setAllowInvalid(false)
      .build();
    sheet.getRange("F7:F106").setDataValidation(priceRule);
    
    // 4. Status validation - dropdown list with statuses
    const statuses = ["Completed", "Pending", "Processing", "Cancelled", "On Hold"];
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(statuses, true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange("J7:J106").setDataValidation(statusRule);
    
    // 5. Total validation - formula check
    // Set the G column to be calculated from E and F
    for (let i = 7; i <= 106; i++) {
      const formula = `=E${i}*F${i}`;
      sheet.getRange(`G${i}`).setFormula(formula);
    }
    sheet.getRange("G7:G106").protect().setDescription("Total (Calculated Field)");
    
    // Add validation descriptions
    sheet.getRange("A115:B115").merge().setValue("Data Validation Rules Applied:").setFontWeight("bold");
    sheet.getRange("A116:B116").merge().setValue("Category: Must be one of the predefined categories");
    sheet.getRange("A117:B117").merge().setValue("Quantity: Must be a positive whole number");
    sheet.getRange("A118:B118").merge().setValue("Price: Must be a positive number");
    sheet.getRange("A119:B119").merge().setValue("Total: Protected calculated field (Quantity × Price)");
    sheet.getRange("A120:B120").merge().setValue("Status: Must be one of the predefined statuses");
    
    // Add color coding to cells with validation
    sheet.getRange("C6").setBackground("#D9EAD3"); // Green for Category header
    sheet.getRange("E6").setBackground("#D9EAD3"); // Green for Quantity header
    sheet.getRange("F6").setBackground("#D9EAD3"); // Green for Price header
    sheet.getRange("G6").setBackground("#FCE5CD"); // Orange for Total header (calculated)
    sheet.getRange("J6").setBackground("#D9EAD3"); // Green for Status header
    
    // Show a success message
    ui.alert('Data validation rules have been applied successfully!');
  }
  
  /**
   * Creates advanced filter views for the sample data
   */
  function createAdvancedFilters() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.DATA_LAB);
    
    if (!sheet) {
      SpreadsheetApp.getUi().alert('Data Handling Lab sheet not found. Please initialize SheetsLab first.');
      return;
    }
    
    // Check if sample data exists
    if (sheet.getRange("A6").getValue() !== "ID") {
      SpreadsheetApp.getUi().alert('Please generate sample data first.');
      return;
    }
    
    // Get confirmation from user
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'Create Advanced Filters',
      'This will create several filter views for the sample data. Continue?',
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) {
      return;
    }
    
    // Remove any existing filter views
    const existingFilters = ss.getFilterViews();
    for (let i = 0; i < existingFilters.length; i++) {
      if (existingFilters[i].getRange().getSheet().getName() === CONFIG.SHEETS.DATA_LAB) {
        existingFilters[i].remove();
      }
    }
    
    // Get the data range
    const dataRange = sheet.getRange("A6:J106");
    
    // 1. Create a filter for high-value orders (Total > $1000)
    const highValueFilter = ss.addFilterView();
    highValueFilter.setRange(dataRange);
    highValueFilter.setTitle("High Value Orders");
    
    // Apply the filter criteria: Total > 1000
    highValueFilter.getFilterCriteria()
      .setColumnPosition(6) // G column, 0-indexed from the range start
      .whenNumberGreaterThan(1000);
    
    // 2. Create a filter for pending orders
    const pendingFilter = ss.addFilterView();
    pendingFilter.setRange(dataRange);
    pendingFilter.setTitle("Pending Orders");
    
    // Apply the filter criteria: Status = "Pending"
    pendingFilter.getFilterCriteria()
      .setColumnPosition(9) // J column, 0-indexed from the range start
      .whenTextEqualTo("Pending");
    
    // 3. Create a filter for hardware products
    const hardwareFilter = ss.addFilterView();
    hardwareFilter.setRange(dataRange);
    hardwareFilter.setTitle("Hardware Products");
    
    // Apply the filter criteria: Category = "Hardware"
    hardwareFilter.getFilterCriteria()
      .setColumnPosition(2) // C column, 0-indexed from the range start
      .whenTextEqualTo("Hardware");
    
    // 4. Create a filter for Q1 orders
    const q1Filter = ss.addFilterView();
    q1Filter.setRange(dataRange);
    q1Filter.setTitle("Q1 Orders");
    
    // Apply the filter criteria: Date between Jan 1 and Mar 31
    const q1Start = new Date(2023, 0, 1); // Jan 1, 2023
    const q1End = new Date(2023, 2, 31); // Mar 31, 2023
    
    q1Filter.getFilterCriteria()
      .setColumnPosition(1) // B column, 0-indexed from the range start
      .whenDateAfter(q1Start)
      .whenDateBefore(q1End);
    
    // 5. Create a filter combining multiple criteria
    const combinedFilter = ss.addFilterView();
    combinedFilter.setRange(dataRange);
    combinedFilter.setTitle("High Value North Region");
    
    // Apply filter for Total > 800
    combinedFilter.getFilterCriteria()
      .setColumnPosition(6) // G column, 0-indexed from the range start
      .whenNumberGreaterThan(800);
    
    // Apply filter for Region = "North"
    combinedFilter.getFilterCriteria()
      .setColumnPosition(8) // I column, 0-indexed from the range start
      .whenTextEqualTo("North");
    
    // Add a note about the created filters
    sheet.getRange("A125:D125").merge().setValue("Advanced Filter Views Created:").setFontWeight("bold");
    sheet.getRange("A126:D126").merge().setValue("• High Value Orders: Orders with total value > $1,000");
    sheet.getRange("A127:D127").merge().setValue("• Pending Orders: Orders with status 'Pending'");
    sheet.getRange("A128:D128").merge().setValue("• Hardware Products: Orders in the 'Hardware' category");
    sheet.getRange("A129:D129").merge().setValue("• Q1 Orders: Orders from January to March");
    sheet.getRange("A130:D130").merge().setValue("• High Value North Region: Orders > $800 in the North region");
    sheet.getRange("A131:D131").merge().setValue("(Access these filters from the filter button in the toolbar)");
    
    // Show a success message
    ui.alert('Advanced filters have been created successfully!');
  }
  
  /**
   * Runs a data transformation on the sample data
   */
  function runDataTransformation() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.DATA_LAB);
    
    if (!sheet) {
      SpreadsheetApp.getUi().alert('Data Handling Lab sheet not found. Please initialize SheetsLab first.');
      return;
    }
    
    // Check if sample data exists
    if (sheet.getRange("A6").getValue() !== "ID") {
      SpreadsheetApp.getUi().alert('Please generate sample data first.');
      return;
    }
    
    // Get confirmation from user
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'Run Data Transformation',
      'This will create a pivot table and data summary based on the sample data. Continue?',
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) {
      return;
    }
    
    // Clear any existing transformation output
    sheet.getRange("M4:S50").clear();
    
    // Add a title for the transformation section
    sheet.getRange("M4:S4").merge().setValue("Data Transformation Results").setFontWeight("bold");
    sheet.getRange("M5:S5").merge().setValue("Automatically generated summary from sample data");
    
    // Create a basic summary
    sheet.getRange("M7:N7").setValues([["Total Orders:", "100"]]);
    
    // Create a mini pivot table for Category totals
    sheet.getRange("M9:O9").setValues([["Category", "Count", "Total Sales"]]).setFontWeight("bold");
    
    // Get all the data from the sample
    const data = sheet.getRange("A7:J106").getValues();
    
    // Initialize category summaries
    const categorySummary = {};
    const categories = CONFIG.DEMO_DATA.CATEGORIES;
    
    // Initialize with zeros
    for (const category of categories) {
      categorySummary[category] = {
        count: 0,
        totalSales: 0
      };
    }
    
    // Process each row in the data
    for (const row of data) {
      const category = row[2]; // Category is in column C (index 2)
      const total = row[6];    // Total is in column G (index 6)
      
      if (category && categorySummary[category]) {
        categorySummary[category].count++;
        categorySummary[category].totalSales += total;
      }
    }
    
    // Write category summary to the sheet
    let rowIndex = 10;
    for (const category of categories) {
      sheet.getRange(rowIndex, 13, 1, 3).setValues([[
        category,
        categorySummary[category].count,
        categorySummary[category].totalSales
      ]]);
      rowIndex++;
    }
    
    // Format the total sales column
    sheet.getRange("O10:O" + (9 + categories.length)).setNumberFormat("$#,##0.00");
    
    // Add a total row
    sheet.getRange(rowIndex, 13, 1, 3).setValues([[
      "TOTAL",
      data.length,
      "=SUM(O10:O" + (rowIndex - 1) + ")"
    ]]);
    sheet.getRange(rowIndex, 13, 1, 3).setFontWeight("bold");
    sheet.getRange("O" + rowIndex).setNumberFormat("$#,##0.00");
    
    // Create a status summary
    rowIndex += 3;
    sheet.getRange(rowIndex, 13, 1, 3).setValues([["Status", "Count", "Percentage"]]).setFontWeight("bold");
    rowIndex++;
    
    // Get unique statuses
    const statuses = [...new Set(data.map(row => row[9]))]; // Status is in column J (index 9)
    
    // Count orders by status
    const statusSummary = {};
    for (const status of statuses) {
      if (status) {
        statusSummary[status] = 0;
      }
    }
    
    for (const row of data) {
      const status = row[9]; // Status is in column J (index 9)
      if (status && statusSummary.hasOwnProperty(status)) {
        statusSummary[status]++;
      }
    }
    
    // Write status summary to the sheet
    for (const status of statuses) {
      if (status) {
        const count = statusSummary[status];
        const percentage = count / data.length;
        
        sheet.getRange(rowIndex, 13, 1, 3).setValues([[
          status,
          count,
          percentage
        ]]);
        rowIndex++;
      }
    }
    
    // Format the percentage column
    sheet.getRange("O" + (rowIndex - statuses.length) + ":O" + (rowIndex - 1)).setNumberFormat("0.0%");
    
    // Create a simple chart
    const chartBuilder = sheet.newChart();
    const chartRange = sheet.getRange("M10:O" + (9 + categories.length));
    
    const chart = chartBuilder
      .addRange(chartRange)
      .setChartType(Charts.ChartType.COLUMN)
      .setPosition(rowIndex + 2, 13, 0, 0)
      .setOption('title', 'Sales by Category')
      .setOption('legend', {position: 'top'})
      .setOption('hAxis', {title: 'Category'})
      .setOption('vAxis', {title: 'Amount ($)'})
      .setOption('series', {
        0: {targetAxisIndex: 1}, // Count on right axis
        1: {targetAxisIndex: 0}  // Total sales on left axis
      })
      .setOption('vAxes', {
        0: {title: 'Sales ($)', format: '$#,##0'},
        1: {title: 'Count'}
      })
      .build();
    
    sheet.insertChart(chart);
    
    // Auto-resize the transformation columns
    for (let i = 13; i <= 19; i++) {
      sheet.autoResizeColumn(i);
    }
    
    // Add borders to the tables
    sheet.getRange("M9:O" + (10 + categories.length)).setBorder(true, true, true, true, true, true);
    sheet.getRange("M" + (rowIndex - statuses.length - 1) + ":O" + (rowIndex - 1)).setBorder(true, true, true, true, true, true);
    
    // Show a success message
    ui.alert('Data transformation completed successfully!');
  }