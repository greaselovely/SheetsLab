/**
 * FormulaExamples.gs
 * Advanced formula demonstrations for SheetsLab
 * 
 * This file contains functions to demonstrate advanced formula
 * techniques in Google Sheets.
 * 
 * @version 1.0.0
 */

/**
 * Navigates to the Array Formulas section in the Formula Lab sheet
 */
function navigateToArrayFormulas() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.FORMULA_LAB);
    
    if (!sheet) {
      SpreadsheetApp.getUi().alert('Formula Lab sheet not found. Please initialize SheetsLab first.');
      return;
    }
    
    // Activate the sheet
    sheet.activate();
    
    // Check if the array formulas section exists
    const findRange = sheet.createTextFinder('Array Formula Examples').findNext();
    
    if (findRange) {
      // Navigate to the existing section
      const row = findRange.getRow();
      sheet.setActiveRange(sheet.getRange(row, 1));
    } else {
      // Create the array formulas section
      createArrayFormulaExamples(sheet);
    }
  }
  
  /**
   * Navigates to the Query Function section in the Formula Lab sheet
   */
  function navigateToQueryFunctions() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.FORMULA_LAB);
    
    if (!sheet) {
      SpreadsheetApp.getUi().alert('Formula Lab sheet not found. Please initialize SheetsLab first.');
      return;
    }
    
    // Activate the sheet
    sheet.activate();
    
    // Check if the query functions section exists
    const findRange = sheet.createTextFinder('QUERY Function Examples').findNext();
    
    if (findRange) {
      // Navigate to the existing section
      const row = findRange.getRow();
      sheet.setActiveRange(sheet.getRange(row, 1));
    } else {
      // Create the query functions section
      createQueryFunctionExamples(sheet);
    }
  }
  
  /**
   * Navigates to the Custom Function section in the Formula Lab sheet
   */
  function navigateToCustomFunctions() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.FORMULA_LAB);
    
    if (!sheet) {
      SpreadsheetApp.getUi().alert('Formula Lab sheet not found. Please initialize SheetsLab first.');
      return;
    }
    
    // Activate the sheet
    sheet.activate();
    
    // Check if the custom functions section exists
    const findRange = sheet.createTextFinder('Custom Function Examples').findNext();
    
    if (findRange) {
      // Navigate to the existing section
      const row = findRange.getRow();
      sheet.setActiveRange(sheet.getRange(row, 1));
    } else {
      // Create the custom functions section
      createCustomFunctionExamples(sheet);
    }
  }
  
  /**
   * Creates array formula examples in the Formula Lab sheet
   * @param {SpreadsheetApp.Sheet} sheet - The Formula Lab sheet
   */
  function createArrayFormulaExamples(sheet) {
    // Clear some space for the examples
    sheet.getRange("A6:Z100").clear();
    
    // Create sample data table
    sheet.getRange("M6:P6").setValues([["Month", "Region", "Sales", "Costs"]]).setFontWeight("bold");
    
    // Generate sample data
    const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun"];
    const regions = ["North", "South", "East", "West"];
    const sampleData = [];
    
    for (let i = 0; i < months.length; i++) {
      for (let j = 0; j < regions.length; j++) {
        const sales = Math.round(Math.random() * 5000 + 1000);
        const costs = Math.round(sales * (0.4 + Math.random() * 0.2));
        sampleData.push([months[i], regions[j], sales, costs]);
      }
    }
    
    // Write sample data
    sheet.getRange(7, 13, sampleData.length, 4).setValues(sampleData);
    
    // Format numbers
    sheet.getRange(7, 15, sampleData.length, 2).setNumberFormat("$#,##0");
    
    // Add title
    sheet.getRange("A6:F6").merge().setValue("Array Formula Examples")
      .setFontWeight("bold")
      .setFontSize(14)
      .setBackground("#E6E6E6");
    
    // Example 1: Basic Array Formula
    sheet.getRange("A8:A8").setValue("Example 1: Calculate Profit with a Single Array Formula")
      .setFontWeight("bold");
    
    sheet.getRange("A9:A9").setValue("Formula:");
    sheet.getRange("B9:F9").merge().setValue("=ArrayFormula(M7:M & \": \" & N7:N & \" - Profit: \" & O7:O-P7:P)")
      .setFontStyle("italic");
    
    sheet.getRange("A10:A10").setValue("Result:");
    sheet.getRange("B10").setFormula("=ArrayFormula(M7:M & \": \" & N7:N & \" - Profit: \" & O7:O-P7:P)");
    
    // Example 2: SUMIF with Arrays
    sheet.getRange("A13:A13").setValue("Example 2: Calculate Sum of Sales by Region with SUMIF")
      .setFontWeight("bold");
    
    sheet.getRange("A14:A14").setValue("Formula:");
    sheet.getRange("B14:F14").merge().setValue("=ArrayFormula(UNIQUE(N7:N) & \": \" & SUMIF(N7:N, UNIQUE(N7:N), O7:O))")
      .setFontStyle("italic");
    
    sheet.getRange("A15:A15").setValue("Result:");
    sheet.getRange("B15").setFormula("=ArrayFormula(UNIQUE(N7:N) & \": \" & TEXT(SUMIF(N7:N, UNIQUE(N7:N), O7:O), \"$#,##0\"))");
    
    // Example 3: Conditional Array Calculations
    sheet.getRange("A18:A18").setValue("Example 3: Flag High-Performing Regions (Sales > $20,000)")
      .setFontWeight("bold");
    
    sheet.getRange("A19:A19").setValue("Formula:");
    sheet.getRange("B19:F19").merge().setValue("=ArrayFormula(IF(N7:N=\"\", \"\", IF(SUMIF(N7:N, N7:N, O7:O) > 20000, N7:N & \" - High Performer\", N7:N & \" - Standard\")))")
      .setFontStyle("italic");
    
    sheet.getRange("A20:A20").setValue("Result:");
    sheet.getRange("B20").setFormula("=ArrayFormula(IF(N7:N=\"\", \"\", IF(SUMIF(N7:N, N7:N, O7:O) > 20000, N7:N & \" - High Performer\", N7:N & \" - Standard\")))");
    
    // Example 4: Multi-column calculations
    sheet.getRange("A23:A23").setValue("Example 4: Create a Summary Table with Multiple Array Calculations")
      .setFontWeight("bold");
    
    sheet.getRange("A24:D24").setValues([["Region", "Total Sales", "Total Costs", "Profit Margin"]]).setFontWeight("bold");
    
    sheet.getRange("A25").setFormula("=UNIQUE(N7:N)");
    sheet.getRange("B25").setFormula("=ArrayFormula(SUMIF(N7:N, A25:A, O7:O))");
    sheet.getRange("C25").setFormula("=ArrayFormula(SUMIF(N7:N, A25:A, P7:P))");
    sheet.getRange("D25").setFormula("=ArrayFormula((B25:B-C25:C)/B25:B)");
    
    // Format the summary table
    sheet.getRange("B25:C28").setNumberFormat("$#,##0");
    sheet.getRange("D25:D28").setNumberFormat("0.00%");
    
    // Example 5: Dynamic ranges
    sheet.getRange("A30:A30").setValue("Example 5: Dynamic Range Processing with FILTER")
      .setFontWeight("bold");
    
    sheet.getRange("A31:A31").setValue("Formula:");
    sheet.getRange("B31:F31").merge().setValue("=ArrayFormula(AVERAGE(FILTER(O7:O, N7:N=\"North\")))")
      .setFontStyle("italic");
    
    sheet.getRange("A32:A32").setValue("Result:");
    sheet.getRange("B32").setFormula("=ArrayFormula(\"Average Sales in North Region: \" & TEXT(AVERAGE(FILTER(O7:O, N7:N=\"North\")), \"$#,##0.00\"))");
    
    // Example 6: Complex conditional logic
    sheet.getRange("A34:A34").setValue("Example 6: Complex Conditional Analysis")
      .setFontWeight("bold");
    
    sheet.getRange("A35:A35").setValue("Formula:");
    sheet.getRange("B35:F35").merge().setValue("=ArrayFormula(IF(O7:O>3000, IF(P7:P<O7:O*0.5, \"High Profit\", \"Standard\"), \"Low Sales\"))")
      .setFontStyle("italic");
    
    sheet.getRange("A36:D36").setValues([["Month", "Region", "Sales", "Performance"]]).setFontWeight("bold");
    sheet.getRange("A37").setFormula("=M7:M");
    sheet.getRange("B37").setFormula("=N7:N");
    sheet.getRange("C37").setFormula("=O7:O");
    sheet.getRange("D37").setFormula("=ArrayFormula(IF(O7:O>3000, IF(P7:P<O7:O*0.5, \"High Profit\", \"Standard\"), \"Low Sales\"))");
    
    // Format the performance table
    sheet.getRange("C37:C60").setNumberFormat("$#,##0");
    
    // Add conditional formatting to the performance column
    const performanceRange = sheet.getRange("D37:D60");
    const rules = sheet.getConditionalFormatRules();
    
    // High Profit rule (green)
    const highProfitRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("High Profit")
      .setBackground("#D9EAD3")
      .setRanges([performanceRange])
      .build();
    
    // Low Sales rule (light red)
    const lowSalesRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("Low Sales")
      .setBackground("#F4C7C3")
      .setRanges([performanceRange])
      .build();
    
    rules.push(highProfitRule);
    rules.push(lowSalesRule);
    sheet.setConditionalFormatRules(rules);
    
    // Add notes about array formulas
    sheet.getRange("A40:E42").merge().setValue(
      "Notes about Array Formulas:\n" +
      "• Array formulas perform calculations on multiple cells at once\n" +
      "• They can replace complex multi-cell formulas with a single formula\n" +
      "• In Google Sheets, use ArrayFormula() to enable array calculations"
    ).setFontStyle("italic");
    
    // Format the entire examples section
    sheet.getRange("A6:F42").setBorder(true, true, true, true, false, false);
    
    // Auto-resize columns
    for (let i = 1; i <= 6; i++) {
      sheet.autoResizeColumn(i);
    }
    for (let i = 13; i <= 16; i++) {
      sheet.autoResizeColumn(i);
    }
    
    // Set the active range to the title
    sheet.setActiveRange(sheet.getRange("A6"));
  
  
  /**
   * Creates QUERY function examples in the Formula Lab sheet
   * @param {SpreadsheetApp.Sheet} sheet - The Formula Lab sheet
   */
  function createQueryFunctionExamples(sheet) {
    // Clear some space for the examples
    sheet.getRange("A6:Z100").clear();
    
    // Create sample data table with more fields
    sheet.getRange("M6:S6").setValues([
      ["Date", "Country", "Product", "Category", "Quantity", "Price", "Revenue"]
    ]).setFontWeight("bold");
    
    // Generate sample data
    const dates = [
      new Date(2023, 0, 15), new Date(2023, 1, 2), new Date(2023, 1, 28), 
      new Date(2023, 2, 10), new Date(2023, 3, 5), new Date(2023, 4, 12),
      new Date(2023, 5, 20), new Date(2023, 6, 1), new Date(2023, 6, 30)
    ];
    
    const countries = ["USA", "Canada", "UK", "Germany", "France", "Japan", "Australia"];
    const products = ["Laptop", "Phone", "Tablet", "Monitor", "Keyboard", "Mouse", "Headphones"];
    const categories = ["Electronics", "Accessories", "Peripherals"];
    
    const sampleData = [];
    
    for (let i = 0; i < 50; i++) {
      const date = dates[Math.floor(Math.random() * dates.length)];
      const country = countries[Math.floor(Math.random() * countries.length)];
      const product = products[Math.floor(Math.random() * products.length)];
      
      // Assign a sensible category
      let category;
      if (product === "Laptop" || product === "Phone" || product === "Tablet") {
        category = "Electronics";
      } else if (product === "Keyboard" || product === "Mouse") {
        category = "Peripherals";
      } else {
        category = "Accessories";
      }
      
      const quantity = Math.floor(Math.random() * 10) + 1;
      const price = Math.round((Math.random() * 500 + 50) * 100) / 100;
      const revenue = quantity * price;
      
      sampleData.push([date, country, product, category, quantity, price, revenue]);
    }
    
    // Write sample data
    sheet.getRange(7, 13, sampleData.length, 7).setValues(sampleData);
    
    // Format numbers and dates
    sheet.getRange(7, 13, sampleData.length, 1).setNumberFormat("yyyy-MM-dd");
    sheet.getRange(7, 18, sampleData.length, 2).setNumberFormat("$#,##0.00");
    
    // Add title
    sheet.getRange("A6:F6").merge().setValue("QUERY Function Examples")
      .setFontWeight("bold")
      .setFontSize(14)
      .setBackground("#E6E6E6");
    
    // Example 1: Basic QUERY with selection and filter
    sheet.getRange("A8:A8").setValue("Example 1: Basic QUERY with Column Selection and Filtering")
      .setFontWeight("bold");
    
    sheet.getRange("A9:A9").setValue("Formula:");
    sheet.getRange("B9:F9").merge().setValue("=QUERY(M6:S56, \"SELECT N, P, SUM(R) WHERE N != '' GROUP BY N, P ORDER BY SUM(R) DESC LABEL SUM(R) 'Total Revenue'\")")
      .setFontStyle("italic");
    
    sheet.getRange("A10:A10").setValue("Result:");
    sheet.getRange("A11:C11").setValues([["Country", "Product", "Total Revenue"]]).setFontWeight("bold");
    sheet.getRange("A12").setFormula("=QUERY(M6:S56, \"SELECT N, P, SUM(R) WHERE N != '' GROUP BY N, P ORDER BY SUM(R) DESC LABEL SUM(R) 'Total Revenue'\")");
    
    // Format the result
    sheet.getRange("C12:C25").setNumberFormat("$#,##0.00");
    
    // Example 2: Date filtering
    sheet.getRange("A28:A28").setValue("Example 2: Date Filtering and Aggregation")
      .setFontWeight("bold");
    
    sheet.getRange("A29:A29").setValue("Formula:");
    sheet.getRange("B29:F29").merge().setValue("=QUERY(M6:S56, \"SELECT MONTH(M), SUM(R) WHERE M >= date '2023-02-01' AND M <= date '2023-06-30' GROUP BY MONTH(M) ORDER BY MONTH(M) LABEL MONTH(M) 'Month', SUM(R) 'Monthly Revenue'\")")
      .setFontStyle("italic");
    
    sheet.getRange("A30:A30").setValue("Result:");
    sheet.getRange("A31:B31").setValues([["Month", "Monthly Revenue"]]).setFontWeight("bold");
    sheet.getRange("A32").setFormula("=QUERY(M6:S56, \"SELECT MONTH(M), SUM(R) WHERE M >= date '2023-02-01' AND M <= date '2023-06-30' GROUP BY MONTH(M) ORDER BY MONTH(M) LABEL MONTH(M) 'Month', SUM(R) 'Monthly Revenue'\")");
    
    // Format the result
    sheet.getRange("B32:B40").setNumberFormat("$#,##0.00");
    
    // Example 3: Complex filtering and calculation
    sheet.getRange("A42:A42").setValue("Example 3: Complex Filtering with Multiple Conditions")
      .setFontWeight("bold");
    
    sheet.getRange("A43:A43").setValue("Formula:");
    sheet.getRange("B43:F43").merge().setValue("=QUERY(M6:S56, \"SELECT O, Q, AVG(Q), SUM(R) WHERE (O = 'Electronics' OR O = 'Peripherals') AND R > 100 GROUP BY O, Q HAVING AVG(Q) > 3 ORDER BY O, AVG(Q) DESC LABEL AVG(Q) 'Avg Quantity', SUM(R) 'Total Revenue'\")")
      .setFontStyle("italic");
    
    sheet.getRange("A44:A44").setValue("Result:");
    sheet.getRange("A45:D45").setValues([["Category", "Product", "Avg Quantity", "Total Revenue"]]).setFontWeight("bold");
    sheet.getRange("A46").setFormula("=QUERY(M6:S56, \"SELECT O, P, AVG(Q), SUM(R) WHERE (O = 'Electronics' OR O = 'Peripherals') AND R > 100 GROUP BY O, P HAVING AVG(Q) > 3 ORDER BY O, AVG(Q) DESC LABEL AVG(Q) 'Avg Quantity', SUM(R) 'Total Revenue'\")");
    
    // Format the result
    sheet.getRange("C46:C55").setNumberFormat("0.00");
    sheet.getRange("D46:D55").setNumberFormat("$#,##0.00");
    
    // Example 4: Pivot-like query
    sheet.getRange("A57:A57").setValue("Example 4: QUERY as a Pivot Table")
      .setFontWeight("bold");
    
    sheet.getRange("A58:A58").setValue("Formula:");
    sheet.getRange("B58:F58").merge().setValue("=QUERY(M6:S56, \"SELECT O, SUM(R) PIVOT N LABEL O 'Category'\")")
      .setFontStyle("italic");
    
    sheet.getRange("A59:A59").setValue("Result (Revenue by Category and Country):");
    sheet.getRange("A60").setFormula("=QUERY(M6:S56, \"SELECT O, SUM(R) PIVOT N LABEL O 'Category'\")");
    
    // Format the result
    sheet.getRange("B60:H65").setNumberFormat("$#,##0.00");
    
    // Add notes about QUERY function
    sheet.getRange("A70:F73").merge().setValue(
      "Notes about the QUERY Function:\n" +
      "• QUERY uses a SQL-like syntax that's powerful for data analysis\n" +
      "• It can filter, sort, group, pivot, and aggregate data in a single formula\n" +
      "• The first parameter is the data range; the second is the query string\n" +
      "• It supports various functions like SUM(), AVG(), COUNT(), MAX(), MIN(), etc."
    ).setFontStyle("italic");
    
    // Format the entire examples section
    sheet.getRange("A6:F73").setBorder(true, true, true, true, false, false);
    
    // Auto-resize columns
    for (let i = 1; i <= 6; i++) {
      sheet.autoResizeColumn(i);
    }
    for (let i = 13; i <= 19; i++) {
      sheet.autoResizeColumn(i);
    }
    
    // Set the active range to the title
    sheet.setActiveRange(sheet.getRange("A6"));
  }
  
  /**
   * Creates custom function examples in the Formula Lab sheet
   * @param {SpreadsheetApp.Sheet} sheet - The Formula Lab sheet
   */
  function createCustomFunctionExamples(sheet) {
    // Clear some space for the examples
    sheet.getRange("A6:Z100").clear();
    
    // Add title
    sheet.getRange("A6:F6").merge().setValue("Custom Function Examples")
      .setFontWeight("bold")
      .setFontSize(14)
      .setBackground("#E6E6E6");
    
    // Explanation
    sheet.getRange("A7:F10").merge().setValue(
      "Custom functions are JavaScript functions that you can create in Google Apps Script " + 
      "to extend Google Sheets with your own formulas. They can be used directly in cells " +
      "just like built-in functions. Below are examples of useful custom functions."
    ).setWrap(true);
    
    // Example 1: Basic custom function
    sheet.getRange("A12:A12").setValue("Example 1: FORMAT_PHONE")
      .setFontWeight("bold");
    
    sheet.getRange("A13:F15").merge().setValue(
      "This function formats a 10-digit number as a US phone number (XXX) XXX-XXXX.\n\n" +
      "Usage: =FORMAT_PHONE(1234567890)"
    ).setWrap(true);
    
    sheet.getRange("A16:B16").setValues([["Input", "Result"]]).setFontWeight("bold");
    sheet.getRange("A17").setValue("1234567890");
    sheet.getRange("B17").setValue("This will display as (123) 456-7890 when the custom function is implemented");
    
    // Display the code
    sheet.getRange("A19:A19").setValue("Function Code:")
      .setFontWeight("bold");
    
    sheet.getRange("A20:F29").merge().setValue(
      "/**\n" +
      " * Formats a number as a US phone number.\n" +
      " *\n" +
      " * @param {number} number The phone number to format.\n" +
      " * @return {string} The formatted phone number.\n" +
      " * @customfunction\n" +
      " */\n" +
      "function FORMAT_PHONE(number) {\n" +
      "  // Convert to string and remove non-numeric characters\n" +
      "  const numStr = String(number).replace(/\\D/g, '');\n" +
      "  \n" +
      "  // Check if it's a valid 10-digit number\n" +
      "  if (numStr.length !== 10) {\n" +
      "    return 'Invalid phone number';\n" +
      "  }\n" +
      "  \n" +
      "  // Format as (XXX) XXX-XXXX\n" +
      "  return '(' + numStr.substring(0, 3) + ') ' + \n" +
      "         numStr.substring(3, 6) + '-' + \n" +
      "         numStr.substring(6);\n" +
      "}"
    ).setFontFamily("Courier New");
    
    // Add notes about custom functions
    sheet.getRange("A76:F80").merge().setValue(
      "Notes about Custom Functions:\n" +
      "• Custom functions must be written in Google Apps Script (JavaScript)\n" +
      "• They can accept parameters and return various data types including arrays\n" +
      "• Use the @customfunction JSDoc tag to make them appear in the function autocomplete\n" +
      "• Custom functions run on Google's servers, not in the browser\n" +
      "• They have limitations such as not being able to access certain services when called from a cell"
    ).setFontStyle("italic");
    
    // Example 4: Advanced custom function with 2D arrays
    sheet.getRange("A82:A82").setValue("Example 4: TRANSPOSE_AND_SUM")
      .setFontWeight("bold");
    
    sheet.getRange("A83:F85").merge().setValue(
      "This advanced function takes a range of cells, transposes it, and adds a totals row and column.\n\n" +
      "Usage: =TRANSPOSE_AND_SUM(A1:C3)"
    ).setWrap(true);
    
    // Create sample data for the example
    sheet.getRange("H82:J84").setValues([
      [10, 20, 30],
      [40, 50, 60],
      [70, 80, 90]
    ]);
    
    sheet.getRange("H81:J81").setValues([["Sample Input Range:"]]).setFontWeight("bold");
    
    sheet.getRange("A86:A86").setValue("Result would look like (when implemented):");
    sheet.getRange("A87:D90").setValues([
      [10, 40, 70, "Row Sum"],
      [20, 50, 80, "Row Sum"],
      [30, 60, 90, "Row Sum"],
      ["Col Sum", "Col Sum", "Col Sum", "Total"]
    ]);
    
    // Display the code
    sheet.getRange("A92:A92").setValue("Function Code:")
      .setFontWeight("bold");
    
    sheet.getRange("A93:F108").merge().setValue(
      "/**\n" +
      " * Transposes a range and adds sum totals for rows and columns.\n" +
      " *\n" +
      " * @param {Range} range The input range.\n" +
      " * @return {Array<Array<number|string>>} The transposed range with totals.\n" +
      " * @customfunction\n" +
      " */\n" +
      "function TRANSPOSE_AND_SUM(range) {\n" +
      "  // Get values from the range\n" +
      "  const values = range.map(row => row.slice());\n" +
      "  \n" +
      "  // Transpose the values (rows become columns, columns become rows)\n" +
      "  const transposed = [];\n" +
      "  for (let i = 0; i < values[0].length; i++) {\n" +
      "    transposed[i] = [];\n" +
      "    for (let j = 0; j < values.length; j++) {\n" +
      "      transposed[i][j] = values[j][i];\n" +
      "    }\n" +
      "  }\n" +
      "  \n" +
      "  // Add row sums\n" +
      "  for (let i = 0; i < transposed.length; i++) {\n" +
      "    let rowSum = 0;\n" +
      "    for (let j = 0; j < transposed[i].length; j++) {\n" +
      "      rowSum += Number(transposed[i][j]);\n" +
      "    }\n" +
      "    transposed[i].push('Row Sum: ' + rowSum);\n" +
      "  }\n" +
      "  \n" +
      "  // Add column sum row\n" +
      "  const colSumRow = [];\n" +
      "  let totalSum = 0;\n" +
      "  \n" +
      "  // Calculate each column sum\n" +
      "  for (let j = 0; j < transposed[0].length - 1; j++) {\n" +
      "    let colSum = 0;\n" +
      "    for (let i = 0; i < transposed.length; i++) {\n" +
      "      colSum += Number(transposed[i][j]);\n" +
      "    }\n" +
      "    colSumRow.push('Col Sum: ' + colSum);\n" +
      "    totalSum += colSum;\n" +
      "  }\n" +
      "  \n" +
      "  // Add the total sum\n" +
      "  colSumRow.push('Total: ' + totalSum);\n" +
      "  \n" +
      "  // Add the column sum row to the result\n" +
      "  transposed.push(colSumRow);\n" +
      "  \n" +
      "  return transposed;\n" +
      "}"
    ).setFontFamily("Courier New");
    
    // Format the entire examples section
    sheet.getRange("A6:F108").setBorder(true, true, true, true, false, false);
    
    // Auto-resize columns
    for (let i = 1; i <= 10; i++) {
      sheet.autoResizeColumn(i);
    }
    
    // Set the active range to the title
    sheet.setActiveRange(sheet.getRange("A6"));
  }
    
    // Example 2: Working with arrays
    sheet.getRange("A31:A31").setValue("Example 2: EXTRACT_EMAILS")
      .setFontWeight("bold");
    
    sheet.getRange("A32:F34").merge().setValue(
      "This function extracts all email addresses from a text and returns them as an array.\n\n" +
      "Usage: =EXTRACT_EMAILS(\"Contact us at support@example.com or sales@example.com\")"
    ).setWrap(true);
    
    sheet.getRange("A35:B35").setValues([["Input", "Result"]]).setFontWeight("bold");
    sheet.getRange("A36").setValue("Contact us at support@example.com or sales@example.com");
    sheet.getRange("B36").setValue("This will return an array of ['support@example.com', 'sales@example.com'] when implemented");
    
    // Display the code
    sheet.getRange("A38:A38").setValue("Function Code:")
      .setFontWeight("bold");
    
    sheet.getRange("A39:F48").merge().setValue(
      "/**\n" +
      " * Extracts email addresses from text.\n" +
      " *\n" +
      " * @param {string} text The text to search for emails.\n" +
      " * @return {string[]} An array of extracted email addresses.\n" +
      " * @customfunction\n" +
      " */\n" +
      "function EXTRACT_EMAILS(text) {\n" +
      "  if (typeof text !== 'string') {\n" +
      "    return 'Input must be text';\n" +
      "  }\n" +
      "  \n" +
      "  // Regular expression for matching emails\n" +
      "  const emailRegex = /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\\.[a-zA-Z]{2,}/g;\n" +
      "  \n" +
      "  // Extract all matches\n" +
      "  const matches = text.match(emailRegex);\n" +
      "  \n" +
      "  // Return matches or empty array if none found\n" +
      "  return matches || [];\n" +
      "}"
    ).setFontFamily("Courier New");
    
    // Example 3: Custom function with multiple arguments
    sheet.getRange("A50:A50").setValue("Example 3: DATE_DIFF_BUSINESS_DAYS")
      .setFontWeight("bold");
    
    sheet.getRange("A51:F53").merge().setValue(
      "This function calculates the number of business days (excluding weekends) between two dates.\n\n" +
      "Usage: =DATE_DIFF_BUSINESS_DAYS(startDate, endDate)"
    ).setWrap(true);
    
    sheet.getRange("A54:C54").setValues([["Start Date", "End Date", "Result"]]).setFontWeight("bold");
    sheet.getRange("A55").setValue(new Date(2023, 3, 3)); // April 3, 2023 (Monday)
    sheet.getRange("B55").setValue(new Date(2023, 3, 14)); // April 14, 2023 (Friday)
    sheet.getRange("C55").setValue("This will return 10 business days when implemented");
    
    // Format dates
    sheet.getRange("A55:B55").setNumberFormat("yyyy-MM-dd");
    
    // Display the code
    sheet.getRange("A57:A57").setValue("Function Code:")
      .setFontWeight("bold");
    
    sheet.getRange("A58:F74").merge().setValue(
      "/**\n" +
      " * Calculates the number of business days between two dates.\n" +
      " *\n" +
      " * @param {Date} startDate The start date.\n" +
      " * @param {Date} endDate The end date.\n" +
      " * @return {number} The number of business days.\n" +
      " * @customfunction\n" +
      " */\n" +
      "function DATE_DIFF_BUSINESS_DAYS(startDate, endDate) {\n" +
      "  // Validate inputs\n" +
      "  if (!(startDate instanceof Date) || !(endDate instanceof Date)) {\n" +
      "    return 'Both inputs must be dates';\n" +
      "  }\n" +
      "  \n" +
      "  // Ensure start date is before end date\n" +
      "  if (startDate > endDate) {\n" +
      "    const temp = startDate;\n" +
      "    startDate = endDate;\n" +
      "    endDate = temp;\n" +
      "  }\n" +
      "  \n" +
      "  // Clone dates to avoid modifying the originals\n" +
      "  let currentDate = new Date(startDate.getTime());\n" +
      "  const lastDate = new Date(endDate.getTime());\n" +
      "  \n" +
      "  // Set to start of day\n" +
      "  currentDate.setHours(0, 0, 0, 0);\n" +
      "  lastDate.setHours(0, 0, 0, 0);\n" +
      "  \n" +
      "  let businessDays = 0;\n" +
      "  \n" +
      "  // Count business days\n" +
      "  while (currentDate <= lastDate) {\n" +
      "    const dayOfWeek = currentDate.getDay();\n" +
      "    \n" +
      "    // 0 = Sunday, 6 = Saturday\n" +
      "    if (dayOfWeek !== 0 && dayOfWeek !== 6) {\n" +
      "      businessDays++;\n" +
      "    }\n" +
      "    \n" +
      "    // Move to next day\n" +
      "    currentDate.setDate(currentDate.getDate() + 1);\n" +
      "  }\n" +
      "  \n" +
      "  return businessDays;\n" +
      "}"
    ).setFontFamily("Courier New");
  }