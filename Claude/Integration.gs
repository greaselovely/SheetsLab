/**
 * Integration.gs
 * External API connections and service integrations for SheetsLab
 * 
 * This file contains functions to demonstrate integration capabilities
 * with external APIs and other Google services.
 * 
 * @version 1.0.0
 */

/**
 * Shows a dialog for connecting to external APIs
 */
function showApiConnectionDialog() {
  // Create the HTML for the dialog
  const htmlOutput = HtmlService.createTemplateFromFile('ApiConnectionDialog')
    .evaluate()
    .setWidth(CONFIG.UI.DIALOG_WIDTH)
    .setHeight(CONFIG.UI.DIALOG_HEIGHT);
  
  // Display the dialog
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'External API Connection');
}

/**
 * Fetches data from a public API and imports it into the sheet
 * @param {Object} options - Parameters for the API request
 * @return {Object} Result object with success status and message
 */
function fetchApiData(options) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.INTEGRATION_LAB);
    
    if (!sheet) {
      return {
        success: false,
        message: 'Integration Lab sheet not found. Please initialize SheetsLab first.'
      };
    }
    
    // Determine which API to use based on the options
    let apiUrl;
    let apiTitle;
    
    switch (options.api) {
      case 'random-user':
        apiUrl = 'https://randomuser.me/api/?results=' + options.count;
        apiTitle = 'Random User Generator API';
        break;
      case 'countries':
        apiUrl = 'https://restcountries.com/v3.1/all';
        apiTitle = 'Countries REST API';
        break;
      case 'open-library':
        apiUrl = 'https://openlibrary.org/search.json?q=' + encodeURIComponent(options.query) + '&limit=' + options.count;
        apiTitle = 'Open Library API';
        break;
      case 'exchange-rates':
        apiUrl = 'https://open.er-api.com/v6/latest/USD';
        apiTitle = 'Exchange Rates API';
        break;
      case 'iss-location':
        apiUrl = 'http://api.open-notify.org/iss-now.json';
        apiTitle = 'ISS Current Location API';
        break;
      case 'iss-people':
        apiUrl = 'http://api.open-notify.org/astros.json';
        apiTitle = 'ISS People in Space API';
        break;
      default:
        return {
          success: false,
          message: 'Invalid API selected.'
        };
    }
    
    // Make the API request
    const response = UrlFetchApp.fetch(apiUrl);
    const json = JSON.parse(response.getContentText());
    
    // Clear previous data
    sheet.getRange("A6:Z100").clear();
    
    // Set a title for the imported data
    sheet.getRange("A6:D6").merge().setValue('Data from ' + apiTitle)
      .setFontWeight("bold")
      .setBackground("#E6E6E6");
    
    // Process the data based on the API type
    switch (options.api) {
      case 'random-user':
        processRandomUserData(sheet, json);
        break;
      case 'countries':
        processCountriesData(sheet, json);
        break;
      case 'open-library':
        processOpenLibraryData(sheet, json);
        break;
      case 'exchange-rates':
        processExchangeRatesData(sheet, json);
        break;
      case 'iss-location':
        processIssLocationData(sheet, json);
        break;
      case 'iss-people':
        processIssPeopleData(sheet, json);
        break;
    }
    
    // Format and add borders
    sheet.getRange("A7:Z50").applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
    
    return {
      success: true,
      message: 'Data successfully imported from ' + apiTitle
    };
  } catch (error) {
    console.error('Error fetching API data:', error);
    return {
      success: false,
      message: 'Error fetching data: ' + error.toString()
    };
  }
}

/**
 * Processes data from the Random User Generator API
 * @param {SpreadsheetApp.Sheet} sheet - The sheet to write data to
 * @param {Object} json - The parsed JSON response
 */
function processRandomUserData(sheet, json) {
  // Set headers
  const headers = ['Name', 'Email', 'Phone', 'Gender', 'Age', 'City', 'Country', 'Picture'];
  sheet.getRange(7, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
  
  // Process each user
  const users = json.results;
  const data = [];
  
  for (const user of users) {
    data.push([
      `${user.name.first} ${user.name.last}`,
      user.email,
      user.phone,
      user.gender,
      user.dob.age,
      user.location.city,
      user.location.country,
      `=IMAGE("${user.picture.thumbnail}", 1)`
    ]);
  }
  
  // Write the data to the sheet
  sheet.getRange(8, 1, data.length, data[0].length).setValues(data);
  
  // Auto-resize columns
  for (let i = 1; i <= headers.length; i++) {
    sheet.autoResizeColumn(i);
  }
  
  // Add hyperlinks to emails
  const emailRange = sheet.getRange(8, 2, data.length, 1);
  const emailValues = emailRange.getValues();
  
  for (let i = 0; i < emailValues.length; i++) {
    const email = emailValues[i][0];
    const formula = `=HYPERLINK("mailto:${email}", "${email}")`;
    sheet.getRange(8 + i, 2).setFormula(formula);
  }
}

/**
 * Processes data from the Countries REST API
 * @param {SpreadsheetApp.Sheet} sheet - The sheet to write data to
 * @param {Object} json - The parsed JSON response
 */
function processCountriesData(sheet, json) {
  // Set headers
  const headers = ['Name', 'Capital', 'Region', 'Subregion', 'Population', 'Area (km²)', 'Languages', 'Currencies', 'Flag'];
  sheet.getRange(7, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
  
  // Sort countries by name
  json.sort((a, b) => a.name.common.localeCompare(b.name.common));
  
  // Take only the first 100 countries to avoid overwhelming the sheet
  const countries = json.slice(0, 100);
  const data = [];
  
  for (const country of countries) {
    // Handle potential missing data
    const capital = country.capital && country.capital.length > 0 ? country.capital[0] : 'N/A';
    
    // Get languages as a comma-separated string
    let languages = 'N/A';
    if (country.languages) {
      languages = Object.values(country.languages).join(', ');
    }
    
    // Get currencies as a comma-separated string
    let currencies = 'N/A';
    if (country.currencies) {
      currencies = Object.values(country.currencies)
        .map(c => `${c.name} (${c.symbol || 'N/A'})`)
        .join(', ');
    }
    
    data.push([
      country.name.common,
      capital,
      country.region || 'N/A',
      country.subregion || 'N/A',
      country.population || 0,
      country.area || 0,
      languages,
      currencies,
      `=IMAGE("${country.flags.png}", 1)`
    ]);
  }
  
  // Write the data to the sheet
  sheet.getRange(8, 1, data.length, data[0].length).setValues(data);
  
  // Format number columns
  sheet.getRange(8, 5, data.length, 1).setNumberFormat("#,##0");
  sheet.getRange(8, 6, data.length, 1).setNumberFormat("#,##0.00");
  
  // Auto-resize columns
  for (let i = 1; i <= headers.length - 1; i++) { // Skip the flag column
    sheet.autoResizeColumn(i);
  }
  
  // Set flag column width
  sheet.setColumnWidth(9, 120);
}

/**
 * Processes data from the Open Library API
 * @param {SpreadsheetApp.Sheet} sheet - The sheet to write data to
 * @param {Object} json - The parsed JSON response
 */
function processOpenLibraryData(sheet, json) {
  // Set headers
  const headers = ['Title', 'Author', 'Year', 'Publisher', 'ISBN', 'Language', 'Subject'];
  sheet.getRange(7, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
  
  // Process each book
  const books = json.docs;
  const data = [];
  
  for (const book of books) {
    // Extract author name(s)
    let authorNames = 'Unknown';
    if (book.author_name && book.author_name.length > 0) {
      authorNames = book.author_name.join(', ');
    }
    
    // Extract year
    const year = book.first_publish_year || 'Unknown';
    
    // Extract publisher
    let publisher = 'Unknown';
    if (book.publisher && book.publisher.length > 0) {
      publisher = book.publisher[0];
    }
    
    // Extract ISBN
    let isbn = 'N/A';
    if (book.isbn && book.isbn.length > 0) {
      isbn = book.isbn[0];
    }
    
    // Extract language
    let language = 'Unknown';
    if (book.language && book.language.length > 0) {
      language = book.language.join(', ');
    }
    
    // Extract subject
    let subject = 'N/A';
    if (book.subject && book.subject.length > 0) {
      subject = book.subject.slice(0, 3).join(', ');
    }
    
    data.push([
      book.title,
      authorNames,
      year,
      publisher,
      isbn,
      language,
      subject
    ]);
  }
  
  // Write the data to the sheet
  sheet.getRange(8, 1, data.length, data[0].length).setValues(data);
  
  // Auto-resize columns
  for (let i = 1; i <= headers.length; i++) {
    sheet.autoResizeColumn(i);
  }
  
  // Make title column wider
  sheet.setColumnWidth(1, 300);
  
  // Add a link to Open Library for each book
  sheet.getRange(7, headers.length + 1).setValue('Link').setFontWeight("bold");
  
  for (let i = 0; i < books.length; i++) {
    const key = books[i].key;
    if (key) {
      const formula = `=HYPERLINK("https://openlibrary.org${key}", "View on Open Library")`;
      sheet.getRange(8 + i, headers.length + 1).setFormula(formula);
    } else {
      sheet.getRange(8 + i, headers.length + 1).setValue('N/A');
    }
  }
}

/**
 * Processes data from the Exchange Rates API
 * @param {SpreadsheetApp.Sheet} sheet - The sheet to write data to
 * @param {Object} json - The parsed JSON response
 */
function processExchangeRatesData(sheet, json) {
  // Set title and information
  sheet.getRange("A7:C7").merge().setValue('Base Currency: USD (US Dollar)')
    .setFontWeight("bold");
  sheet.getRange("A8:C8").merge().setValue('Last Update: ' + json.time_last_update_utc);
  
  // Set headers
  const headers = ['Currency Code', 'Currency Name', 'Rate (vs USD)', 'USD Value'];
  sheet.getRange(10, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
  
  // Process each currency
  const rates = json.rates;
  const data = [];
  
  // Define common currency names (since the API doesn't provide them)
  const currencyNames = {
    USD: 'US Dollar',
    EUR: 'Euro',
    GBP: 'British Pound',
    JPY: 'Japanese Yen',
    CAD: 'Canadian Dollar',
    AUD: 'Australian Dollar',
    CHF: 'Swiss Franc',
    CNY: 'Chinese Yuan',
    INR: 'Indian Rupee',
    BRL: 'Brazilian Real',
    RUB: 'Russian Ruble',
    KRW: 'South Korean Won',
    SGD: 'Singapore Dollar',
    NZD: 'New Zealand Dollar',
    MXN: 'Mexican Peso',
    HKD: 'Hong Kong Dollar',
    TRY: 'Turkish Lira',
    ZAR: 'South African Rand',
    SEK: 'Swedish Krona',
    NOK: 'Norwegian Krone'
  };
  
  // Convert object to array for sorting
  const ratesArray = Object.entries(rates);
  ratesArray.sort((a, b) => a[0].localeCompare(b[0]));
  
  for (const [code, rate] of ratesArray) {
    data.push([
      code,
      currencyNames[code] || code + ' Currency',
      rate,
      100 / rate // Value of 100 USD in this currency
    ]);
  }
  
  // Write the data to the sheet
  sheet.getRange(11, 1, data.length, data[0].length).setValues(data);
  
  // Format number columns
  sheet.getRange(11, 3, data.length, 1).setNumberFormat("0.00000");
  sheet.getRange(11, 4, data.length, 1).setNumberFormat("0.00");
  
  // Auto-resize columns
  for (let i = 1; i <= headers.length; i++) {
    sheet.autoResizeColumn(i);
  }
  
  // Create a conversion calculator
  sheet.getRange("F10:I10").merge().setValue('Currency Converter')
    .setFontWeight("bold")
    .setBackground("#E6E6E6")
    .setHorizontalAlignment("center");
  
  sheet.getRange("F11").setValue('Amount:');
  sheet.getRange("G11").setValue(100);
  
  sheet.getRange("F12").setValue('From:');
  sheet.getRange("G12").setValue('USD');
  
  sheet.getRange("F13").setValue('To:');
  sheet.getRange("G13").setValue('EUR');
  
  sheet.getRange("F14").setValue('Result:');
  
  // Create a conversion formula
  const formula = '=G11 * VLOOKUP(G13, A11:C' + (10 + data.length) + ', 3, FALSE) / VLOOKUP(G12, A11:C' + (10 + data.length) + ', 3, FALSE)';
  sheet.getRange("G14").setFormula(formula).setNumberFormat("0.00");
  
  // Create data validation for currency selection
  const currencyCodes = data.map(row => row[0]);
  const currencyRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(currencyCodes, true)
    .build();
  
  sheet.getRange("G12").setDataValidation(currencyRule);
  sheet.getRange("G13").setDataValidation(currencyRule);
  
  // Add a small chart
  const chartRange = sheet.getRange("A11:B" + Math.min(21, 10 + data.length));
  
  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(chartRange)
    .setPosition(20, 6, 0, 0)
    .setOption('title', 'Major Currencies')
    .setOption('legend', {position: 'none'})
    .setOption('height', 300)
    .setOption('width', 400)
    .build();
  
  sheet.insertChart(chart);
}

/**
 * Processes data from the ISS Current Location API
 * @param {SpreadsheetApp.Sheet} sheet - The sheet to write data to
 * @param {Object} json - The parsed JSON response
 */
function processIssLocationData(sheet, json) {
  // Extract data from the API response
  const timestamp = json.timestamp;
  const datetime = new Date(timestamp * 1000); // Convert Unix timestamp to JavaScript Date
  const latitude = parseFloat(json.iss_position.latitude);
  const longitude = parseFloat(json.iss_position.longitude);
  
  // Add headers and current position
  sheet.getRange("A7:B7").setValues([["Current ISS Location (Updated at)", datetime]]);
  sheet.getRange("B7").setNumberFormat("yyyy-MM-dd HH:mm:ss");
  
  sheet.getRange("A8:B8").setValues([["Latitude", latitude]]);
  sheet.getRange("A9:B9").setValues([["Longitude", longitude]]);
  
  // Add map visualization
  sheet.getRange("A11:D11").merge().setValue("ISS Position Visualization")
    .setFontWeight("bold");
  
  // Create a static map image using Google Maps Static API (doesn't require an API key for basic usage)
  const mapUrl = `https://maps.googleapis.com/maps/api/staticmap?center=${latitude},${longitude}&zoom=2&size=600x300&maptype=terrain&markers=color:red%7C${latitude},${longitude}`;
  
  // Add the map image
  sheet.getRange("A12").setFormula(`=IMAGE("${mapUrl}", 4, 600, 300)`);
  
  // Add a coordinate plane visualization
  sheet.getRange("A14:D14").merge().setValue("ISS Position on Earth Coordinate Plane")
    .setFontWeight("bold");
  
  // Create a coordinate grid
  createCoordinateGrid(sheet, 15, 1, latitude, longitude);
  
  // Add information about the ISS
  sheet.getRange("A28:D28").merge().setValue("About the International Space Station")
    .setFontWeight("bold");
  
  const issInfo = [
    ["Altitude", "~408 km (253 mi)"],
    ["Orbital Speed", "~28,000 km/h (17,500 mph)"],
    ["Orbital Period", "~92 minutes (16 orbits per day)"],
    ["Launch Date", "November 20, 1998"],
    ["Mass", "~420,000 kg (925,000 lb)"],
    ["Length", "109 m (358 ft)"],
    ["Width", "73 m (240 ft)"],
    ["Pressurized Volume", "915 m³ (32,300 ft³)"],
    ["Data Source", "Open Notify API (http://api.open-notify.org/)"]
  ];
  
  sheet.getRange(29, 1, issInfo.length, 2).setValues(issInfo);
  
  // Add a refresh button for real-time updates
  sheet.getRange("D8:E8").merge().setValue("Click to Refresh ISS Location")
    .setBackground("#4285F4")
    .setFontColor("white")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
}

/**
 * Creates a simple coordinate grid to visualize the ISS position
 * @param {SpreadsheetApp.Sheet} sheet - The sheet to add the grid to
 * @param {number} startRow - The starting row for the grid
 * @param {number} startCol - The starting column for the grid
 * @param {number} latitude - The ISS latitude
 * @param {number} longitude - The ISS longitude
 */
function createCoordinateGrid(sheet, startRow, startCol, latitude, longitude) {
  // Create a 10x20 grid representing the Earth's coordinates
  const rows = 10; // From 90°N to 90°S in 18° increments
  const cols = 20; // From 180°W to 180°E in 18° increments
  
  // Create the grid data
  const gridData = [];
  for (let i = 0; i < rows; i++) {
    const row = [];
    for (let j = 0; j < cols; j++) {
      // Leave empty for now
      row.push("");
    }
    gridData.push(row);
  }
  
  // Write the grid to the sheet
  sheet.getRange(startRow, startCol, rows, cols).setValues(gridData);
  
  // Add borders to create a grid
  sheet.getRange(startRow, startCol, rows, cols).setBorder(true, true, true, true, true, true);
  
  // Set cell sizes to make a more square grid
  for (let j = startCol; j < startCol + cols; j++) {
    sheet.setColumnWidth(j, 30);
  }
  for (let i = startRow; i < startRow + rows; i++) {
    sheet.setRowHeight(i, 30);
  }
  
  // Calculate which cell the ISS is in
  const latIndex = Math.floor((90 - latitude) / 180 * rows);
  const lonIndex = Math.floor((longitude + 180) / 360 * cols);
  
  // Ensure indices are within bounds
  const boundedLatIndex = Math.min(Math.max(latIndex, 0), rows - 1);
  const boundedLonIndex = Math.min(Math.max(lonIndex, 0), cols - 1);
  
  // Mark the ISS position with a red background
  sheet.getRange(startRow + boundedLatIndex, startCol + boundedLonIndex)
    .setBackground("#F4C7C3")
    .setValue("ISS");
  
  // Add coordinate labels
  sheet.getRange(startRow - 1, startCol + Math.floor(cols / 2)).setValue("North Pole")
    .setHorizontalAlignment("center");
  sheet.getRange(startRow + rows, startCol + Math.floor(cols / 2)).setValue("South Pole")
    .setHorizontalAlignment("center");
  sheet.getRange(startRow + Math.floor(rows / 2), startCol - 1).setValue("West")
    .setHorizontalAlignment("right");
  sheet.getRange(startRow + Math.floor(rows / 2), startCol + cols).setValue("East")
    .setHorizontalAlignment("left");
}

/**
 * Processes data from the ISS People in Space API
 * @param {SpreadsheetApp.Sheet} sheet - The sheet to write data to
 * @param {Object} json - The parsed JSON response
 */
function processIssPeopleData(sheet, json) {
  // Extract data from the API response
  const numberOfPeople = json.number;
  const people = json.people;
  
  // Display the number of people in space
  sheet.getRange("A7:B7").setValues([["Number of People Currently in Space:", numberOfPeople]])
    .setFontWeight("bold");
  
  // Add table headers
  sheet.getRange("A9:C9").setValues([["Name", "Craft", "Days in Space (Est.)"]]).setFontWeight("bold");
  
  // Process data for each astronaut
  const astronautData = [];
  
  for (const person of people) {
    // For demo purposes, generate a random number of days in space
    // In a real application, you would get this data from another source
    const daysInSpace = Math.floor(Math.random() * 180) + 1;
    
    astronautData.push([
      person.name,
      person.craft,
      daysInSpace
    ]);
  }
  
  // Write data to the sheet
  sheet.getRange(10, 1, astronautData.length, 3).setValues(astronautData);
  
  // Add some formatting
  sheet.getRange(10, 3, astronautData.length, 1).setNumberFormat("0");
  
  // Auto-resize columns
  for (let i = 1; i <= 3; i++) {
    sheet.autoResizeColumn(i);
  }
  
  // Add a simple bar chart showing days in space
  const chartRange = sheet.getRange(10, 1, astronautData.length, 3);
  
  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(chartRange)
    .setPosition(10 + astronautData.length + 2, 1, 0, 0)
    .setOption('title', 'Days in Space by Astronaut')
    .setOption('legend', {position: 'none'})
    .setOption('hAxis', {title: 'Days'})
    .setOption('vAxis', {title: 'Astronaut'})
    .setOption('height', 300)
    .setOption('width', 500)
    .build();
  
  sheet.insertChart(chart);
  
  // Add information about astronauts in space
  sheet.getRange("A25:D25").merge().setValue("About People in Space")
    .setFontWeight("bold");
  
  const info = [
    ["The International Space Station (ISS) is typically home to 6-7 astronauts at any given time."],
    ["Astronauts usually stay on the ISS for about 6 months, though some missions last longer."],
    ["Expedition crews are made up of astronauts from various space agencies including NASA, Roscosmos, ESA, JAXA, and CSA."],
    ["The ISS has been continuously occupied since November 2, 2000."],
    ["Data Source: Open Notify API (http://api.open-notify.org/)"]
  ];
  
  sheet.getRange(26, 1, info.length, 1).setValues(info);
}

/**
 * Shows a dialog with data import options
 */
function showDataImportOptions() {
  // Create the HTML for the dialog
  const htmlOutput = HtmlService.createTemplateFromFile('DataImportDialog')
    .evaluate()
    .setWidth(CONFIG.UI.DIALOG_WIDTH)
    .setHeight(CONFIG.UI.DIALOG_HEIGHT);
  
  // Display the dialog
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Import External Data');
}

/**
 * Imports sample calendar events to the sheet
 * @return {Object} Result object with success status and message
 */
function importCalendarEvents() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.INTEGRATION_LAB);
    
    if (!sheet) {
      return {
        success: false,
        message: 'Integration Lab sheet not found. Please initialize SheetsLab first.'
      };
    }
    
    // Clear previous data
    sheet.getRange("A6:Z100").clear();
    
    // Set a title for the imported data
    sheet.getRange("A6:D6").merge().setValue('Calendar Integration Example')
      .setFontWeight("bold")
      .setBackground("#E6E6E6");
    
    // Set headers
    const headers = ['Event Title', 'Start Time', 'End Time', 'Location', 'Description'];
    sheet.getRange(7, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
    
    // Get calendar events for the next 30 days
    const now = new Date();
    const thirtyDaysLater = new Date(now.getTime() + 30 * 24 * 60 * 60 * 1000);
    
    const calendar = CalendarApp.getDefaultCalendar();
    const events = calendar.getEvents(now, thirtyDaysLater);
    
    if (events.length === 0) {
      // No events found, so create some sample events
      sheet.getRange("A8:E8").merge().setValue('No calendar events found. Using sample data instead.')
        .setFontStyle("italic");
      
      // Create sample data
      const sampleData = [
        ['Team Meeting', new Date(now.getTime() + 2 * 24 * 60 * 60 * 1000), new Date(now.getTime() + 2 * 24 * 60 * 60 * 1000 + 60 * 60 * 1000), 'Conference Room A', 'Weekly team sync-up'],
        ['Project Review', new Date(now.getTime() + 5 * 24 * 60 * 60 * 1000), new Date(now.getTime() + 5 * 24 * 60 * 60 * 1000 + 2 * 60 * 60 * 1000), 'Virtual Meeting', 'End of sprint review'],
        ['Client Call', new Date(now.getTime() + 7 * 24 * 60 * 60 * 1000), new Date(now.getTime() + 7 * 24 * 60 * 60 * 1000 + 30 * 60 * 1000), 'Phone', 'Discuss new requirements'],
        ['Team Lunch', new Date(now.getTime() + 10 * 24 * 60 * 60 * 1000), new Date(now.getTime() + 10 * 24 * 60 * 60 * 1000 + 90 * 60 * 1000), 'Downtown Cafe', 'Team building'],
        ['Training Session', new Date(now.getTime() + 14 * 24 * 60 * 60 * 1000), new Date(now.getTime() + 14 * 24 * 60 * 60 * 1000 + 4 * 60 * 60 * 1000), 'Training Room B', 'New tools workshop']
      ];
      
      sheet.getRange(9, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
      
      // Format date columns
      sheet.getRange(9, 2, sampleData.length, 2).setNumberFormat("yyyy-MM-dd hh:mm a");
      
      // Add note about sample data
      sheet.getRange("A" + (10 + sampleData.length) + ":E" + (10 + sampleData.length)).merge()
        .setValue('Note: In a real integration, this would show actual calendar events from your Google Calendar.')
        .setFontStyle("italic");
    } else {
      // Process real calendar events
      const data = [];
      
      for (const event of events) {
        data.push([
          event.getTitle(),
          event.getStartTime(),
          event.getEndTime(),
          event.getLocation() || 'N/A',
          event.getDescription() || 'N/A'
        ]);
      }
      
      // Write the data to the sheet
      sheet.getRange(8, 1, data.length, data[0].length).setValues(data);
      
      // Format date columns
      sheet.getRange(8, 2, data.length, 2).setNumberFormat("yyyy-MM-dd hh:mm a");
    }
    
    // Auto-resize columns
    for (let i = 1; i <= headers.length; i++) {
      sheet.autoResizeColumn(i);
    }
    
    return {
      success: true,
      message: 'Calendar data imported successfully!'
    };
  } catch (error) {
    console.error('Error importing calendar events:', error);
    return {
      success: false,
      message: 'Error importing calendar data: ' + error.toString()
    };
  }
}

/**
 * Imports sample Gmail messages to the sheet
 * @return {Object} Result object with success status and message
 */
function importGmailMessages() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.INTEGRATION_LAB);
    
    if (!sheet) {
      return {
        success: false,
        message: 'Integration Lab sheet not found. Please initialize SheetsLab first.'
      };
    }
    
    // Clear previous data
    sheet.getRange("A6:Z100").clear();
    
    // Set a title for the imported data
    sheet.getRange("A6:D6").merge().setValue('Gmail Integration Example')
      .setFontWeight("bold")
      .setBackground("#E6E6E6");
    
    // Set headers
    const headers = ['From', 'Subject', 'Date', 'Snippet', 'Label'];
    sheet.getRange(7, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
    
    // Try to get Gmail threads
    const threads = GmailApp.getInboxThreads(0, 10);
    
    if (threads.length === 0) {
      // No threads found, so create some sample data
      sheet.getRange("A8:E8").merge().setValue('No Gmail messages found. Using sample data instead.')
        .setFontStyle("italic");
      
      // Create sample data
      const now = new Date();
      const sampleData = [
        ['john.doe@example.com', 'Project Update: Q3 Goals', new Date(now.getTime() - 2 * 24 * 60 * 60 * 1000), 'Here are the updated goals for Q3 based on our discussion yesterday...', 'Inbox'],
        ['team-notifications@company.com', 'New Comment on Task #1234', new Date(now.getTime() - 3 * 24 * 60 * 60 * 1000), 'Sarah added a comment to task #1234: "Let\'s discuss this at the next meeting"', 'Notifications'],
        ['sales@vendor.com', 'Your recent order #56789', new Date(now.getTime() - 5 * 24 * 60 * 60 * 1000), 'Thank you for your recent order. Your items will be shipped within 2 business days...', 'Orders'],
        ['newsletter@industry-news.com', 'Weekly Industry Roundup', new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000), 'Top stories this week: New market trends, upcoming conferences, and technology innovations...', 'Newsletters'],
        ['support@service.com', 'Your support ticket resolved', new Date(now.getTime() - 10 * 24 * 60 * 60 * 1000), 'Your recent support ticket #45678 has been resolved. Please let us know if you have any further issues...', 'Support']
      ];
      
      sheet.getRange(9, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
      
      // Format date column
      sheet.getRange(9, 3, sampleData.length, 1).setNumberFormat("yyyy-MM-dd hh:mm a");
      
      // Add note about sample data
      sheet.getRange("A" + (10 + sampleData.length) + ":E" + (10 + sampleData.length)).merge()
        .setValue('Note: In a real integration, this would show actual messages from your Gmail inbox.')
        .setFontStyle("italic");
    } else {
      // Process real Gmail threads
      const data = [];
      
      for (const thread of threads) {
        const messages = thread.getMessages();
        const message = messages[0]; // Get the first message in the thread
        
        data.push([
          message.getFrom(),
          message.getSubject(),
          message.getDate(),
          message.getPlainBody().substring(0, 100) + '...',
          thread.getFirstMessageSubject().length > 0 ? 'Inbox' : 'No Label'
        ]);
      }
      
      // Write the data to the sheet
      sheet.getRange(8, 1, data.length, data[0].length).setValues(data);
      
      // Format date column
      sheet.getRange(8, 3, data.length, 1).setNumberFormat("yyyy-MM-dd hh:mm a");
    }
    
    // Auto-resize columns
    for (let i = 1; i <= headers.length; i++) {
      sheet.autoResizeColumn(i);
    }
    
    return {
      success: true,
      message: 'Gmail data imported successfully!'
    };
  } catch (error) {
    console.error('Error importing Gmail messages:', error);
    return {
      success: false,
      message: 'Error importing Gmail data: ' + error.toString()
    };
  }
}

/**
 * Shows a dialog for email automation demo
 */
function showEmailAutomationDialog() {
  // Create the HTML for the dialog
  const htmlOutput = HtmlService.createTemplateFromFile('EmailAutomationDialog')
    .evaluate()
    .setWidth(CONFIG.UI.DIALOG_WIDTH)
    .setHeight(CONFIG.UI.DIALOG_HEIGHT);
  
  // Display the dialog
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Email Automation Demo');
}

/**
 * Actually sends a test email based on form input (real implementation)
 * @param {Object} formData - Email form data
 * @return {Object} Result object with success status and message
 */
function sendTestEmail(formData) {
  try {
    // Create the email body with proper formatting
    let body = formData.body;
    
    // If including data from the sheet, add it to the email
    if (formData.includeData) {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName(CONFIG.SHEETS.INTEGRATION_LAB);
      
      if (sheet) {
        // Get some sample data to include
        const dataRange = sheet.getRange("A7:C15").getValues();
        
        if (formData.dataFormat === 'table') {
          // Create an HTML table for the data
          let tableHtml = '<br><br><table border="1" cellpadding="5" style="border-collapse:collapse;">';
          
          // Add headers
          tableHtml += '<tr style="background-color:#f3f3f3;font-weight:bold;">';
          for (let i = 0; i < 3; i++) {
            tableHtml += '<th>' + (dataRange[0][i] || '') + '</th>';
          }
          tableHtml += '</tr>';
          
          // Add data rows
          for (let i = 1; i < dataRange.length; i++) {
            tableHtml += '<tr>';
            for (let j = 0; j < 3; j++) {
              tableHtml += '<td>' + (dataRange[i][j] || '') + '</td>';
            }
            tableHtml += '</tr>';
          }
          
          tableHtml += '</table>';
          
          // Add the table to the email body
          body += '<br><br><p>Attached Data:</p>' + tableHtml;
        } else {
          // For CSV format, we'll just note that it would be attached
          body += '<br><br><p>Note: In a real implementation, a CSV file would be attached here.</p>';
        }
      }
    }
    
    // Check if email is actually being sent or simulated
    const actualSend = false; // Set to true to actually send emails (requires Gmail permission)
    
    if (actualSend) {
      // Actually send the email using GmailApp
      GmailApp.sendEmail(
        formData.to,
        formData.subject,
        // Plain text version
        body.replace(/<[^>]*>/g, ''), // Remove HTML tags for plain text
        {
          htmlBody: body, // HTML version
          name: 'SheetsLab Demo'
        }
      );
      
      return {
        success: true,
        message: 'Email successfully sent to ' + formData.to
      };
    } else {
      // Simulate email sending
      console.log('Email Data:', formData);
      
      // Simulate sending delay
      Utilities.sleep(2000);
      
      return {
        success: true,
        message: 'Email would be sent to ' + formData.to + ' (This is a simulation, no actual email was sent)'
      };
    }
  } catch (error) {
    console.error('Error sending test email:', error);
    return {
      success: false,
      message: 'Error sending email: ' + error.toString()
    };
  }
}

/**
 * Schedules recurring emails based on the configuration
 * @param {Object} config - Email configuration
 * @return {Object} Result with success status and trigger ID
 */
function scheduleRecurringEmail(config) {
  try {
    // Delete any existing triggers with the same label
    const existingTriggers = ScriptApp.getProjectTriggers();
    for (const trigger of existingTriggers) {
      if (trigger.getHandlerFunction() === 'sendScheduledEmail') {
        ScriptApp.deleteTrigger(trigger);
      }
    }
    
    // Create the appropriate trigger based on the schedule type
    let trigger;
    switch (config.schedule) {
      case 'daily':
        trigger = ScriptApp.newTrigger('sendScheduledEmail')
          .timeBased()
          .atHour(9) // 9 AM
          .everyDays(1)
          .create();
        break;
      case 'weekly':
        trigger = ScriptApp.newTrigger('sendScheduledEmail')
          .timeBased()
          .onWeekDay(ScriptApp.WeekDay.MONDAY)
          .atHour(9) // 9 AM
          .create();
        break;
      case 'monthly':
        trigger = ScriptApp.newTrigger('sendScheduledEmail')
          .timeBased()
          .onMonthDay(1) // 1st day of the month
          .atHour(9) // 9 AM
          .create();
        break;
      default:
        return {
          success: false,
          message: 'Invalid schedule type'
        };
    }
    
    // Store the configuration in Properties service
    const props = PropertiesService.getScriptProperties();
    props.setProperty('emailConfig', JSON.stringify(config));
    
    return {
      success: true,
      message: 'Email scheduled successfully!',
      triggerId: trigger.getUniqueId()
    };
  } catch (error) {
    console.error('Error scheduling email:', error);
    return {
      success: false,
      message: 'Error scheduling email: ' + error.toString()
    };
  }
}

/**
 * Handler for sending scheduled emails
 * Called by the time-based trigger
 */
function sendScheduledEmail() {
  try {
    // Get the email configuration from Properties service
    const props = PropertiesService.getScriptProperties();
    const configStr = props.getProperty('emailConfig');
    
    if (!configStr) {
      console.error('No email configuration found');
      return;
    }
    
    const config = JSON.parse(configStr);
    
    // Send the email
    sendTestEmail(config);
    
    // Log the success
    console.log('Scheduled email sent successfully at ' + new Date());
  } catch (error) {
    console.error('Error sending scheduled email:', error);
  }
}