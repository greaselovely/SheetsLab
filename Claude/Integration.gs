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
      
      // Add a button to create calendar event (simulated with a shape)
      sheet.getRange("A" + (sheet.getLastRow() + 3) + ":B" + (sheet.getLastRow() + 3)).merge()
        .setValue('Create Sample Event')
        .setFontWeight("bold")
        .setBackground("#4285F4")
        .setFontColor("white")
        .setHorizontalAlignment("center");
      
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
   * Sends a test email based on form input
   * @param {Object} formData - Email form data
   * @return {Object} Result object with success status and message
   */
  function sendTestEmail(formData) {
    try {
      // In a real implementation, this would actually send emails
      // For demo purposes, we'll just log the data and return success
      
      console.log('Email Data:', formData);
      
      // Simulate email sending delay
      Utilities.sleep(2000);
      
      return {
        success: true,
        message: 'Email would be sent to ' + formData.to + ' (This is a simulation, no actual email was sent)'
      };
    } catch (error) {
      console.error('Error sending test email:', error);
      return {
        success: false,
        message: 'Error in email simulation: ' + error.toString()
      };
    }
  }