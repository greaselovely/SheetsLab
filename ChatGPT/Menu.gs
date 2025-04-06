// Menu.gs
// SheetsLab - Custom menu creation and handling

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("SheetsLab")
    .addItem("Setup Sheets", "setupSheets")
    .addItem("Show Sidebar", "showNavigationSidebar")
    .addItem("Open Data Tools", "openDataTools")
    .addItem("View Dashboard", "openDashboard")
    .addItem("Demo ISS API", "demoApiIntegration")
    .addItem("Demo CSV Import", "demoCsvImport")
    .addSeparator()
    .addSubMenu(ui.createMenu("Advanced")
      .addItem("Formula Examples", "openFormulaExamples")
      .addItem("Integration Settings", "openIntegrationSettings"))
    .addToUi();
}

function openDataTools() {
  SpreadsheetApp.getUi().alert("Data Tools functionality coming soon!");
}

function openDashboard() {
  try {
    createDashboardChart();
    SpreadsheetApp.getUi().alert("Dashboard chart created in the DASHBOARD sheet.");
  } catch(e) {
    SpreadsheetApp.getUi().alert("Error creating dashboard chart: " + e.message);
  }
}

function openFormulaExamples() {
  SpreadsheetApp.getUi().alert("Formula Examples functionality coming soon!");
}

function openIntegrationSettings() {
  SpreadsheetApp.getUi().alert("Integration Settings functionality coming soon!");
}
