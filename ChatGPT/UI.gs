// UI.gs
// SheetsLab - UI elements like sidebars and modals

function showNavigationSidebar() {
    var html = HtmlService.createTemplateFromFile('NavigationSidebar')
        .evaluate()
        .setTitle('Navigation');
    SpreadsheetApp.getUi().showSidebar(html);
  }
  
  function openAdvancedFormDialog() {
    var html = HtmlService.createTemplateFromFile('AdvancedFormDialog')
        .evaluate()
        .setWidth(400)
        .setHeight(300);
    SpreadsheetApp.getUi().showModalDialog(html, 'Advanced Form');
  }
  
  function openProgressDemo() {
    var html = HtmlService.createTemplateFromFile('ProgressDemo')
        .evaluate()
        .setWidth(300)
        .setHeight(200);
    SpreadsheetApp.getUi().showModalDialog(html, 'Progress Demo');
  }
  
  function openApiConnectionDialog() {
    var html = HtmlService.createTemplateFromFile('ApiConnectionDialog')
        .evaluate()
        .setWidth(400)
        .setHeight(300);
    SpreadsheetApp.getUi().showModalDialog(html, 'API Connection');
  }
  
  function openDataImportDialog() {
    var html = HtmlService.createTemplateFromFile('DataImportDialog')
        .evaluate()
        .setWidth(400)
        .setHeight(300);
    SpreadsheetApp.getUi().showModalDialog(html, 'Data Import');
  }
  
  function openEmailAutomationDialog() {
    var html = HtmlService.createTemplateFromFile('EmailAutomationDialog')
        .evaluate()
        .setWidth(400)
        .setHeight(300);
    SpreadsheetApp.getUi().showModalDialog(html, 'Email Automation');
  }
  