// Global variables
const ui = SpreadsheetApp.getUi();
const cache = CacheService.getUserCache();
const inputs = ['dataSheet', 'currentSheet', 'myRange', 'cleanedQuery', 'sheetName', 'useActiveSheet']

function onOpen() {
    ui.createMenu('Queries')
        .addItem('QueryMe', 'showSidebar')
        .addItem('Clear Cache', 'clearCache')
        .addItem('Test Cache', 'testCache')
        .addToUi();
}

// HTML Builders
function showSidebar() {
    var html = HtmlService.createTemplateFromFile('sidebar');
    ui.showSidebar(html.evaluate().setTitle('QueryMe'));
}

function picker() {
    var html = HtmlService.createTemplateFromFile('picker').evaluate();
    ui.showModalDialog(html, 'Pick Your Sheet');
}

// Get data ranges from sheet
function getActiveRange() {
    const sheet = SpreadsheetApp.getActiveSheet()
    const activeSheet = sheet.getName();
    const range = sheet.getDataRange().getA1Notation();
    return activeSheet + "!" + range
}

function buildQuery(formInputs, query) {
    if (formInputs.help) {
        walkthroughQuery(formInputs, query)
    } else {
        freeFormQuery(formInputs)
    }
}

function walkthroughQuery(formInputs, query) {
    let select = 'SELECT ' + formInputs.selectInput + ' ';
    query = select + query;
    const queryFunction = "=QUERY(" + formInputs.myRange + ', "' + query + '"' + ", 1)"
    const newSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet();
    newSheet.getRange(1, 1).setValue(queryFunction);
}

function freeFormQuery(formData) {
    let queryFunction;
    if (formData.useActiveSheet) {
        queryFunction = "=QUERY(" + formData.myRange + ', "' + formData.cleanedQuery + '"' + ", 1)"
    } else {
        let dataRange = 'IMPORTRANGE(' + '"' + formData.dataSheet + '", ' + '"' + formData.myRange + '")';
        queryFunction = "=QUERY(" + dataRange + ', "' + formData.cleanedQuery + '"' + ", 1)"
    }

    var newSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet();
    newSheet.getRange(1, 1).setValue(queryFunction);
}

// Cache Update, Get, and Clear
function updateCache (data) {
    cache.putAll(data);
}

function getFromCache() {
    var data = cache.getAll(inputs);
    return data
}


// OAuth Token for verification and authentication
function getOAuthToken() {
    DriveApp.getRootFolder();
    return ScriptApp.getOAuthToken();
}

// Include function allowing files to access other files
function include(filename) {
    return HtmlService.createHtmlOutputFromFile(filename)
        .getContent();
}

// For error messages
function displayError(error) {
    ui.alert(error)
}
