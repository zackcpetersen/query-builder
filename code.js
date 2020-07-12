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

// Query Builders
// keeping this in case I want to go back to walk along
// function createQuery(sheetName, myRange, select, column, compare, value) {
//   var query = 'SELECT ' + select + ' WHERE ' + column + ' ' + compare + ' ' + value;
//   var currentSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
//   var funct = "=QUERY(" + currentSheet + '!' + myRange + ', "' + query + '"' + ", 1)"
//   var newSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet().setName(sheetName);
//   newSheet.getRange(1, 1).setValue(funct);
// }

function buildQuery(formInputs) {
    if (formInputs.cleanedQuery) {
        freeFormQuery(formInputs)
    } else {
        walkthroughQuery(formInputs)
    }
}

function buildWalkthroughQuery(formInputs) {
    // take all 'query' inputs and append them to the original (I don't think order matters?)
}

function walkthroughQuery(formInputs) {
    // let query = buildWalkthroughQuery(formInputs)
    const query = 'SELECT ' + formInputs.selectInput + ' WHERE ' + formInputs.column + ' ' + formInputs.compareInput + ' ' + formInputs.value;
    const queryFunction = "=QUERY(" + formInputs.myRange + ', "' + query + '"' + ", 1)"
    ui.alert(queryFunction)
    // const newSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet().setName(sheetName);
    // newSheet.getRange(1, 1).setValue(funct);
}

function freeFormQuery(formData) {
    let queryFunction;
    if (formData.useActiveSheet) {
        queryFunction = "=QUERY(" + formData.myRange + ', "' + formData.cleanedQuery + '"' + ", 1)"
    } else {
        let dataRange = 'IMPORTRANGE(' + '"' + formData.dataSheet + '", ' + '"' + formData.myRange + '")';
        queryFunction = "=QUERY(" + dataRange + ', "' + formData.cleanedQuery + '"' + ", 1)"
    }

    ui.alert(queryFunction)

    // var newSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet();
    // newSheet.getRange(1, 1).setValue(queryFunction);
}

// Cache Update, Get, and Clear
function updateCache (data) {
    ui.alert(Object.entries(data))
    cache.putAll(data, 999);
}

function testCache() {
    var data = getFromCache();
    ui.alert(Object.entries(data))
}

function getFromCache() {
    var data = cache.getAll(inputs);
    return data
}

function clearCache() {
    cache.removeAll(inputs);
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
