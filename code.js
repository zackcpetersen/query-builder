// Global variables
const ui = SpreadsheetApp.getUi();
const cache = CacheService.getUserCache();
const inputs = ['dataSheet', 'range', 'cleanedQuery']

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

// Get data range from sheet
function getActiveRange() {
    const sheet = SpreadsheetApp.getActiveSheet()
    const activeSheet = sheet.getName();
    const range = sheet.getDataRange().getA1Notation();
    return activeSheet + "!" + range
}

// Cache Update, Get, and Clear
function updateCache (data) {
    cache.putAll(data);
}

function getFromCache() {
    return cache.getAll(inputs);
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

// OAuth Token for verification and authentication
function getOAuthToken() {
    DriveApp.getRootFolder();
    return ScriptApp.getOAuthToken();
}


class Query {
    constructor(formData, filters) {
        this.data = formData.data;
        this.range = formData.range;
        this.longQuery = formData.longQuery;
        this.filters = filters;
        this.help = formData.help;
        this.useActiveSheet = formData.useActiveSheet;
        this.selectInput = formData.selectInput;
        this.cleanedQuery = formData.cleanedQuery;
    }
    buildQuery() {
        if (this.help) {
            this.query = this.walkthroughQuery();
        } else {
            this.query = this.freeFormQuery();
        }
        return this.query;
    }
    walkthroughQuery() {
        let select = 'SELECT ' + this.selectInput + ' ';
        let query = select + this.filters;
        return "=QUERY(" + this.range + ', "' + query + '"' + ", 1)"
    }
    freeFormQuery() {
        if (this.useActiveSheet) {
            return "=QUERY(" + this.range + ', "' + this.cleanedQuery + '"' + ", 1)"
        }
        let dataRange = 'IMPORTRANGE(' + '"' + this.data + '", ' + '"' + this.range + '")';
        return "=QUERY(" + dataRange + ', "' + this.cleanedQuery + '"' + ", 1)"
    }
    addQueryToSheet() {
        let newSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet();
        newSheet.getRange(1, 1).setValue(this.query);
    }
}

function addQuery(formData, filters) {
    let myQuery = new Query(formData, filters)
    myQuery.buildQuery();
    myQuery.addQueryToSheet();
}
