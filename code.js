const ui = SpreadsheetApp.getUi();
const cache = CacheService.getUserCache();
const inputs = ['dataSheet', 'range', 'cleanedQuery', 'sheetName']

function onOpen() {
    ui.createMenu('Queries')
        .addItem('QueryBuilder', 'showSidebar')
        .addToUi();
}

function addQuery(formData, filters) {
    let myQuery = new Query(formData, filters)
    myQuery.addQueryToSheet();
}

function showSidebar() {
    const html = HtmlService.createTemplateFromFile('sidebar');
    ui.showSidebar(html.evaluate().setTitle('QueryBuilder'));
}

function picker() {
    const html = HtmlService.createTemplateFromFile('picker').evaluate();
    ui.showModalDialog(html, 'Pick Your SPREADSHEET');
}

function helpBox() {
    const html = HtmlService.createTemplateFromFile('helpbox').evaluate();
    ui.showModelessDialog(html, "How to use QueryBuilder")
}

function updateCache(data) {
    cache.putAll(data);
}

function getCache() {
    return cache.getAll(inputs);
}

// Get data range from sheet
function getActiveRange() {
    const sheet = SpreadsheetApp.getActiveSheet()
    const activeSheet = sheet.getName();
    const range = sheet.getDataRange().getA1Notation();
    return activeSheet + "!" + range
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
        this.dataSheet = formData.dataSheet;
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
            return this.walkthroughQuery()
        }
        return this.freeFormQuery()
    }
    walkthroughQuery() {
        let select = 'SELECT ' + this.selectInput + ' ';
        let query = select + this.filters;
        if (this.useActiveSheet === 'true') {
            return "=QUERY(" + this.range + ', "' + query + '"' + ", 1)"
        }
        let dataRange = 'IMPORTRANGE(' + '"' + this.dataSheet + '", ' + '"' + this.range + '")';
        return "=QUERY(" + dataRange + ', "' + query + '"' + ", 1)"
    }
    freeFormQuery() {
        if (this.useActiveSheet === 'true') {
            return "=QUERY(" + this.range + ', "' + this.cleanedQuery + '"' + ", 1)"
        }
        let dataRange = 'IMPORTRANGE(' + '"' + this.dataSheet + '", ' + '"' + this.range + '")';
        return "=QUERY(" + dataRange + ', "' + this.cleanedQuery + '"' + ", 1)"
    }
    addQueryToSheet() {
        const query = this.buildQuery()
        let newSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet();
        newSheet.getRange(1, 1).setValue(query);
    }
}
