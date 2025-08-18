// ################################## Web App page ##################################

function doGet() {
    let html = HtmlService
        .createTemplateFromFile("FilePicker")
        .evaluate()
        .setTitle("Google Drive")
    return html;

};

function include(filename) {
    return HtmlService.createHtmlOutputFromFile(filename)
        .getContent();
}

// #### Server - client connectors. ####


function pickerConfig() {
    DriveApp.getRootFolder()
    return {
        oauthToken: ScriptApp.getOAuthToken(),
        developerKey: PropertiesService.getScriptProperties().getProperty("developerKey")
    }
};


function storeDriveSelections(fileId) {
    // Append current list of files and folders.
    let storedDocs = JSON.parse(PropertiesService.getUserProperties()
        .getProperty("files"));

    let updateArray = () => {
        //Combine current list with incoming and remove duplicates.
        return [...new Map([...fileId, ...storedDocs].map(item => [item.id, item])).values()]

    };

    // IF not stored ids just input the fileId otherwise add both to array.
    let docsAll = (storedDocs === null) ? fileId : updateArray();


    //Add storedDocs to selected docs;
    PropertiesService.getUserProperties()
        .setProperty("files", JSON.stringify(docsAll))

    // Allows us to only keep these properties when using is working on saved properties.
    PropertiesService.getUserProperties()
        .setProperty("filePick", JSON.stringify(true));

    PropertiesService.getUserProperties().setProperty("fileId", JSON.stringify(fileId[0].id))
    PropertiesService.getUserProperties().setProperty("fileUrl", JSON.stringify(fileId[0].url))
    PropertiesService.getUserProperties().setProperty("fileName", JSON.stringify(fileId[0].name))
};

function getSheetNames() {
    const id = JSON.parse(PropertiesService.getUserProperties().getProperty('fileId'))
    const ss = SpreadsheetApp.openById(id);
    const sheets = ss.getSheets();
    const sheetNames = sheets.map(sheet => sheet.getName());
    //console.log(sheetNames)
    return sheetNames;
}


function getSheetData(sheetName) {
    const spreadsheetId = JSON.parse(PropertiesService.getUserProperties().getProperty('fileId'))

    const sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);

    const firstRowValues = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    //console.log("getSheetData: ", JSON.stringify(firstRowValues));
    return firstRowValues;


}

function getColumnValues(headerName, sheetName) {
    // Add caching for repeated calls
    const cacheKey = `${sheetName}_${headerName}_values`;
    const cache = CacheService.getUserCache();
    const cachedValues = cache.get(cacheKey);

    if (cachedValues) {
        return JSON.parse(cachedValues);
    }

    const spreadsheetId = JSON.parse(PropertiesService.getUserProperties().getProperty('fileId'));
    const sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);

    if (!sheet) {
        throw new Error(`Sheet "${sheetName}" not found in spreadsheet.`);
    }

    // Get all data at once instead of reading headers separately
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    const allData = sheet.getRange(1, 1, lastRow, lastCol).getValues();

    // Process headers in first row
    const headers = allData[0].map((header, index) => {
        if (typeof header === 'number' && header % 1 === 0) {
            return Math.round(header).toString();
        }
        return header;
    });

    const columnIndex = headers.indexOf(headerName);
    if (columnIndex === -1) {
        throw new Error(`Header "${headerName}" not found in spreadsheet headers!`);
    }

    // Extract column values efficiently
    const values = allData.slice(1) // Skip header row
        .map(row => row[columnIndex])
        .filter(value => value !== "");

    // Cache the results for 6 minutes
    cache.put(cacheKey, JSON.stringify(values), 360);

    return values;
}


function clearFilesFromPropServ() {

    PropertiesService.getUserProperties()
        .deleteProperty("files");

    PropertiesService.getUserProperties()
        .deleteProperty("fileId");
};