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
    //headerName = headerName.replace(/\r?\n|\r/g, " ")
    //console.log(`Getting column values for header: "${headerName}"`);
    const spreadsheetId = JSON.parse(PropertiesService.getUserProperties().getProperty('fileId'));
    //console.log(`Spreadsheet ID: ${spreadsheetId}`);
    const sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);
    //console.log(`Sheet name: ${sheetName}, exists: ${sheet !== null}`);

    if (!sheet) {
        //console.error(`Sheet "${sheetName}" not found in spreadsheet!`);
        throw new Error(`Sheet "${sheetName}" not found in spreadsheet.`);
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    // console.log(`Headers found (${headers.length})`);
    // Log each header with its index for clarity
    // Process headers to handle number conversion
    const processedHeaders = headers.map((header, index) => {
        let processedHeader = header;

        //Convert floating point numbers back to integers for matching
        if (typeof header === 'number' && header % 1 === 0) {
            // console.log("getColumnValues: ")
            processedHeader = Math.round(header).toString();
        }

        // console.log(`Header[${index+1}] original: "${header}" (${typeof header}), processed: "${processedHeader}"`);
        return processedHeader;
    });

    // console.log(`Looking for header: "${headerName}"`);
    const columnIndex = processedHeaders.indexOf(headerName) + 1;
    //console.log(`Column index for "${headerName}": ${columnIndex}`);

    if (columnIndex < 1) {
        //console.error(`Header "${headerName}" not found in spreadsheet headers!`);
        throw new Error(`Header "${headerName}" not found in spreadsheet headers!`)
    }

    const range = sheet.getRange(2, columnIndex, sheet.getLastRow() - 1, 1);
    const values = range.getValues().map(function (row) {
        return row[0];
    });
    //console.log(`Retrieved ${values.length} values from column ${columnIndex}`);
    const nonEmptyValues = values.filter(function (value) {
        return value !== "";
    });
    //console.log(nonEmptyValues)
    return nonEmptyValues;
}


function clearFilesFromPropServ() {

    PropertiesService.getUserProperties()
        .deleteProperty("files");

    PropertiesService.getUserProperties()
        .deleteProperty("fileId");
};