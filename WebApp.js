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

// function getColumnValues(headerName, sheetName) {
//     const cacheKey = `${sheetName}_${headerName}_richvalues`;
//     const cache = CacheService.getUserCache();
//     const cachedValues = cache.get(cacheKey);
//
//     if (cachedValues) {
//         return JSON.parse(cachedValues);
//     }
//
//     const spreadsheetId = JSON.parse(PropertiesService.getUserProperties().getProperty('fileId'));
//     const sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);
//
//     if (!sheet) {
//         throw new Error(`Sheet "${sheetName}" not found in spreadsheet.`);
//     }
//
//     const lastRow = sheet.getLastRow();
//     const lastCol = sheet.getLastColumn();
//     const dataRange = sheet.getRange(1, 1, lastRow, lastCol);
//
//     const allData = dataRange.getValues();
//     const allRichText = dataRange.getRichTextValues();
//     const allFormats = dataRange.getNumberFormats();
//     const allFontLines = dataRange.getFontLines(); // 🔧 NEW: Get strikethrough for ALL cells
//
//     const headers = allData[0].map((header, index) => {
//         if (typeof header === 'number' && header % 1 === 0) {
//             return Math.round(header).toString();
//         }
//         return header;
//     });
//
//     const columnIndex = headers.indexOf(headerName);
//     if (columnIndex === -1) {
//         throw new Error(`Header "${headerName}" not found in spreadsheet headers!`);
//     }
//
//     const values = [];
//     for (let i = 1; i < allData.length; i++) {
//         const cellValue = allData[i][columnIndex];
//         const richTextValue = allRichText[i][columnIndex];
//         const numberFormat = allFormats[i][columnIndex];
//         const fontLine = allFontLines[i][columnIndex]; // 🔧 Get strikethrough for this cell
//
//         if (cellValue === "") continue;
//
//         const richText = richTextValue.getText();
//
//         if (richText) {
//             // Rich text has content - use it with formatting
//             values.push({
//                 text: richText,
//                 richText: serializeRichText(richTextValue)
//             });
//         } else if (cellValue) {
//             // 🔧 Rich text is empty (number/date/percentage case)
//             let displayValue = cellValue;
//             if (numberFormat && numberFormat.includes('%') && typeof cellValue === 'number') {
//                 displayValue = Math.round(cellValue * 100) + '%';
//             }
//
//             // 🔧 Create synthetic rich text for formatted numbers with strikethrough
//             let syntheticRichText = null;
//             if (fontLine === 'line-through') {
//                 syntheticRichText = {
//                     text: displayValue.toString(),
//                     runs: [{
//                         startIndex: 0,
//                         endIndex: displayValue.toString().length,
//                         textStyle: {
//                             bold: false,
//                             italic: false,
//                             underline: false,
//                             strikethrough: true, // 🔧 Apply strikethrough from cell format
//                             fontFamily: null,
//                             fontSize: null,
//                             foregroundColor: null
//                         }
//                     }]
//                 };
//             }
//
//             values.push({
//                 text: displayValue.toString(),
//                 richText: syntheticRichText
//             });
//         }
//     }
//
//     cache.put(cacheKey, JSON.stringify(values), 360);
//     return values;
// }



function clearFilesFromPropServ() {

    PropertiesService.getUserProperties()
        .deleteProperty("files");

    PropertiesService.getUserProperties()
        .deleteProperty("fileId");
};