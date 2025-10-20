/**
 * Formats a number with thousands separators
 * @param {Number|String} number - Number to format
 * @returns {String} Formatted number
 */
function convertNumberFormat(number) {
    var parts = number.toString().split(".");
    parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, "'");
    return parts.join(".");
}

/**
 * Converts array data to nested array format
 * @param {Array} arr - Input array
 * @param {Array} fields - Fields to apply number formatting to
 * @returns {Array} Formatted nested array
 */
function convertToNestedArray(arr, fields) {
    console.log("\n=== INSIDE convertToNestedArray ===");
    // console.log("arr length: " + arr.length);
    // console.log("fields: " + JSON.stringify(fields));

    if (!arr || arr.length === 0) {
        console.error("arr is empty or undefined!");
        return [];
    }

    var firstItem = arr[0];
    var firstKey = Object.keys(firstItem)[0];
    var firstValue = firstItem[firstKey];

    // console.log("firstKey: " + firstKey);
    // console.log("firstValue type: " + typeof firstValue);
    // console.log("firstValue is array: " + Array.isArray(firstValue));

    if (!firstValue || !Array.isArray(firstValue) || firstValue.length === 0) {
        // console.error("firstValue is not a valid array!");
        // console.log("firstValue: " + JSON.stringify(firstValue));
        return [];
    }

    var length = firstValue.length;
    // console.log("Will create " + length + " records");

    var result = [];

    for (var i = 0; i < length; i++) {
        var obj = [];
        for (var j = 0; j < arr.length; j++) {
            var key = Object.keys(arr[j])[0];
            var innerObj = {};
            var newKey = key
                .replace(/\n/g, "__")
                .replace(/\(/g, "<")
                .replace(/\)/g, ">");

            var cellValue = arr[j][key][i];

            // 🔧 DEBUG: Log cell value
            if (i === 0) {
                //console.log("Column " + j + " (" + key + "), first value: " + JSON.stringify(cellValue));
            }

            // 🔧 Handle both old format (plain text) and new format (rich text object)
            if (cellValue && typeof cellValue === 'object' && cellValue.text !== undefined) {
                // New format with rich text
                if (fields && fields.includes(key)) {
                    // Apply number formatting to the text
                    var formattedText = convertNumberFormat(cellValue.text);
                    innerObj[newKey] = {
                        text: formattedText,
                        richText: cellValue.richText
                    };
                } else {
                    innerObj[newKey] = cellValue;
                }
            } else {
                // Old format - plain text (backward compatibility)
                if (fields && fields.includes(key)) {
                    innerObj[newKey] = convertNumberFormat(cellValue);
                } else {
                    innerObj[newKey] = cellValue;
                }
            }

            obj.push(innerObj);
        }
        result.push(obj);
    }

    // console.log("convertToNestedArray returning " + result.length + " records");
    if (result.length > 0) {
        // console.log("First record: " + JSON.stringify(result[0], null, 2));
    }

    return result;
}

/**
 * Adjusts the padding of all cells in tables to control row height
 * @param {Body} body - The document body containing tables
 * @param {Number} paddingTop - Top padding value in points
 * @param {Number} paddingBottom - Bottom padding value in points
 */
// function changePaddings(body, paddingTop, paddingBottom) {
//     var tables = body.getTables();
//
//     for (var t = 0; t < tables.length; t++) {
//         var table = tables[t];
//         table.setPaddingTop()
//         for (var r = 0; r < table.getNumRows(); r++) {
//             var row = table.getRow(r);
//
//             for (var c = 0; c < row.getNumCells(); c++) {
//                 row.getCell(c).setPaddingTop(paddingTop).setPaddingBottom(paddingBottom);
//             }
//         }
//     }
// }

/**
 * Processes a document element and applies data values to placeholders
 * @param {Element} element - The document element to process
 * @param {Object} dataObject - The data to insert into placeholders
 * @returns {Element} The processed element
 */
function processElement(element, dataObject) {
    var processedElement = element.copy();

    if (element.getType() == DocumentApp.ElementType.PARAGRAPH ||
        element.getType() == DocumentApp.ElementType.TABLE) {

        for (var k = 0; k < dataObject.length; k++) {
            var obj = dataObject[k];
            for (var prop in obj) {
                if (obj.hasOwnProperty(prop)) {
                    if (obj[prop] instanceof Object) {
                        var date = new Date(obj[prop]);
                        var formattedDate =
                            ("0" + date.getDate()).slice(-2) +
                            "/" +
                            ("0" + (date.getMonth() + 1)).slice(-2) +
                            "/" +
                            date.getFullYear();
                        processedElement.asText().replaceText(`{{${prop}}}`, formattedDate);
                    } else {
                        processedElement.asText().replaceText(`{{${prop}}}`, `${obj[prop]}`);
                    }
                }
            }
        }
    }

    return processedElement;
}

/**
 * Handles document creation for both PDF and Google Docs formats
 * @param {Array} dataObjects - Array of data objects to merge
 * @param {String} fileId - ID of the template document
 * @param {Number} startIndex - Starting index for processing
 * @param {Number} endIndex - Ending index for processing
 * @param {Boolean} pdf - Whether to generate PDFs
 * @returns {String} ID of the created document
 */
function copyAndUpdateDoc2(dataObjects, fileId, startIndex, endIndex, pdf) {
    var originalDocument = DriveApp.getFileById(fileId);
    var copiedDocument = originalDocument.makeCopy();
    var newDocumentId = copiedDocument.getId();

    // Adjust index ranges
    if (startIndex && endIndex) {
        startIndex = startIndex - 2;
        endIndex = endIndex - 2;
    } else {
        startIndex = 0;
        endIndex = dataObjects.length - 1;
    }

    try {
        if (pdf) {
            console.log("Generating PDFs from template ID:", fileId);
            // PDF Generation approach
            return generatePdfs(dataObjects, fileId, startIndex, endIndex, copiedDocument);
        } else {
            console.log("Generating Google Docs from template ID:", fileId);
            // Google Docs Generation approach
            return generateGoogleDoc(dataObjects, newDocumentId, startIndex, endIndex);
        }
    } catch (error) {
        throw new Error(error && error.message ? error.message : String(error));
    }
}

/**
 * Generates individual PDFs for each data record
 * @param {Array} dataObjects - Array of data objects to merge
 * @param {String} templateId - ID of the template document
 * @param {Number} startIndex - Starting index for processing
 * @param {Number} endIndex - Ending index for processing
 * @param {File} copiedDocument - Copy of the template document
 * @returns {String} Result message
 */
function generatePdfs(dataObjects, templateId, startIndex, endIndex, copiedDocument) {
    var templateDoc = DocumentApp.openById(templateId);
    var templateBody = templateDoc.getBody();
    var numChildren = templateBody.getNumChildren();

    // Preprocess template children into JS objects for faster access (same as Google Doc)
    var templateElements = [];
    for (var j = 0; j < numChildren; j++) {
        var child = templateBody.getChild(j);
        var type = child.getType();
        if (type === DocumentApp.ElementType.PARAGRAPH) {
            templateElements.push({type: 'paragraph', element: child.copy()});
        } else if (type === DocumentApp.ElementType.TABLE) {
            templateElements.push({type: 'table', element: child.copy()});
        }
    }

    try {
        // Create a folder to store all PDFs
        var createdFiles = [];

        // Process each document individually (same loop structure as Google Doc)
        for (var x = startIndex; x <= endIndex; x++) {
            var firstPropertyValue = Object.values(dataObjects[x][0])[0];
            var newDocument = DocumentApp.create("Merged Letter " + firstPropertyValue.text);
            var newBody = newDocument.getBody();
            newBody.setAttributes(templateBody.getAttributes());

            // Use preprocessed template elements (same as Google Doc)
            for (var k = 0; k < templateElements.length; k++) {
                var item = templateElements[k];
                if (item.type === 'paragraph') {
                    processFormattedParagraph(item.element.asParagraph(), newBody, dataObjects[x]);
                } else if (item.type === 'table') {
                    try {
                        processFormattedTable(item.element.asTable(), newBody, dataObjects[x], k);
                    } catch (tableError) {
                        newBody.appendParagraph("[Table placeholder]");
                    }
                }
            }

            removeEmptyFirstParagraph(newBody);
            newDocument.saveAndClose();

            // Convert to PDF and clean up
            var docFile = DriveApp.getFileById(newDocument.getId());
            var pdfBlob = docFile.getAs("application/pdf");
            var pdfFile = DriveApp.createFile(pdfBlob).setName("Merged Letter " + firstPropertyValue.text + ".pdf");
            docFile.setTrashed(true);
            createdFiles.push(pdfFile.getUrl());

        }

        return createdFiles.length > 0 ? createdFiles[0] : null;
    } catch (error) {
        throw new Error(error && error.message ? error.message : String(error));
    }
}

function preprocessTemplateElements(body, numChildren) {
    var elements = [];
    for (var j = 0; j < numChildren; j++) {
        var child = body.getChild(j);
        var type = child.getType();
        if (type === DocumentApp.ElementType.PARAGRAPH || type === DocumentApp.ElementType.TABLE) {
            elements.push({
                type: type,
                element: child.copy(),
                text: child.getText() // Cache text content
            });
        }
    }
    return elements;
}

function processTemplateElementsBatch(templateElements, targetBody, dataObject) {
    templateElements.forEach((item, index) => {
        if (item.type === DocumentApp.ElementType.PARAGRAPH) {
            if (item.text.includes('{{')) {
                // Process with placeholder replacement
                processFormattedParagraph(item.element.asParagraph(), targetBody, dataObject);
            } else {
                // No placeholders - just copy the paragraph
                var copiedPara = item.element.copy();
                targetBody.appendParagraph(copiedPara);
            }
        } else if (item.type === DocumentApp.ElementType.TABLE) {
            if (item.text.includes('{{')) {
                // Process with placeholder replacement
                try {
                    // Pass the current body child count as the insert position
                    var currentPosition = targetBody.getNumChildren();
                    processFormattedTable(item.element.asTable(), targetBody, dataObject, currentPosition);
                } catch (tableError) {
                    console.error("Error processing table:", tableError);
                    targetBody.appendParagraph("[Table placeholder]");
                }
            } else {
                // No placeholders - just copy the table
                var copiedTable = item.element.copy();
                targetBody.appendTable(copiedTable);
            }
        }
    });
}
/**
 * Generates a Google Doc with merged data
 * @param {Array} dataObjects - Array of data objects to merge
 * @param {String} documentId - ID of the template document
 * @param {Number} startIndex - Starting index for processing
 * @param {Number} endIndex - Ending index for processing
 * @returns {String} ID of the created document
 */
function generateGoogleDoc(dataObjects, templateId, startIndex, endIndex) {
    var templateDoc = DocumentApp.openById(templateId);
    var templateBody = templateDoc.getBody();
    var numChildren = templateBody.getNumChildren();
    // Preprocess template children into JS objects for faster access
    var templateElements = [];
    for (var j = 0; j < numChildren; j++) {
        var child = templateBody.getChild(j);
        var type = child.getType();
        if (type === DocumentApp.ElementType.PARAGRAPH) {
            templateElements.push({type: 'paragraph', element: child.copy()});
        } else if (type === DocumentApp.ElementType.TABLE) {
            templateElements.push({type: 'table', element: child.copy()});
        }
    }
    var newDocIds = [];
    try {
        for (var x = startIndex; x <= endIndex; x++) {
            var firstPropertyValue = Object.values(dataObjects[x][0])[0];
            var newDocument = DocumentApp.create("Merged Letter " + firstPropertyValue.text);
            var newBody = newDocument.getBody();
            newBody.setAttributes(templateBody.getAttributes());
            // Use preprocessed template elements
            for (var k = 0; k < templateElements.length; k++) {
                var item = templateElements[k];
                if (item.type === 'paragraph') {
                    processFormattedParagraph(item.element.asParagraph(), newBody, dataObjects[x]);
                } else if (item.type === 'table') {
                    try {
                        processFormattedTable(item.element.asTable(), newBody, dataObjects[x], k);
                    } catch (tableError) {
                        newBody.appendParagraph("[Table placeholder]");
                    }
                }
            }
            removeEmptyFirstParagraph(newBody);
            newDocument.saveAndClose();
            newDocIds.push(newDocument.getId());
        }
        return newDocIds;
    } catch (error) {
        throw new Error(error && error.message ? error.message : String(error));
    }
}


/**
 * Processes a table while preserving spacing and formatting
 * @param {Table} sourceTable - The original table
 * @param {Body} targetBody - The target document body
 * @param {Object} dataObject - Data for replacing placeholders
 * @param {any} index - Data for replacing placeholders
 * @param {Body} sourceBody - Data for replacing placeholders
 * @returns {Table} The new table in the target document
 */
function processFormattedTable(sourceTable, targetBody, dataObject, index) {
    try {
        var tableCopy = sourceTable.copy();
        var numRows = tableCopy.getNumRows();

        for (var r = 0; r < numRows; r++) {
            var row = tableCopy.getRow(r);
            var numCells = row.getNumCells();

            for (var c = 0; c < numCells; c++) {
                var cell = row.getCell(c);
                var cellText = cell.getText();

                if (dataObject && cellText.indexOf("{{") >= 0) {
                    var cellTextElement = cell.editAsText();

                    // CAPTURE ORIGINAL FORMATTING BEFORE REPLACEMENT
                    var originalFontSize = cellTextElement.getFontSize(0);
                    console.log("Original font size formattedTable: " + originalFontSize + " in cell with text: " + cellText);
                    var originalFontFamily = cellTextElement.getFontFamily(0);
                    var originalForegroundColor = cellTextElement.getForegroundColor(0);

                    var richTextReplacements = [];

                    // First pass: collect all replacements and their formatting
                    for (var k = 0; k < dataObject.length; k++) {
                        var obj = dataObject[k];
                        for (var prop in obj) {
                            if (obj.hasOwnProperty(prop)) {
                                var placeholder = "{{" + prop + "}}";
                                var placeholderIndex = cellText.indexOf(placeholder);

                                if (placeholderIndex > -1) {
                                    var replacement = obj[prop];
                                    var replacementText = null;
                                    var richTextInfo = null;

                                    // Handle Date objects
                                    if (replacement instanceof Date) {
                                        replacementText = ("0" + replacement.getDate()).slice(-2) + "/" +
                                            ("0" + (replacement.getMonth() + 1)).slice(-2) + "/" +
                                            replacement.getFullYear();
                                    }
                                    else if (replacement && typeof replacement === 'object' &&
                                        typeof replacement.getMonth === 'function' &&
                                        replacement.text === undefined) {
                                        var date = new Date(replacement);
                                        replacementText = ("0" + date.getDate()).slice(-2) + "/" +
                                            ("0" + (date.getMonth() + 1)).slice(-2) + "/" +
                                            date.getFullYear();
                                    }
                                    // Handle rich text objects
                                    else if (replacement && typeof replacement === 'object' && replacement.text !== undefined) {
                                        replacementText = replacement.text;
                                        richTextInfo = replacement.richText;
                                    }
                                    else {
                                        replacementText = String(replacement);
                                    }

                                    richTextReplacements.push({
                                        placeholder: placeholder,
                                        text: replacementText,
                                        richText: richTextInfo,
                                        originalIndex: placeholderIndex
                                    });
                                }
                            }
                        }
                    }

                    // Second pass: replace text (preserves cell formatting)
                    for (var rtIdx = 0; rtIdx < richTextReplacements.length; rtIdx++) {
                        var repInfo = richTextReplacements[rtIdx];
                        var escapedPlaceholder = escapeRegexChars(repInfo.placeholder);
                        cellTextElement.replaceText(escapedPlaceholder, repInfo.text);
                    }

                    // RESTORE ORIGINAL FORMATTING TO ENTIRE CELL if it was lost
                    var updatedCellText = cell.getText();

                    // Third pass: apply rich text formatting on top
                    for (var rtIdx = 0; rtIdx < richTextReplacements.length; rtIdx++) {
                        var repInfo = richTextReplacements[rtIdx];

                        if (!repInfo.richText || !repInfo.richText.runs) continue;

                        // Find where the replaced text is now
                        var textStart = updatedCellText.indexOf(repInfo.text);
                        if (textStart === -1) continue;

                        console.log("Table cell: applying formatting to '" + repInfo.text + "' at position " + textStart);

                        // Apply formatting from each run
                        var runs = repInfo.richText.runs;
                        for (var runIdx = 0; runIdx < runs.length; runIdx++) {
                            var run = runs[runIdx];
                            var style = run.textStyle;

                            var docStartIdx = textStart + run.startIndex;
                            var docEndIdx = textStart + run.endIndex - 1;

                            try {
                                if (style.bold !== null && style.bold !== undefined) {
                                    cellTextElement.setBold(docStartIdx, docEndIdx, style.bold);
                                }
                                if (style.italic !== null && style.italic !== undefined) {
                                    cellTextElement.setItalic(docStartIdx, docEndIdx, style.italic);
                                }
                                if (style.underline !== null && style.underline !== undefined) {
                                    cellTextElement.setUnderline(docStartIdx, docEndIdx, style.underline);
                                }
                                if (style.strikethrough !== null && style.strikethrough !== undefined) {
                                    console.log("  Applying strikethrough to table cell chars " + docStartIdx + "-" + docEndIdx);
                                    cellTextElement.setStrikethrough(docStartIdx, docEndIdx, style.strikethrough);
                                }
                                // Only override font if explicitly set in rich text
                                if (style.fontFamily) {
                                    cellTextElement.setFontFamily(docStartIdx, docEndIdx, originalFontFamily);
                                }
                                if (style.fontSize) {
                                    cellTextElement.setFontSize(docStartIdx, docEndIdx, originalFontSize);
                                }
                                if (style.foregroundColor) {
                                    cellTextElement.setForegroundColor(docStartIdx, docEndIdx, style.foregroundColor);
                                }
                            } catch (styleError) {
                                console.log("  Error applying style: " + styleError);
                            }
                        }
                    }
                }
            }
        }

        var newTable = targetBody.insertTable(index, tableCopy);
        return newTable;
    } catch (error) {
        throw new Error(error && error.message ? error.message : String(error));
    }
}







// function processFormattedTable(sourceTable, targetBody, dataObject, index) {
//     try {
//         var tableCopy = sourceTable.copy();
//         var numRows = tableCopy.getNumRows();
//
//         for (var r = 0; r < numRows; r++) {
//             var row = tableCopy.getRow(r);
//             var numcells = row.getNumCells();
//             console.log("Processing row " + r + " with " + numcells + " cells");
//             //Process each cell in the row
//             for (var c = 0; c < numcells; c++) {
//                 var cell = row.getCell(c);
//                 var cellText = cell.getText();
//                 if (dataObject && cellText.indexOf("{{") >= 0) {
//                     for (var k = 0; k < dataObject.length; k++) {
//                         var obj = dataObject[k];
//                         for (var prop in obj) {
//                             if (obj.hasOwnProperty(prop)) {
//                                 var placeholder = "{{" + prop + "}}";
//                                 if (cellText.indexOf(placeholder) > -1) {
//                                     console.log("placeholder: ", placeholder);
//                                     var replacement = obj[prop];
//
//                                     // Check for actual Date objects FIRST
//                                     if (replacement instanceof Date) {
//                                         var date = replacement;
//                                         replacement = ("0" + date.getDate()).slice(-2) + "/" +
//                                             ("0" + (date.getMonth() + 1)).slice(-2) + "/" +
//                                             date.getFullYear();
//                                     }
//                                     // Check for Date-like objects (but NOT our rich text format)
//                                     else if (replacement && typeof replacement === 'object' &&
//                                         typeof replacement.getMonth === 'function' &&
//                                         replacement.text === undefined) {
//                                         var date = new Date(replacement);
//                                         replacement = ("0" + date.getDate()).slice(-2) + "/" +
//                                             ("0" + (date.getMonth() + 1)).slice(-2) + "/" +
//                                             date.getFullYear();
//                                     }
//                                     // Handle our rich text objects
//                                     else if (replacement && typeof replacement === 'object' && replacement.text !== undefined) {
//                                         if (replacement.richText.runs.strikethrough) {
//                                             console.log("Rich text has strikethrough formatting");
//                                             cell.editAsText().setStrikethrough(true);
//                                         }
//                                     }
//                                     //console.log("Replacing " + placeholder + " with " + replacement.text);
//                                     var escapedPlaceholder = escapeRegexChars(placeholder);
//                                     cell.asText().replaceText(escapedPlaceholder, replacement.text);
//                                 }
//                             }
//                         }
//                     }
//                 }
//             }
//         }
//
//         var newTable = targetBody.insertTable(index, tableCopy);
//         return newTable;
//     } catch (error) {
//         throw new Error(error && error.message ? error.message : String(error));
//     }
// }

function escapeRegexChars(str) {
    return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * Helper function to replace placeholders in text
 * @param {String} text - Text containing placeholders
 * @param {Object} dataObject - Data for replacements
 * @returns {String} Text with placeholders replaced
 */
// ============================================
// FIX 1: Replace getColumnValues() in WebApp.js
// Enhanced with Rich Text support + Cache mechanism preserved
// ============================================

function getColumnValues(headerName, sheetName) {
    const cacheKey = `${sheetName}_${headerName}_richvalues`;
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

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    const dataRange = sheet.getRange(1, 1, lastRow, lastCol);

    const allData = dataRange.getValues();
    const allRichText = dataRange.getRichTextValues();
    const allFormats = dataRange.getNumberFormats();
    const allFontLines = dataRange.getFontLines(); // 🔧 NEW: Get strikethrough for ALL cells

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

    const values = [];
    for (let i = 1; i < allData.length; i++) {
        const cellValue = allData[i][columnIndex];
        const richTextValue = allRichText[i][columnIndex];
        const numberFormat = allFormats[i][columnIndex];
        const fontLine = allFontLines[i][columnIndex]; // 🔧 Get strikethrough for this cell

        if (cellValue === "") continue;

        const richText = richTextValue.getText();

        if (richText) {
            // Rich text has content - use it with formatting
            values.push({
                text: richText,
                richText: serializeRichText(richTextValue)
            });
        } else if (cellValue) {
            // 🔧 Rich text is empty (number/date/percentage case)
            let displayValue = cellValue;
            if (numberFormat && numberFormat.includes('%') && typeof cellValue === 'number') {
                displayValue = Math.round(cellValue * 100) + '%';
            }

            // 🔧 Create synthetic rich text for formatted numbers with strikethrough
            let syntheticRichText = null;
            if (fontLine === 'line-through') {
                syntheticRichText = {
                    text: displayValue.toString(),
                    runs: [{
                        startIndex: 0,
                        endIndex: displayValue.toString().length,
                        textStyle: {
                            bold: false,
                            italic: false,
                            underline: false,
                            strikethrough: true, // 🔧 Apply strikethrough from cell format
                            fontFamily: null,
                            fontSize: null,
                            foregroundColor: null
                        }
                    }]
                };
            }

            values.push({
                text: displayValue.toString(),
                richText: syntheticRichText
            });
        }
    }

    cache.put(cacheKey, JSON.stringify(values), 360);
    return values;
}

// 🔧 NEW: Serialize RichTextValue for caching
function serializeRichText(richTextValue) {
    const text = richTextValue.getText();
    if (!text) return null;

    const runs = [];
    let currentIndex = 0;

    while (currentIndex < text.length) {
        const textStyle = richTextValue.getTextStyle(currentIndex, currentIndex + 1);

        // Find the length of this style run
        let runLength = 1;
        while (currentIndex + runLength < text.length) {
            const nextStyle = richTextValue.getTextStyle(
                currentIndex + runLength,
                currentIndex + runLength + 1
            );

            // Check if styles match
            if (!stylesMatch(textStyle, nextStyle)) break;
            runLength++;
        }

        runs.push({
            startIndex: currentIndex,
            endIndex: currentIndex + runLength,
            textStyle: {
                bold: textStyle.isBold(),
                italic: textStyle.isItalic(),
                underline: textStyle.isUnderline(),
                strikethrough: textStyle.isStrikethrough(),
                fontFamily: textStyle.getFontFamily(),
                fontSize: textStyle.getFontSize(),
                foregroundColor: textStyle.getForegroundColor()
            }
        });

        currentIndex += runLength;
    }

    return {text: text, runs: runs};
}

// 🔧 NEW: Helper to compare text styles
function stylesMatch(style1, style2) {
    return style1.isBold() === style2.isBold() &&
        style1.isItalic() === style2.isItalic() &&
        style1.isUnderline() === style2.isUnderline() &&
        style1.isStrikethrough() === style2.isStrikethrough() &&
        style1.getFontFamily() === style2.getFontFamily() &&
        style1.getFontSize() === style2.getFontSize() &&
        style1.getForegroundColor() === style2.getForegroundColor();
}


// ============================================
// FIX 2 (REVISED): Update convertToNestedArray() in helpersFunctions.js
// Better handling of data structure with more robust checks
// ============================================

function convertToNestedArray(arr, fields) {
    console.log("\n=== INSIDE convertToNestedArray ===");
    console.log("arr length: " + arr.length);
    console.log("fields: " + JSON.stringify(fields));

    if (!arr || arr.length === 0) {
        console.error("arr is empty or undefined!");
        return [];
    }

    var firstItem = arr[0];
    var firstKey = Object.keys(firstItem)[0];
    var firstValue = firstItem[firstKey];

    console.log("firstKey: " + firstKey);
    console.log("firstValue type: " + typeof firstValue);
    console.log("firstValue is array: " + Array.isArray(firstValue));

    if (!firstValue || !Array.isArray(firstValue) || firstValue.length === 0) {
        console.error("firstValue is not a valid array!");
        console.log("firstValue: " + JSON.stringify(firstValue));
        return [];
    }

    var length = firstValue.length;
    console.log("Will create " + length + " records");

    var result = [];

    for (var i = 0; i < length; i++) {
        var obj = [];
        for (var j = 0; j < arr.length; j++) {
            var key = Object.keys(arr[j])[0];
            var innerObj = {};
            var newKey = key
                .replace(/\n/g, "__")
                .replace(/\(/g, "<")
                .replace(/\)/g, ">");

            var cellValue = arr[j][key][i];

            // 🔧 DEBUG: Log cell value
            if (i === 0) {
                //console.log("Column " + j + " (" + key + "), first value: " + JSON.stringify(cellValue));
            }

            // 🔧 Handle both old format (plain text) and new format (rich text object)
            if (cellValue && typeof cellValue === 'object' && cellValue.text !== undefined) {
                // New format with rich text
                if (fields && fields.includes(key)) {
                    // Apply number formatting to the text
                    var formattedText = convertNumberFormat(cellValue.text);
                    innerObj[newKey] = {
                        text: formattedText,
                        richText: cellValue.richText
                    };
                } else {
                    innerObj[newKey] = cellValue;
                }
            } else {
                // Old format - plain text (backward compatibility)
                if (fields && fields.includes(key)) {
                    innerObj[newKey] = convertNumberFormat(cellValue);
                } else {
                    innerObj[newKey] = cellValue;
                }
            }

            obj.push(innerObj);
        }
        result.push(obj);
    }

    // console.log("convertToNestedArray returning " + result.length + " records");
    if (result.length > 0) {
        // console.log("First record: " + JSON.stringify(result[0], null, 2));
    }

    return result;
}


// ============================================
// FIX 3: Update replacePlaceholders() in helpersFunctions.js
// Track rich text replacements
// ============================================

function replacePlaceholders(text, dataObject) {
    var result = text;
    var richTextReplacements = []; // Track rich text info

    for (var k = 0; k < dataObject.length; k++) {
        var obj = dataObject[k];
        for (var prop in obj) {
            if (obj.hasOwnProperty(prop)) {
                var placeholder = "{{" + prop + "}}";
                var replacement = obj[prop];

                // 🔧 FIX: Check for Date objects FIRST, before checking for other objects
                if (replacement instanceof Date) {
                    console.log("Replacement is a Date object:", replacement);
                    var date = replacement;
                    replacement = ("0" + date.getDate()).slice(-2) +
                        "/" +
                        ("0" + (date.getMonth() + 1)).slice(-2) +
                        "/" +
                        date.getFullYear();
                }
                // 🔧 Check if it's a Date-like object with getMonth method
                else if (replacement && typeof replacement === 'object' &&
                    typeof replacement.getMonth === 'function' &&
                    !replacement.text) {
                    console.log("Replacement is a Date-like object:", replacement);
                    var date = new Date(replacement);
                    replacement = ("0" + date.getDate()).slice(-2) +
                        "/" +
                        ("0" + (date.getMonth() + 1)).slice(-2) +
                        "/" +
                        date.getFullYear();
                }
                // 🔧 Handle rich text objects (our custom format)
                else if (replacement && typeof replacement === 'object' && replacement.text !== undefined) {
                    // console.log("Replacement is a rich text object:", replacement);
                    var placeholderIndex = result.indexOf(placeholder);
                    if (placeholderIndex !== -1) {
                        richTextReplacements.push({
                            placeholder: placeholder,
                            index: placeholderIndex,
                            richText: replacement.richText,
                            text: replacement.text
                        });
                    }
                    replacement = replacement.text;
                }

                // Replace placeholder
                var escapedPlaceholder = escapeRegexChars(placeholder);
                result = result.replace(new RegExp(escapedPlaceholder, 'g'), replacement);
            }
        }
    }

    return {
        text: result,
        richTextReplacements: richTextReplacements
    };
}


/**
 * Formats a date object to a string
 * @param {Date} dateObj - Date object to format
 * @returns {String} Formatted date string
 */
function formatDate(dateObj) {
    return (
        ("0" + dateObj.getDate()).slice(-2) +
        "/" +
        ("0" + (dateObj.getMonth() + 1)).slice(-2) +
        "/" +
        dateObj.getFullYear()
    );
}

function processFormattedParagraph(sourceParagraph, targetBody, dataObject) {
    try {
        var newParagraph = targetBody.appendParagraph('');
        var paragraphAttributes = sourceParagraph.getAttributes();

        if (paragraphAttributes[DocumentApp.Attribute.FONT_SIZE]) {
            //console.log("Source paragraph FONT_SIZE:", paragraphAttributes[DocumentApp.Attribute.FONT_SIZE]);
        }
        newParagraph.setAttributes(paragraphAttributes);

        // First, identify and process positioned images (including wrap-text)
        var positionedImages = sourceParagraph.getPositionedImages();
        var wrapTextImages = [];

        if (positionedImages && positionedImages.length > 0) {
            console.log("Processing " + positionedImages.length + " positioned images");

            for (var i = 0; i < positionedImages.length; i++) {
                try {
                    var posImage = positionedImages[i];
                    var layout = posImage.getLayout();
                    var isWrapText = (layout === DocumentApp.PositionedLayout.WRAP_TEXT);

                    console.log("Image " + i + " layout:", getLayoutString(layout), "- Wrap text:", isWrapText);

                    // Store wrap-text images for special handling
                    if (isWrapText) {
                        wrapTextImages.push({
                            image: posImage,
                            index: i,
                            layout: layout
                        });
                    }

                    // Process the positioned image
                    var newPositionedImage = processPositionedImage(posImage, newParagraph);

                    if (isWrapText) {
                        console.log("Successfully processed wrap-text image");
                    }

                } catch (posImageError) {
                    console.log("Error processing positioned image: " + posImageError);
                    handlePositionedImageFallback(posImage, newParagraph);
                }
            }
        }

        // Process all child elements from the original paragraph
        var numChildren = sourceParagraph.getNumChildren();
        // console.log("Processing " + numChildren + " child elements");

        for (var i = 0; i < numChildren; i++) {
            var child = sourceParagraph.getChild(i);
            var type = child.getType();

            if (type === DocumentApp.ElementType.TEXT) {
                processTextElement(child, newParagraph, dataObject);
            } else if (type === DocumentApp.ElementType.INLINE_IMAGE) {
                processInlineImage(child, newParagraph);
            }
        }

        // Final formatting adjustments for spacing paragraphs
        applyFinalParagraphFormatting(sourceParagraph, newParagraph);

        return newParagraph;

    } catch (error) {
        throw new Error(error && error.message ? error.message : String(error));
    }
}

function processPositionedImage(posImage, targetParagraph) {
    try {
        var width = posImage.getWidth();
        var height = posImage.getHeight();
        var blob = posImage.getBlob();
        var layout = posImage.getLayout();
        var leftOffset = posImage.getLeftOffset();
        var topOffset = posImage.getTopOffset();

        var newPositionedImage = targetParagraph.addPositionedImage(blob);
        newPositionedImage.setWidth(width);
        newPositionedImage.setHeight(height);
        newPositionedImage.setLayout(layout);
        newPositionedImage.setLeftOffset(leftOffset);
        newPositionedImage.setTopOffset(topOffset);

        // Log wrap-text specific details
        if (layout === DocumentApp.PositionedLayout.WRAP_TEXT) {
            console.log("Wrap-text image processed - Width:", width, "Height:", height,
                "Offsets:", leftOffset, topOffset);
        }

        return newPositionedImage;

    } catch (error) {
        console.log("Error in processPositionedImage: " + error);
        throw error;
    }
}

function handlePositionedImageFallback(posImage, targetParagraph) {
    try {
        var fallbackImage = targetParagraph.appendInlineImage(posImage.getBlob());
        fallbackImage.setWidth(posImage.getWidth());
        fallbackImage.setHeight(posImage.getHeight());
        console.log("Added positioned image as inline image fallback (wrap-text lost)");
    } catch (fallbackError) {
        console.log("Fallback also failed: " + fallbackError);
        targetParagraph.appendText("[Positioned Image placeholder]");
    }
}

function processTextElement(textElement, targetParagraph, dataObject) {
    try {
        var sourceTextElement = textElement.asText();
        var originalText = sourceTextElement.getText();

        // Replace placeholders
        var processedText = originalText;
        var richTextReplacements = [];

        if (dataObject && originalText.indexOf("{{") >= 0) {
            var replacementResult = replacePlaceholders(originalText, dataObject);
            processedText = replacementResult.text;
            richTextReplacements = replacementResult.richTextReplacements;

            console.log("processTextElement: Found " + richTextReplacements.length + " rich text replacements");
        }

        // Add the processed text to the new paragraph
        var appendedText = targetParagraph.appendText(processedText);

        // Apply formatting from spreadsheet cells
        if (richTextReplacements.length > 0) {
            for (var r = 0; r < richTextReplacements.length; r++) {
                var repInfo = richTextReplacements[r];

                console.log("Processing replacement " + r + ": text='" + repInfo.text + "', hasRichText=" + (repInfo.richText !== null));

                // Skip if no rich text formatting (richText is null)
                if (!repInfo.richText || !repInfo.richText.runs) {
                    console.log("  -> Skipping (no richText)");
                    continue;
                }

                var startIndex = processedText.indexOf(repInfo.text);
                if (startIndex === -1) {
                    console.log("  -> Could not find text in processed text");
                    continue;
                }

                console.log("  -> Found at index " + startIndex + ", has " + repInfo.richText.runs.length + " runs");

                // Apply each formatting run
                var runs = repInfo.richText.runs;
                for (var runIdx = 0; runIdx < runs.length; runIdx++) {
                    var run = runs[runIdx];
                    var style = run.textStyle;

                    var docStartIdx = startIndex + run.startIndex;
                    var docEndIdx = startIndex + run.endIndex - 1;

                    console.log("  -> Run " + runIdx + ": chars " + docStartIdx + "-" + docEndIdx +
                        ", strikethrough=" + style.strikethrough);

                    try {
                        // Apply ALL formatting
                        if (style.bold !== null && style.bold !== undefined) {
                            appendedText.setBold(docStartIdx, docEndIdx, style.bold);
                        }
                        if (style.italic !== null && style.italic !== undefined) {
                            appendedText.setItalic(docStartIdx, docEndIdx, style.italic);
                        }
                        if (style.underline !== null && style.underline !== undefined) {
                            appendedText.setUnderline(docStartIdx, docEndIdx, style.underline);
                        }
                        if (style.strikethrough !== null && style.strikethrough !== undefined) {
                            console.log("  -> APPLYING STRIKETHROUGH to chars " + docStartIdx + "-" + docEndIdx);
                            appendedText.setStrikethrough(docStartIdx, docEndIdx, style.strikethrough);
                        }
                        if (style.fontFamily) {
                            appendedText.setFontFamily(docStartIdx, docEndIdx, style.fontFamily);
                        }
                        if (style.fontSize) {
                            appendedText.setFontSize(docStartIdx, docEndIdx, style.fontSize);
                        }
                        if (style.foregroundColor) {
                            appendedText.setForegroundColor(docStartIdx, docEndIdx, style.foregroundColor);
                        }
                    } catch (styleError) {
                        console.log("  -> ERROR applying style: " + styleError);
                    }
                }
            }
        }

        // Copy template formatting for non-replaced text
        var textLength = processedText.length;
        if (textLength > 0) {
            for (var charIndex = 0; charIndex < Math.min(originalText.length, textLength); charIndex++) {
                var isReplaced = false;
                for (var r = 0; r < richTextReplacements.length; r++) {
                    var repInfo = richTextReplacements[r];
                    var startIdx = processedText.indexOf(repInfo.text);
                    if (startIdx !== -1 && charIndex >= startIdx && charIndex < startIdx + repInfo.text.length) {
                        isReplaced = true;
                        break;
                    }
                }

                if (!isReplaced) {
                    try {
                        var sourceAttributes = sourceTextElement.getAttributes(charIndex);
                        if (sourceAttributes) {
                            appendedText.setAttributes(charIndex, charIndex, sourceAttributes);
                        }
                    } catch (charError) {
                        // Silently skip
                    }
                }
            }
        }

    } catch (error) {
        console.log("Error processing text element: " + error);
    }
}


function processInlineImage(imageElement, targetParagraph) {
    try {
        var image = imageElement.asInlineImage().copy();
        var width = image.getWidth();
        var height = image.getHeight();
        var blob = image.getBlob();

        var newImage = targetParagraph.appendInlineImage(blob);
        newImage.setWidth(width);
        newImage.setHeight(height);

        console.log("Processed inline image (no text wrapping)");

    } catch (imageError) {
        console.log("Error processing inline image: " + imageError);
        targetParagraph.appendText("[Image placeholder]");
    }
}

function applyFinalParagraphFormatting(sourceParagraph, targetParagraph) {
    try {
        var paragraphText = targetParagraph.getText();
        if (paragraphText.length <= 1) { // Empty or just newline character
            var sourceParagraphAttributes = sourceParagraph.getAttributes();

            if (sourceParagraphAttributes[DocumentApp.Attribute.FONT_SIZE]) {
                var fontSize = sourceParagraphAttributes[DocumentApp.Attribute.FONT_SIZE];
                targetParagraph.editAsText().setFontSize(fontSize);
            }

            if (sourceParagraphAttributes[DocumentApp.Attribute.UNDERLINE]) {
                var underLine = sourceParagraphAttributes[DocumentApp.Attribute.UNDERLINE];
                targetParagraph.editAsText().setUnderline(underLine);
            }
        }
    } catch (finalError) {
        console.log("Error in final paragraph formatting: " + finalError);
    }
}

function getLayoutString(layout) {
    // Helper function to convert layout enum to readable string
    switch (layout) {
        case DocumentApp.PositionedLayout.ABOVE_TEXT:
            return "ABOVE_TEXT";
        case DocumentApp.PositionedLayout.BELOW_TEXT:
            return "BELOW_TEXT";
        case DocumentApp.PositionedLayout.BREAK_BOTH:
            return "BREAK_BOTH";
        case DocumentApp.PositionedLayout.BREAK_LEFT:
            return "BREAK_LEFT";
        case DocumentApp.PositionedLayout.BREAK_RIGHT:
            return "BREAK_RIGHT";
        case DocumentApp.PositionedLayout.WRAP_TEXT:
            return "WRAP_TEXT";
        default:
            return "UNKNOWN";
    }
}

// Utility function to analyze paragraph images for debugging
function analyzeParagraphImages(paragraph) {
    console.log("=== PARAGRAPH IMAGE ANALYSIS ===");

    // Check inline images
    var inlineImages = [];
    var numChildren = paragraph.getNumChildren();

    for (var i = 0; i < numChildren; i++) {
        var child = paragraph.getChild(i);
        if (child.getType() === DocumentApp.ElementType.INLINE_IMAGE) {
            inlineImages.push(i);
        }
    }

    console.log("Inline images found:", inlineImages.length);

    // Check positioned images
    var positionedImages = paragraph.getPositionedImages();
    console.log("Positioned images found:", positionedImages ? positionedImages.length : 0);

    if (positionedImages && positionedImages.length > 0) {
        for (var i = 0; i < positionedImages.length; i++) {
            var posImage = positionedImages[i];
            var layout = posImage.getLayout();
            console.log("Positioned image " + i + ":");
            console.log("  Layout:", getLayoutString(layout));
            console.log("  Wrap text:", layout === DocumentApp.PositionedLayout.WRAP_TEXT);
            console.log("  Dimensions:", posImage.getWidth() + "x" + posImage.getHeight());
            console.log("  Offsets:", posImage.getLeftOffset(), posImage.getTopOffset());
        }
    }

    return {
        inlineCount: inlineImages.length,
        positionedCount: positionedImages ? positionedImages.length : 0,
        wrapTextCount: positionedImages ?
            positionedImages.filter(img => img.getLayout() === DocumentApp.PositionedLayout.WRAP_TEXT).length : 0
    };
}

function removeEmptyFirstParagraph(targetBody) {
    try {
        console.log("=== CHECKING FOR EMPTY FIRST PARAGRAPH ===");
        var paras = targetBody.getParagraphs();
        var firstPara = paras[0];
        if (paras.length > 1 && !firstPara.getText().trim()) {
            console.log("Found empty first paragraph, removing it...");
            firstPara.removeFromParent();
            console.log("Empty first paragraph removed successfully");
        } else {
            console.log("No empty first paragraph found or only one paragraph exists");
        }
    } catch (error) {
        console.log("Error removing empty first paragraph: " + error);
    }
}
