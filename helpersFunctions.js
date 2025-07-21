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
    var result = [];
    var length = arr[0][Object.keys(arr[0])[0]].length;

    for (var i = 0; i < length; i++) {
        var obj = [];
        for (var j = 0; j < arr.length; j++) {
            var key = Object.keys(arr[j])[0];
            var innerObj = {};
            var newKey = key
                .replace(/\n/g, "__")
                .replace(/\(/g, "<")
                .replace(/\)/g, ">");
            if (fields && fields.includes(key)) {
                innerObj[newKey] = convertNumberFormat(arr[j][key][i]);
            } else {
                innerObj[newKey] = arr[j][key][i];
            }
            obj.push(innerObj);
        }
        result.push(obj);
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
            // PDF Generation approach
            return generatePdfs(dataObjects, fileId, startIndex, endIndex, copiedDocument);
        } else {
            // Google Docs Generation approach
            return generateGoogleDoc(dataObjects, newDocumentId, startIndex, endIndex);
        }
    } catch (error) {
        console.log("ERROR in copyAndUpdateDoc2: " + error.toString());
        throw error;
        //throw new Error("Document processing failed: " + error.message);
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
    var body = templateDoc.getBody();
    var numChildren = body.getNumChildren();
    var currentChildType = null;

    try {
        for (var x = startIndex; x <= endIndex; x++) {
            // Get identifier for file naming
            var firstPropertyValue = Object.values(dataObjects[x][0])[0];

            // Create a new document for each record
            var newDocument = DocumentApp.create("Merged Letter " + firstPropertyValue);
            var newDocumentBody = newDocument.getBody();


            newDocumentBody.setAttributes(body.getAttributes());
            // Process all elements from template
            for (var j = 0; j < numChildren; j++) {
                var child = body.getChild(j);
                currentChildType = child.getType();
                var paragraphVisited = false;
                if (currentChildType === DocumentApp.ElementType.PARAGRAPH) {
                    // Use the enhanced paragraph processor to preserve formatting
                    processFormattedParagraph(child.asParagraph(), newDocumentBody, dataObjects[x]);
                } else if (currentChildType === DocumentApp.ElementType.TABLE) {
                    try {
                        processFormattedTable(child.asTable(), newDocumentBody, dataObjects[x], j)
                    } catch (tableError) {
                        console.log("Error with table: " + tableError);
                        newDocumentBody.appendParagraph("[Table placeholder]");
                    }
                } else {
                    console.log("Skipping unsupported element type: " + currentChildType);
                }
            }

            // Apply table padding adjustments
            // changePaddings(newDocumentBody, 0, 0);
            // Remove empty first paragraph if it exists
            // Save document
            //removeEmptyFirstParagraph(newDocumentBody);
            newDocument.saveAndClose();

            // Get parent folder for PDF output
            var parentFolder = DriveApp.getFileById(newDocument.getId()).getParents().next();

            // Add delay to avoid rate limits
            Utilities.sleep(1000);

            // Convert to PDF
            var pdfBlob = DriveApp.getFileById(newDocument.getId()).getAs("application/pdf");
            parentFolder.createFile(pdfBlob);

            // Clean up temporary document
            DriveApp.getFileById(newDocument.getId()).setTrashed(true);

            // Add delay between operations
            Utilities.sleep(1000);
        }

        // Clean up
        DriveApp.getFileById(newDocument.getId()).setTrashed(true);

        return "PDF generation complete";
    } catch (error) {
        console.log("ERROR childType: " + JSON.stringify(currentChildType));
        console.log("ERROR in generatePdfs: " + error.toString());
        console.log("Stack trace: " + (error.stack || "No stack trace available"));
        throw error;
    }
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
    var templateBody = templateDoc.getBody().copy();
    var numChildren = templateBody.getNumChildren();
    var batchSize = 30;
    var newDocIds = [];

    try {
        for (var x = startIndex; x <= endIndex; x++) {
            // Get identifier for file naming
            var firstPropertyValue = Object.values(dataObjects[x][0])[0];
            // Create a new document for each record
            var newDocument = DocumentApp.create("Merged Letter " + firstPropertyValue);
            var newBody = newDocument.getBody();
            newBody.setAttributes(templateBody.getAttributes());

            for (var j = 0; j < numChildren; j++) {
                var child = templateBody.getChild(j);
                var type = child.getType();
                if (type === DocumentApp.ElementType.PARAGRAPH) {
                    processFormattedParagraph(child.asParagraph(), newBody, dataObjects[x]);
                } else if (type === DocumentApp.ElementType.TABLE) {
                    try {
                        processFormattedTable(child.asTable(), newBody, dataObjects[x], j);
                    } catch (tableError) {
                        console.log("Error with table: " + tableError);
                        newBody.appendParagraph("[Table placeholder]");
                    }
                } else {
                    console.log("Skipping unsupported element type: " + type);
                }
                removeEmptyFirstParagraph(newBody);
                // Save changes in batches to avoid timeout
                if (j > 0 && j % batchSize === 0) {
                    newDocument.saveAndClose();
                    newDocument = DocumentApp.openById(newDocument.getId());
                    newBody = newDocument.getBody();
                }
            }

            newDocument.saveAndClose();
            newDocIds.push(newDocument.getId());
        }
        return newDocIds;
    } catch (error) {
        console.log("ERROR in generateGoogleDoc: " + error.toString());
        throw error;
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
        // Create a copy of the source table
        var tableCopy = sourceTable.copy();

        // Process the copied table
        var numRows = tableCopy.getNumRows();
        for (var r = 0; r < numRows; r++) {
            var row = tableCopy.getRow(r);
            var rowText = row.getText();
            var hasPlaceholder = false;

            // Check if row contains placeholders
            if (dataObject && rowText.indexOf("{{") >= 0) {
                hasPlaceholder = true;
                // console.log("Row " + r + " contains placeholders");
                // console.log("dataObject,   ", JSON.stringify(dataObject))
                // Replace placeholders
                for (var k = 0; k < dataObject.length; k++) {
                    var obj = dataObject[k];
                    for (var prop in obj) {
                        if (obj.hasOwnProperty(prop)) {

                            var placeholder = "{{" + prop + "}}";
                            if (rowText.indexOf(placeholder) > -1) {
                                console.log("placeholder: ", placeholder)
                                var replacement = obj[prop];

                                if (obj[prop] instanceof Object) {
                                    var date = new Date(obj[prop]);
                                    replacement = ("0" + date.getDate()).slice(-2) +
                                        "/" +
                                        ("0" + (date.getMonth() + 1)).slice(-2) +
                                        "/" +
                                        date.getFullYear();
                                }

                                console.log("Replacing " + placeholder + " with " + replacement);
                                var escapedPlaceholder = escapeRegexChars(placeholder);
                                row.asText().replaceText(escapedPlaceholder, replacement);
                            }
                        }
                    }
                }
            }
        }

        // Insert the processed table
        var newTable = targetBody.insertTable(index, tableCopy);
        return newTable;
    } catch (error) {
        console.log("Error processing formatted table: " + error);
        throw error;
    }
}

function escapeRegexChars(str) {
    return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
* Get the string representing the given PositionedLayout enum.
* @param {PositionedLayout} PositionedLayout - Enum value.
* @returns {String} English text matching enum.
*/
function getLayoutString(PositionedLayout) {
    var layout;
    switch (PositionedLayout) {
        case DocumentApp.PositionedLayout.ABOVE_TEXT:
            layout = "ABOVE_TEXT";
            break;
        case DocumentApp.PositionedLayout.BREAK_BOTH:
            layout = "BREAK_BOTH";
            break;
        case DocumentApp.PositionedLayout.BREAK_LEFT:
            layout = "BREAK_LEFT";
            break;
        case DocumentApp.PositionedLayout.BREAK_RIGHT:
            layout = "BREAK_RIGHT";
            break;
        case DocumentApp.PositionedLayout.WRAP_TEXT:
            layout = "WRAP_TEXT";
            break;
        default:
            layout = "UNKNOWN";
            break;
    }
    return layout;
}

/**
* Helper function to replace placeholders in text
* @param {String} text - Text containing placeholders
* @param {Object} dataObject - Data for replacements
                                        * @returns {String} Text with placeholders replaced
*/
function replacePlaceholders(text, dataObject) {
    var result = text;
    for (var k = 0; k < dataObject.length; k++) {
        var obj = dataObject[k];
        for (var prop in obj) {
            if (obj.hasOwnProperty(prop)) {
                var placeholder = "{{" + prop + "}}";
                var replacement = obj[prop];

                if (obj[prop] instanceof Object) {
                    var date = new Date(obj[prop]);
                    replacement = ("0" + date.getDate()).slice(-2) +
                        "/" +
                        ("0" + (date.getMonth() + 1)).slice(-2) +
                        "/" +
                        date.getFullYear();
                }

                // Use the escape function for regex-safe replacement
                var escapedPlaceholder = escapeRegexChars(placeholder);
                result = result.replace(new RegExp(escapedPlaceholder, 'g'), replacement);
            }
        }
    }
    return result;
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

/**
 * Processes a paragraph and preserves its formatting, handling cross-document copying
 * @param {GoogleAppsScript.Document.Paragraph} sourceParagraph - The original paragraph
 * @param {GoogleAppsScript.Document.Body} targetBody - The target document body
 * @param {Object} dataObject - Data for replacements
 * @returns {Paragraph} The new paragraph in the target document
 */
function processFormattedParagraph(sourceParagraph, targetBody, dataObject) {
    try {
        var newParagraph = targetBody.appendParagraph('');
        var paragraphAttributes = sourceParagraph.getAttributes();
        // console.log("=== PARAGRAPH ATTRIBUTES ===");
        // console.log("Source paragraph attributes:", JSON.stringify(paragraphAttributes, null, 2));
        if (paragraphAttributes[DocumentApp.Attribute.FONT_SIZE]) {
            console.log("Source paragraph FONT_SIZE:", paragraphAttributes[DocumentApp.Attribute.FONT_SIZE]);
        }
        newParagraph.setAttributes(paragraphAttributes);
        // Process all child elements from the original paragraph
        var numChildren = sourceParagraph.getNumChildren();
        // console.log("Processing " + numChildren + " child elements");
        for (var i = 0; i < numChildren; i++) {
            var child = sourceParagraph.getChild(i);
            var type = child.getType();
            if (type === DocumentApp.ElementType.TEXT) {
                // Get the original text element
                var sourceTextElement = child.asText();
                var originalText = sourceTextElement.getText();
                // console.log("=== TEXT ELEMENT ===");
                // console.log("Original text: '" + originalText + "'");
                // Replace placeholders if needed
                var processedText = originalText;
                if (dataObject && originalText.indexOf("{{") >= 0) {
                    processedText = replacePlaceholders(originalText, dataObject);
                    // console.log("Text after placeholder replacement: '" + processedText + "'");
                }
                // Add the processed text to the new paragraph
                var appendedText = newParagraph.appendText(processedText);
                // Now copy ALL text formatting from the original text element
                // This is crucial for preserving font size and other text-level formatting
                var textLength = processedText.length;
                if (textLength > 0) {
                    // Copy text attributes for each character/range
                    for (var charIndex = 0; charIndex < originalText.length && charIndex < textLength; charIndex++) {
                        try {
                            var sourceAttributes = sourceTextElement.getAttributes(charIndex);
                            if (sourceAttributes) {
                                var targetIndex = Math.min(charIndex, textLength - 1);
                                appendedText.setAttributes(targetIndex, targetIndex, sourceAttributes);
                            }
                        } catch (charError) {
                            // console.log("Could not copy attributes for character " + charIndex + ": " + charError);
                        }
                    }
                    // Also apply overall text formatting to the entire appended text
                    try {
                        var overallTextAttributes = sourceTextElement.getAttributes(0);
                        if (overallTextAttributes) {
                            // console.log("Overall text attributes:", JSON.stringify(overallTextAttributes, null, 2));
                            appendedText.setAttributes(0, textLength - 1, overallTextAttributes);
                        }
                    } catch (overallError) {
                        console.log("Could not apply overall text attributes: " + overallError);
                    }
                }
            } else if (type === DocumentApp.ElementType.INLINE_IMAGE) {
                // Handle inline images (your existing code is fine)
                try {
                    var image = child.asInlineImage().copy();
                    var width = image.getWidth();
                    var height = image.getHeight();
                    var blob = image.getBlob();
                    var newImage = newParagraph.appendInlineImage(blob);
                    newImage.setWidth(width);
                    newImage.setHeight(height);
                } catch (imageError) {
                    console.log("Error processing inline image: " + imageError);
                    newParagraph.appendText("[Image placeholder]");
                }
            }
        }
        // Handle positioned images (your existing code is fine)
        var positionedImages = sourceParagraph.getPositionedImages();
        if (positionedImages && positionedImages.length > 0) {
            // console.log("Processing " + positionedImages.length + " positioned images");
            for (var i = 0; i < positionedImages.length; i++) {
                try {
                    var posImage = positionedImages[i];
                    var width = posImage.getWidth();
                    var height = posImage.getHeight();
                    var blob = posImage.getBlob();
                    var layout = posImage.getLayout();
                    var leftOffset = posImage.getLeftOffset();
                    var topOffset = posImage.getTopOffset();
                    var newPositionedImage = newParagraph.addPositionedImage(blob);
                    newPositionedImage.setWidth(width);
                    newPositionedImage.setHeight(height);
                    newPositionedImage.setLayout(layout);
                    newPositionedImage.setLeftOffset(leftOffset);
                    newPositionedImage.setTopOffset(topOffset);
                    // console.log("Successfully added positioned image with layout: " + getLayoutString(layout));
                } catch (posImageError) {
                    // console.log("Error processing positioned image: " + posImageError);
                    try {
                        var fallbackImage = newParagraph.appendInlineImage(posImage.getBlob());
                        fallbackImage.setWidth(posImage.getWidth());
                        fallbackImage.setHeight(posImage.getHeight());
                        console.log("Added positioned image as inline image fallback");
                    } catch (fallbackError) {
                        console.log("Fallback also failed: " + fallbackError);
                        newParagraph.appendText("[Positioned Image placeholder]");
                    }
                }
            }
        }
        // Final check: ensure the paragraph has the correct font size for spacing
        // This is especially important for empty or minimal text paragraphs used for spacing
        try {
            var finalParagraphAttributes = newParagraph.getAttributes();
            // console.log("=== FINAL PARAGRAPH CHECK ===");
            // console.log("Final paragraph attributes:", JSON.stringify(finalParagraphAttributes, null, 2));
            // if (finalParagraphAttributes[DocumentApp.Attribute.FONT_SIZE]) {
            //     console.log("Final paragraph FONT_SIZE:", finalParagraphAttributes[DocumentApp.Attribute.FONT_SIZE]);
            // }
            // If the paragraph is empty or has minimal text, ensure font size is preserved
            var paragraphText = newParagraph.getText();
            if (paragraphText.length <= 1) { // Empty or just newline character
                var sourceParagraphAttributes = sourceParagraph.getAttributes();
                if (sourceParagraphAttributes[DocumentApp.Attribute.FONT_SIZE]) {
                    var fontSize = sourceParagraphAttributes[DocumentApp.Attribute.FONT_SIZE];
                    // console.log("Ensuring font size " + fontSize + " for spacing paragraph");
                    newParagraph.editAsText().setFontSize(fontSize);
                }

                if (sourceParagraphAttributes[DocumentApp.Attribute.UNDERLINE]) {
                    var underLine = sourceParagraphAttributes[DocumentApp.Attribute.UNDERLINE];
                    newParagraph.editAsText().setUnderline(underLine)
                }
            }
        } catch (finalError) {
            console.log("Error in final paragraph check: " + finalError);
        }
        return newParagraph;
    } catch (error) {
        console.log("Error processing formatted paragraph: " + error);
        console.log("Stack: " + (error.stack || "No stack trace"));
        throw error;
    }
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
