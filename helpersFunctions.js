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
    var body = templateDoc.getBody();
    var numChildren = body.getNumChildren();
    try {
        // Preprocess template children into JS objects for faster access
        var templateElements = [];
        for (var j = 0; j < numChildren; j++) {
            var child = body.getChild(j);
            var type = child.getType();
            if (type === DocumentApp.ElementType.PARAGRAPH) {
                templateElements.push({type: 'paragraph', element: child.copy()});
            } else if (type === DocumentApp.ElementType.TABLE) {
                templateElements.push({type: 'table', element: child.copy()});
            }
        }
        for (var x = startIndex; x <= endIndex; x++) {
            var firstPropertyValue = Object.values(dataObjects[x][0])[0];
            var newDocument = DocumentApp.create("Merged Letter " + firstPropertyValue);
            var newDocumentBody = newDocument.getBody();
            newDocumentBody.setAttributes(body.getAttributes());
            // Use preprocessed template elements
            for (var k = 0; k < templateElements.length; k++) {
                var item = templateElements[k];
                if (item.type === 'paragraph') {
                    processFormattedParagraph(item.element.asParagraph(), newDocumentBody, dataObjects[x]);
                } else if (item.type === 'table') {
                    try {
                        processFormattedTable(item.element.asTable(), newDocumentBody, dataObjects[x], k);
                    } catch (tableError) {
                        newDocumentBody.appendParagraph("[Table placeholder]");
                    }
                }
            }
            removeEmptyFirstParagraph(newDocumentBody);
            newDocument.saveAndClose();
            var parentFolder = DriveApp.getFileById(newDocument.getId()).getParents().next();
            var pdfBlob = DriveApp.getFileById(newDocument.getId()).getAs("application/pdf");
            parentFolder.createFile(pdfBlob);
            DriveApp.getFileById(newDocument.getId()).setTrashed(true);
        }
        return "PDF generation complete";
    } catch (error) {
        throw new Error(error && error.message ? error.message : String(error));
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
            var newDocument = DocumentApp.create("Merged Letter " + firstPropertyValue);
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
        throw new Error(error && error.message ? error.message : String(error));
    }
}

function escapeRegexChars(str) {
    return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
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

function processFormattedParagraph(sourceParagraph, targetBody, dataObject) {
    try {
        var newParagraph = targetBody.appendParagraph('');
        var paragraphAttributes = sourceParagraph.getAttributes();

        if (paragraphAttributes[DocumentApp.Attribute.FONT_SIZE]) {
            console.log("Source paragraph FONT_SIZE:", paragraphAttributes[DocumentApp.Attribute.FONT_SIZE]);
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
        console.log("Processing " + numChildren + " child elements");

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

        // Replace placeholders if needed
        var processedText = originalText;
        if (dataObject && originalText.indexOf("{{") >= 0) {
            processedText = replacePlaceholders(originalText, dataObject);
        }

        // Add the processed text to the new paragraph
        var appendedText = targetParagraph.appendText(processedText);

        // Copy text formatting
        var textLength = processedText.length;
        if (textLength > 0) {
            // Copy character-level attributes
            for (var charIndex = 0; charIndex < originalText.length && charIndex < textLength; charIndex++) {
                try {
                    var sourceAttributes = sourceTextElement.getAttributes(charIndex);
                    if (sourceAttributes) {
                        var targetIndex = Math.min(charIndex, textLength - 1);
                        appendedText.setAttributes(targetIndex, targetIndex, sourceAttributes);
                    }
                } catch (charError) {
                    console.log("Could not copy attributes for character " + charIndex + ": " + charError);
                }
            }

            // Apply overall text formatting
            try {
                var overallTextAttributes = sourceTextElement.getAttributes(0);
                if (overallTextAttributes) {
                    appendedText.setAttributes(0, textLength - 1, overallTextAttributes);
                }
            } catch (overallError) {
                console.log("Could not apply overall text attributes: " + overallError);
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
