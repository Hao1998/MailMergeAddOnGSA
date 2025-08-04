// const URL =
//     "https://script.google.com/a/macros/zoi.kaercher.com/s/AKfycbz14Mw-sJAjQ70eiFWifXJc3072fEIQjoLeAQtxMqel7NsTwJGs4_TfTgPkmZkI5wQ6qQ/exec";


const URL =
    "https://script.google.com/a/macros/zoi.kaercher.com/s/AKfycbzp9Dt9RcsJyc4Ijq6Xco0HVWHII6atY6ToFbRFUr4T/dev";
var sheetName = undefined;
var SHEET_NAME = "";

var card4 = CardService.newCardBuilder();

function loadExistingList(e) {
    var section1 = CardService.newCardSection().setHeader("");
    var section2 = CardService.newCardSection().setHeader("");
    let filePick = JSON.parse(
        PropertiesService.getUserProperties().getProperty("filePick")
    );
    let prop = PropertiesService.getUserProperties().getProperty("files");

    if (prop == null || !filePick) {
        var browseFileBtn = CardService.newTextButton()
            .setText("Browse")
            .setOpenLink(
                CardService.newOpenLink()
                    .setUrl(URL)
                    .setOpenAs(CardService.OpenAs.OVERLAY)
                    .setOnClose(CardService.OnClose.NOTHING)
            );
        var infoText = CardService.newTextParagraph().setText(
            "If you chose a file and it's not showing in the list, click the 'Refresh' button below."
        );

        var reloadButton = CardService.newTextButton()
            .setText("Reload")
            .setOnClickAction(
                CardService.newAction().setFunctionName("loadExistingList")
            );

        section1.addWidget(browseFileBtn).addWidget(getFilesAndFoldersDataWidget());

        section2.addWidget(infoText).addWidget(reloadButton);

        card4.addSection(section1).addSection(section2);

        //card4.addSection(section2)

        return CardService.newNavigation().pushCard(card4.build());
    } else {
        let sheetNames = getSheetNames();
        var sheetDropdown = CardService.newSelectionInput()
            .setType(CardService.SelectionInputType.DROPDOWN)
            .setTitle("Select a sheet/table")
            .setFieldName("selectedSheet")
            .setOnChangeAction(
                CardService.newAction().setFunctionName("getMergeFieldsValues")
            );

        sheetNames.forEach(function (sheetName) {
            sheetDropdown.addItem(sheetName, sheetName, false);
        });
        section1.addWidget(getFilesAndFoldersDataWidget()).addWidget(sheetDropdown);
        card4.addSection(section1);

        return CardService.newNavigation().updateCard(card4.build());
    }
}

function getMergeFieldsValues(e) {
    let sheetName = e && e.formInput && e.formInput.selectedSheet;
    var fileId = DocumentApp.getActiveDocument().getId();

    // 1. Create the TextInput widget
    var textInput = CardService.newTextParagraph().setText(
        "Place your cursor where you want to add a field. You can add multiple fields."
    );

    var fieldsDropdown = CardService.newSelectionInput()
        .setType(CardService.SelectionInputType.DROPDOWN)
        .setTitle("Select a merge field")
        .setFieldName("selectedField")
        .setOnChangeAction(
            CardService.newAction().setFunctionName("insertTextAtCursor")
        );

    var headers = getSheetData(sheetName);
    headers.forEach(function (header) {
        if (typeof header === 'number' && header % 1 === 0) {
            // If it's a whole number, convert to integer string
            header = Math.round(header).toString();
        }
        fieldsDropdown.addItem(header, header, false);
    });

    var buttonSet = CardService.newButtonSet()
        .addButton(
            CardService.newTextButton()
                .setText("Finish & Merge")
                .setOnClickAction(
                    CardService.newAction()
                        .setFunctionName("mergeAndFinish")
                        .setParameters({sheetName: sheetName, fileId: fileId})
                )
                .setBackgroundColor("#FFDE00")
        )
        .addButton(
            CardService.newTextButton()
                .setText("Remove All Placeholders")
                .setOnClickAction(
                    CardService.newAction()
                        .setFunctionName("removeAllPlaceholders")
                        .setParameters({fileId: fileId})
                )
                .setBackgroundColor("#FFDE00")
        );

    var section2 = CardService.newCardSection().setHeader("Select a merge field");

    var section3 = CardService.newCardSection().setHeader("");
    var noteInput = CardService.newTextParagraph().setText(
        "⚠️ Please make sure that your file only contains the merge fields that you need."
    );
    section3.addWidget(noteInput);

    section2.addWidget(fieldsDropdown).addWidget(buttonSet).addWidget(textInput);

    card4.addSection(section2).addSection(section3);
    return CardService.newNavigation().pushCard(card4.build());
}

// TO DELETE
function formatingFieldPage(e) {
    const {sheetName, fileId} = e.parameters;

    var section1 = CardService.newCardSection().setHeader(
        "Select a specif formating for a merge fiel (Optional)"
    );

    // FORMATING SECTION
    var formatingDropdown = CardService.newSelectionInput()
        .setType(CardService.SelectionInputType.DROPDOWN)
        .setTitle("Select a numeric field to format")
        .setFieldName("fieldFormat")
        .setOnChangeAction(
            CardService.newAction()
                .setFunctionName("updateFormatingPage")
                .setParameters({sheetName: sheetName, fileId: fileId})
        );

    var insertedTags = getInsertedTags(fileId);
    insertedTags.forEach(function (header) {
        formatingDropdown.addItem(header, header, false);
    });

    const nextButton = CardService.newTextButton()
        .setText("Skip")
        .setOnClickAction(
            CardService.newAction()
                .setFunctionName("mergeAndFinish")
                .setParameters({sheetName: sheetName, fileId: fileId})
        )
        .setBackgroundColor("#FFDE00");

    section1.addWidget(formatingDropdown).addWidget(nextButton);

    card4.addSection(section1);
    return CardService.newNavigation().pushCard(card4.build());
}

function updateFormatingPage(e) {
    const {sheetName, fileId} = e.parameters;
    const fieldFormat = e && e.formInput && e.formInput.fieldFormat;

    var section1 = CardService.newCardSection().setHeader(
        "Custom number formart of field : " + fieldFormat
    );

    var textIput = CardService.newSelectionInput()
        .setType(CardService.SelectionInputType.DROPDOWN)
        .setTitle("Select a Number Format")
        .setFieldName("numberFormats");
    // .setOnChangeAction(CardService.newAction().setFunctionName('mergeAndFinish').setParameters({ field: fieldFormat , sheetName: sheetName, fileId: fileId}))

    var insertedTag = ["#'##0'##0.00 (ex: 1'234.56)"];
    insertedTag.forEach(function (header) {
        textIput.addItem(header, header, false);
    });

    const nextButton = CardService.newTextButton()
        .setText("Next")
        .setOnClickAction(
            CardService.newAction().setFunctionName("mergeAndFinish").setParameters({
                field: fieldFormat,
                sheetName: sheetName,
                fileId: fileId,
            })
        )
        .setBackgroundColor("#FFDE00");

    section1.addWidget(textIput).addWidget(nextButton);

    card4.addSection(section1);
    return CardService.newNavigation().updateCard(card4.build());
}

// END TO DELETE

function mergeAndFinish(e) {
    const {sheetName, fileId, field} = e.parameters;
    const numberFormats = e && e.formInput && e.formInput.numberFormats;

    var section1 = CardService.newCardSection().setHeader("Merge Records");
    var radioButtonGroup = CardService.newSelectionInput()
        .setType(CardService.SelectionInputType.RADIO_BUTTON)
        .setFieldName("mergeAndFinishTypes")
        .addItem("All", "allLetters", true)
        .addItem("Individual (Recommended)", "individualLetters", false);
    /*  .setOnChangeAction(
        CardService.newAction()
          .setFunctionName("mergeLetterFunction")
          .setParameters({ sheetName: sheetName, fileId: fileId })
      ); */
    //, field: field,  numberFormats: numberFormats}))

    var nextBtn = CardService.newTextButton()
        .setText("Next")
        .setOnClickAction(
            CardService.newAction()
                .setFunctionName("mergeLetterFunction")
                .setParameters({
                    fileId: fileId,
                    sheetName: sheetName,
                })
        )
        .setBackgroundColor("#FFDE00");

    section1.addWidget(radioButtonGroup).addWidget(nextBtn);

    card4.addSection(section1);
    return CardService.newNavigation().pushCard(card4.build());
}

function addValuesFunction(sheetName, fileId, startIndex, endIndex, pdf) {

    console.log("Starting addValuesFunction with:", {sheetName, fileId, startIndex, endIndex, pdf});
    const headersTags = getInsertedTags(fileId);
    console.log("Retrieved headersTags:", headersTags);
    const values = [];

    if (!headersTags || headersTags.length === 0) {
        console.error("No merge fields found in the document!");
        throw new Error("No merge fields found in the document. Please add merge fields and try again.");
    }

    headersTags.forEach((tag) => {
        const obj = {};
        //console.log("tag     :", tag )
        obj[tag] = getColumnValues(tag, sheetName);
        values.push(obj);
    });


    const fieldsToFormat = JSON.parse(
        PropertiesService.getUserProperties().getProperty("fields")
    );

    var formattedValues = convertToNestedArray(values, fieldsToFormat);

    PropertiesService.getUserProperties().deleteProperty("fields");

    copyAndUpdateDoc2(formattedValues, fileId, startIndex, endIndex, pdf);

// showDebugLog();
}

function insertTextAtCursor(e) {
    var document = DocumentApp.getActiveDocument();
    const selectedField = e && e.formInput && e.formInput.selectedField;
    var cursor = document.getCursor();
    if (cursor) {
        var textToInsert = `{{${selectedField}}}`;

        if (textToInsert.match(/(\r\n|\n|\r)/gm)) {
            //  textToInsert.replace(/\r/g, "\n")
            var newTextToInsert = textToInsert
                .replace(/\n/g, "__")
                .replace(/\(/g, "<")
                .replace(/\)/g, ">");
            //.replace(/\(/g, "__(").replace(/\)/g, ")__")
            cursor.insertText(newTextToInsert).setBold(true);
        } else {
            cursor.insertText(textToInsert).setBold(true);
        }
    }
}

function getInsertedTags(fileId) {
    var doc = DocumentApp.openById(fileId);

    var body = doc.getBody();
    var text = body.getText();
    var regex = /{{.+?}}/g; // -> /{{([\s\S]*?)}}/g;
    var matches = text.match(regex);
    let values = [];
    if (matches) {
        for (var i = 0; i < matches.length; i++) {
            var value = matches[i].replace("{{", "").replace("}}", "");
            //values.push(value.replace(/\r?\n|\r/g, " "));
            values.push(
                value.replace(/__/g, "\n").replace(/\</g, "(").replace(/\>/g, ")")
            );
            //.replace(/__\(/g, "(").replace(/__\)/g, ")"))//value.replace(/\r/g, "\n"));

            values.push;
        }
    }

    return values;
}

function getFilesAndFoldersDataWidget() {
    let filePick = JSON.parse(
        PropertiesService.getUserProperties().getProperty("filePick")
    );
    let prop = PropertiesService.getUserProperties().getProperty("files");

    let paragraph = "";
    if (prop == null || !filePick) {
        paragraph = `<i>Use an existing List.</i>`;

        clearFilesFromPropServ(); // Ensures there are no files in the Properties Service;
    } else {
        let docs = JSON.parse(prop);

        let fileName = JSON.parse(
            PropertiesService.getUserProperties().getProperty("fileName")
        );
        let fileUrl = JSON.parse(
            PropertiesService.getUserProperties().getProperty("fileUrl")
        );

        if (fileName && fileUrl) {
            paragraph += `- <a href="${fileUrl}">${fileName}</a><br>`;
        }

        /* docs.forEach(doc => {
          DOC_ID = doc.id
          paragraph += `- <a href="${doc.url}">${doc.name}</a><br>`
        })  */
    }

    PropertiesService.getUserProperties().setProperty(
        "filePick",
        JSON.stringify(false)
    );

    return CardService.newTextParagraph().setText(paragraph);
}

function removeAllPlaceholders(e) {
    var doc = DocumentApp.getActiveDocument();
    var body = doc.getBody();

    // Helper function to recursively remove placeholders in all elements
    function removePlaceholdersFromElement(element) {
        var type = element.getType();
        if (type === DocumentApp.ElementType.PARAGRAPH) {
            var text = element.getText();
            if (text.includes("{{PLZ}}")){
                console.log("{{PLZ}} placeholder found: " + text);
            }
            var newText = text.replace(/{{.+?}}/g, " ");
            if (text.includes("{{PLZ}}")){
                console.log("{{PLZ}} placeholder found new text: ", newText);
            }
            if (text.trim() === "") {
                // Ignore empty/blank original text
                return;
            } else if (newText.trim() === "") {
                // Remove the element if newText is empty/whitespace
                var parent = element.getParent();
                parent.removeChild(element);
            } else if (text !== newText) {
                element.setText(newText);
            }
        } else if (type === DocumentApp.ElementType.TABLE) {
            var table = element.asTable();
            console.log("Processing table with " + table.getNumRows() + " rows.")
            for (var r = 0; r < table.getNumRows(); r++) {
                var row = table.getRow(r);
                for (var c = 0; c < row.getNumCells(); c++) {
                    console.log("Processing cell at row " + r + ", column " + c);
                    var cell = row.getCell(c);
                    console.log("Cell text before: " + cell.getText());

                    // Process each child element in the cell (usually paragraphs)
                    for (var p = 0; p < cell.getNumChildren(); p++) {
                        var childElement = cell.getChild(p);
                        console.log("Processing child element at index " + p + " type: " + childElement.getType());

                        // Recursively process the child element (this handles paragraphs and their text)
                        removePlaceholdersFromElement(childElement);
                    }
                }
            }
        } else if (type === DocumentApp.ElementType.LIST_ITEM) {
            var text = element.getText();
            var newText = text.replace(/{{.+?}}/g, " ");
            if (text.trim() === "") {
                return;
            } else if (newText.trim() === "") {
                var parent = element.getParent();
                parent.removeChild(element);
            } else if (text !== newText) {
                element.setText(newText);
            }
        } else if (element.getNumChildren && element.getNumChildren() > 0) {
            // Recursively process children
            // Note: iterate backwards to safely remove children
            for (var i = element.getNumChildren() - 1; i >= 0; i--) {
                removePlaceholdersFromElement(element.getChild(i));
            }
        }
    }

    removePlaceholdersFromElement(body);
    doc.saveAndClose();
    var card = CardService.newCardBuilder()
        .setHeader(CardService.newCardHeader().setTitle("All placeholders removed!"))
        .addSection(CardService.newCardSection().addWidget(
            CardService.newTextParagraph().setText("All placeholders have been removed from the document, including inside tables.")
        ));
    return CardService.newNavigation().updateCard(card.build());
}
