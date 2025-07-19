var NXT_BTN = true;

function mergeLetterFunction(e) {
  // sheetName: sheetName, fileId: fileId, field: field,  numberFormats: numberFormats
  // const { sheetName, fileId } = e.parameters
  const { sheetName, fileId } = e.parameters;
  var mergeAndFinishTypes = e && e.formInput && e.formInput.mergeAndFinishTypes;

  if (mergeAndFinishTypes === "allLetters") {
    var section1 = CardService.newCardSection().setHeader("Merge All Records");

    var merge = CardService.newButtonSet().addButton(
      CardService.newTextButton()
        .setText("Finish & Merge")
        .setOpenLink(
          CardService.newOpenLink()
            .setUrl(URL + "?sheetName=" + sheetName + "&fileId=" + fileId)
            .setOpenAs(CardService.OpenAs.OVERLAY)
            .setOnClose(CardService.OnClose.RELOAD)
        )
        .setBackgroundColor("#FFDE00")
    );

    var mergePDF = CardService.newButtonSet().addButton(
      CardService.newTextButton()
        .setText("Finish & Merge PDF")
        .setOpenLink(
          CardService.newOpenLink()
            .setUrl(
              URL +
                "?sheetName=" +
                sheetName +
                "&fileId=" +
                fileId +
                "&pdf=true"
            )
            .setOpenAs(CardService.OpenAs.OVERLAY)
            .setOnClose(CardService.OnClose.RELOAD)
        )
        .setBackgroundColor("#FFDE00")
    );

    var formatingBtn = CardService.newTextButton()
      .setText("Change Merge Fields Formating")
      .setOnClickAction(
        CardService.newAction()
          .setFunctionName("changeMergeFieldsFmt")
          .setParameters({ fileId: fileId, sheetName: sheetName })
      )
      .setBackgroundColor("#FFDE00");

    var section2 = CardService.newCardSection().setHeader(
      "Format Merge Fields (Optional)"
    );

    section2.addWidget(formatingBtn);

    section1.addWidget(merge).addWidget(mergePDF); //.addWidget(formatingBtn);

    card4.addSection(section1).addSection(section2);
    return CardService.newNavigation().pushCard(card4.build());
  } else if (mergeAndFinishTypes === "individualLetters") {
    var section1 = CardService.newCardSection().setHeader("Merge Records");
    let fileName = JSON.parse(
      PropertiesService.getUserProperties().getProperty("fileName")
    );
    let fileUrl = JSON.parse(
      PropertiesService.getUserProperties().getProperty("fileUrl")
    );

    var url = CardService.newTextParagraph().setText(
      `- <a href="${fileUrl}">${fileName}</a><br>`
    );
    var sheetNameTxt = CardService.newTextParagraph().setText(
      "Table/Sheet Name: " + sheetName
    );

    var startIndex = CardService.newTextInput()
      .setFieldName("startIndexInput")
      .setTitle("From line number: ");
    var endIndex = CardService.newTextInput()
      .setFieldName("endIndexInput")
      .setTitle("Until line number: ");

    //var imgEx = CacheService.newImage().setAltText('Example').setImageUrl('https://image.png')

    var merge = CardService.newButtonSet().addButton(
      CardService.newTextButton()
        .setText("Confirm")
        .setOnClickAction(
          CardService.newAction()
            .setFunctionName("mergeIndividualFunction")
            .setParameters({ sheetName: sheetName, fileId: fileId })
          //field: field,  numberFormats: numberFormats})
        )
        .setBackgroundColor("#FFDE00")
    );

    section1
      .addWidget(url)
      .addWidget(sheetNameTxt)
      .addWidget(startIndex)
      .addWidget(endIndex)
      .addWidget(merge);

    card4.addSection(section1);
    return CardService.newNavigation().pushCard(card4.build());
  }
}

function changeMergeFieldsFmt(e) {
  const { fileId, sheetName } = e.parameters;
  var section1 = CardService.newCardSection().setHeader("Select a Merge Field");

  // multi selection
  const multiSelect = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.CHECK_BOX)
    .setTitle("Multiple selections are allowed.")
    .setFieldName("selectedField");

  var insertedTags = getInsertedTags(fileId);
  var uniqueTags = [...new Set(insertedTags)];
  uniqueTags.forEach(function (header) {
    //  fieldsDropdown.addItem(header, header, false);
    multiSelect.addItem(header, header, false);
  });

  var section2 = CardService.newCardSection().setHeader(
    "Available Formatation:"
  );

  var uniqueFormatDescription = CardService.newTextParagraph().setText(
    "#'##0'##0.00 (ex: 1'234.56)"
  );

  var nextButton = CardService.newTextButton()
    .setText("Next")
    .setOnClickAction(
      CardService.newAction()
        .setFunctionName("confirmFieldChange")
        .setParameters({
          fileId: fileId,
          sheetName: sheetName,
        })
    )
    .setBackgroundColor("#FFDE00");

  section2.addWidget(uniqueFormatDescription);

  section1.addWidget(multiSelect).addWidget(nextButton);
  card4.addSection(section1).addSection(section2);
  return CardService.newNavigation().pushCard(card4.build());
}

function changeMergeFieldsFmtIndividual(e) {
  const { fileId, sheetName, startIndex, endIndex } = e.parameters;
  var section1 = CardService.newCardSection().setHeader("Select a Merge Field");

  // multi selection
  const multiSelect = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.CHECK_BOX)
    .setTitle("Multiple selections are allowed.")
    .setFieldName("selectedField");

  var insertedTags = getInsertedTags(fileId);
  var uniqueTags = [...new Set(insertedTags)];
  uniqueTags.forEach(function (header) {
    //  fieldsDropdown.addItem(header, header, false);
    multiSelect.addItem(header, header, false);
  });

  var section2 = CardService.newCardSection().setHeader(
    "Available Formatation:"
  );

  var uniqueFormatDescription = CardService.newTextParagraph().setText(
    "#'##0'##0.00 (ex: 1'234.56)"
  );

  var nextButton = CardService.newTextButton()
    .setText("Next")
    .setOnClickAction(
      CardService.newAction()
        .setFunctionName("confirmFieldChangeIndividual")
        .setParameters({
          fileId: fileId,
          sheetName: sheetName,
          startIndex: startIndex,
          endIndex: endIndex,
        })
    )
    .setBackgroundColor("#FFDE00");

  section2.addWidget(uniqueFormatDescription);

  section1.addWidget(multiSelect).addWidget(nextButton);
  card4.addSection(section1).addSection(section2);
  return CardService.newNavigation().pushCard(card4.build());
}
function onFieldChange2(e) {
  const { fileId, sheetName } = e.parameters;
  const field = e.formInputs.selectedField;

  // Build the card to display the selected values
  const card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle("Selected Values"))
    .addSection(
      CardService.newCardSection().addWidget(
        CardService.newTextParagraph().setText(field)
      )
    );

  // Return the action response with the new card
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().pushCard(card.build()))
    .build();
}

function onFieldChange(e) {
  const { fileId, sheetName } = e.parameters;
  var field = e.formInputs.selectedField;

  var section1 = CardService.newCardSection();

  var text = CardService.newTextParagraph().setText(
    "selected merge field: " + field
  );

  var textIput = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setTitle("Select a Number Format")
    .setFieldName("numberFormat")
    .setOnChangeAction(
      CardService.newAction()
        .setFunctionName("confirmFieldChange")
        .setParameters({
          field: field,
          sheetName: sheetName,
          fileId: fileId,
        })
    );

  var suisseFmt = ["#'##0'##0.00 (ex: 1'234.56)"];
  suisseFmt.forEach(function (header) {
    textIput.addItem(header, header, false);
  });

  section1.addWidget(text).addWidget(textIput);
  card4.addSection(section1);
  return CardService.newNavigation().pushCard(card4.build());
}

function confirmFieldChange(e) {
  const { fileId, sheetName } = e.parameters;
  const field = e.formInputs.selectedField;

  var section1 = CardService.newCardSection();

  var text = CardService.newTextParagraph().setText(
    "Selected merge field(s) to be formated: " + field
  );

  var text2 = CardService.newTextParagraph().setText(
    "Formatation : #'##0'##0.00 (ex: 1'234.56) "
  );

  PropertiesService.getUserProperties().setProperty(
    "fields",
    JSON.stringify(field)
  );

  var nextBtn = CardService.newButtonSet().addButton(
    CardService.newTextButton()
      .setText("Finish & Merge")
      .setOpenLink(
        CardService.newOpenLink()
          .setUrl(
            URL +
              "?sheetName=" +
              sheetName +
              "&fileId=" +
              fileId +
              "&field=" +
              field
          )
          .setOpenAs(CardService.OpenAs.OVERLAY)
          .setOnClose(CardService.OnClose.RELOAD)
      )
      .setBackgroundColor("#FFDE00")
  );

  section1.addWidget(text).addWidget(text2).addWidget(nextBtn);
  card4.addSection(section1);
  return CardService.newNavigation().pushCard(card4.build());
}

function confirmFieldChangeIndividual(e) {
  const { fileId, sheetName, startIndex, endIndex } = e.parameters;
  const field = e.formInputs.selectedField;

  var section1 = CardService.newCardSection();

  var text = CardService.newTextParagraph().setText(
    "Selected merge field(s) to be formated: " + field
  );

  var text2 = CardService.newTextParagraph().setText(
    "Formatation : #'##0'##0.00 (ex: 1'234.56) "
  );

  var text3 = CardService.newTextParagraph().setText(
    "Merging from record on line number " +
      startIndex +
      " until number " +
      endIndex
  );

  PropertiesService.getUserProperties().setProperty(
    "fields",
    JSON.stringify(field)
  );

  var nextBtn = CardService.newButtonSet().addButton(
    CardService.newTextButton()
      .setText("Finish & Merge")
      .setOpenLink(
        CardService.newOpenLink()
          .setUrl(
            URL +
              "?sheetName=" +
              sheetName +
              "&fileId=" +
              fileId +
              "&field=" +
              field +
              "&startIndex=" +
              startIndex +
              "&endIndex=" +
              endIndex
          )
          .setOpenAs(CardService.OpenAs.OVERLAY)
          .setOnClose(CardService.OnClose.RELOAD)
      )
      .setBackgroundColor("#FFDE00")
  );

  section1.addWidget(text).addWidget(text2).addWidget(text3).addWidget(nextBtn);
  card4.addSection(section1);
  return CardService.newNavigation().pushCard(card4.build());
}

function mergeIndividualFunction(e) {
  var sheetName = e.parameters.sheetName;
  var fileId = e.parameters.fileId;

  var field = e.parameters.field;

  // Get the value of the startIndex input
  var startIndex = e.formInput.startIndexInput;
  var endIndex = e.formInput.endIndexInput;

  var section1 = CardService.newCardSection().setHeader(
    "Please confirm your choice : "
  );

  var text = CardService.newTextParagraph().setText(
    "Merging from record on line number " +
      startIndex +
      " until number " +
      endIndex
  );

  var btn = CardService.newButtonSet().addButton(
    CardService.newTextButton()
      .setText("Finish & Merge")
      .setOpenLink(
        CardService.newOpenLink()
          .setUrl(
            URL +
              "?sheetName=" +
              sheetName +
              "&fileId=" +
              fileId +
              "&startIndex=" +
              startIndex +
              "&endIndex=" +
              endIndex
          )
          .setOpenAs(CardService.OpenAs.OVERLAY)
          .setOnClose(CardService.OnClose.RELOAD)
      )
      .setBackgroundColor("#FFDE00")
  );

  var btnPDF = CardService.newButtonSet().addButton(
    CardService.newTextButton()
      .setText("Finish & Merge PDF")
      .setOpenLink(
        CardService.newOpenLink()
          .setUrl(
            URL +
              "?sheetName=" +
              sheetName +
              "&fileId=" +
              fileId +
              "&startIndex=" +
              startIndex +
              "&endIndex=" +
              endIndex +
              "&pdf=true"
          )
          .setOpenAs(CardService.OpenAs.OVERLAY)
          .setOnClose(CardService.OnClose.RELOAD)
      )
      .setBackgroundColor("#FFDE00")
  );

  section1.addWidget(text).addWidget(btn).addWidget(btnPDF);

  var formatingBtn = CardService.newTextButton()
    .setText("Change Merge Fields Formating")
    .setOnClickAction(
      CardService.newAction()
        .setFunctionName("changeMergeFieldsFmtIndividual")
        .setParameters({
          fileId: fileId,
          sheetName: sheetName,
          startIndex: startIndex,
          endIndex: endIndex,
        })
    )
    .setBackgroundColor("#FFDE00");

  var section2 = CardService.newCardSection().setHeader(
    "Format Merge Fields (Optional)"
  );

  section2.addWidget(formatingBtn);

  card4.addSection(section1).addSection(section2);
  return CardService.newNavigation().pushCard(card4.build());
}
