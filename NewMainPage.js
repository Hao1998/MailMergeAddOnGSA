var NEXT_BTN_TGL = true;
var CLEAR_BTN_TGL = true;
function onhomepage(e) {
  return createSelectionCard(e);
}

/**
 * ## Initial Card Builder ##
 * Main function to generate the initial card on load.
 * @param {Object} e : Event object.
 * @return {CardService.Card} The card to show the user.
 */
function createSelectionCard(e) {
  const builder = CardService.newCardBuilder();

  builder
    .addSection(lettersDetailsSection())
    .addSection(buildSection())
    .addSection(createNewSectionWithButton());

  return builder.build();
}

/**
 * ## Rebuild Card Builder ##
 * This is activated after first initialisation of the homepage.
 *
 * Main function to generate the layout of the homepage card when button clicked.
 * @param {Object} e : Event object.
 * @return {CardService.Card} The card to show the user.
 */
function rebuildSelectionCard(e) {
  PropertiesService.getUserProperties().setProperty(
    "filePick",
    JSON.stringify(false)
  );

  const builder = CardService.newCardBuilder();

  builder
    .addSection(lettersDetailsSection())
    .addSection(buildSection())
    .addSection(createNewSectionWithButton());

  return CardService.newNavigation().updateCard(builder.build());
}

function lettersDetailsSection() {
  const newSection = CardService.newCardSection()
    .setHeader("Letters")
    .addWidget(
      CardService.newTextParagraph().setText(
        "Send letters to a group of people. You can personalize the letter that each person receives."
      )
    );

  return newSection;
}

function createNewSectionWithButton() {
  const button = CardService.newTextButton()
    .setText("Next")
    .setOnClickAction(
      CardService.newAction().setFunctionName("getMergeFieldsWidget")
    )
    .setDisabled(NEXT_BTN_TGL)
    .setBackgroundColor("#FFDE00");
  const newSection = CardService.newCardSection().addWidget(button);

  return newSection;
}

var card4 = CardService.newCardBuilder();
function getMergeFieldsWidget() {
  var section1 = CardService.newCardSection().setHeader("");
  let sheetNames = getSheetNames();

  var sheetDropdown = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setTitle("Select a sheet/table")
    .setFieldName("selectedSheet");
  // .setOnChangeAction(CardService.newAction().setFunctionName('getMergeFieldsValues'))

  sheetNames.forEach(function (sheetName) {
    sheetDropdown.addItem(sheetName, sheetName, false);
  });

  var nextBtn = CardService.newTextButton()
    .setText("Next")
    .setOnClickAction(
      CardService.newAction().setFunctionName("getMergeFieldsValues")
    )
    .setBackgroundColor("#FFDE00");

  section1.addWidget(sheetDropdown).addWidget(nextBtn);
  card4.addSection(section1);

  return CardService.newNavigation().pushCard(card4.build());
}

/**
 * Builds a section for the card service card.
 *
 * @return {CardService.CardSection}
 */
function buildSection() {
  // ## Widgets ##

  let textWidget = () => {
    return CardService.newTextParagraph().setText("");
  };

  let buttonPickerWidget = () => {
    let button = CardService.newTextButton()
      .setText("Browse")
      .setOpenLink(
        CardService.newOpenLink()
          .setUrl(URL)
          .setOpenAs(CardService.OpenAs.OVERLAY)
          .setOnClose(CardService.OnClose.RELOAD)
      )
      .setBackgroundColor("#FFDE00");

    return button;
  };

  // Must be done on a rebuild.
  let buttonFileRemoveWidget = () => {
    let button = CardService.newTextButton()
      .setText("Delete file")
      .setOnClickAction(
        CardService.newAction().setFunctionName("rebuildSelectionCard")
      )
      .setBackgroundColor("#D1E3F8")
      .setDisabled(CLEAR_BTN_TGL);
    return button;
  };

  var buttonSet = () => {
    let bSet = CardService.newButtonSet()
      .addButton(buttonPickerWidget())
      .addButton(buttonFileRemoveWidget());
    return bSet;
  };

  const detailsSection = CardService.newCardSection()
    .setHeader("Select Recipients")
    .addWidget(textWidget())
    .addWidget(getFilesAndFoldersDataWidget())
    .addWidget(buttonSet());

  return detailsSection;
}

/**
 * Calls the stored files and folders if the file picker is selected
 * otherwise returns a request to select docs and removes files from Property Service.
 *
 * Doc links are added with their title and url as a hyperlink in each page.
 *
 * @return {CardService.TextParagraph} text widget string either request to select docs or a list of doc links.
 */
function getFilesAndFoldersDataWidget() {
  let filePick = JSON.parse(
    PropertiesService.getUserProperties().getProperty("filePick")
  );
  let prop = PropertiesService.getUserProperties().getProperty("files");

  let paragraph = "";
  if (prop == null || !filePick) {
    paragraph = `<i>Use an existing List.</i>`;

    clearFilesFromPropServ();
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
      PropertiesService.getUserProperties();
      NEXT_BTN_TGL = false;
      CLEAR_BTN_TGL = false;
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
