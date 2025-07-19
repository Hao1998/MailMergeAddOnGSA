
var card = CardService.newCardBuilder();

function onHomePage(e) {
  var section1 = CardService.newCardSection().setHeader('Select Document Type:');
  var radioButtonGroup = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.RADIO_BUTTON)
    .setFieldName('documentType')
    .addItem('Letters', 'letters', true)
   // .addItem('E-mail Messages', 'emailMessages', false)
    .setOnChangeAction(CardService.newAction().setFunctionName('updateEmailText'))

  section1.addWidget(radioButtonGroup);
  card.addSection(section1);

  var section3 = CardService.newCardSection().setHeader('Letters');
  const emailText = CardService.newTextParagraph()
    .setText("Send letters to a group of people. You can personalize the letter that each person receives.")

  section3.addWidget(emailText)
  card.addSection(section3)

  var section2 = CardService.newCardSection()
  //.setHeader('Step 1 of 6');
  var nextButton = CardService.newTextButton().setText("Next").setOnClickAction(
    CardService.newAction().setFunctionName('startingDocument')
  )
  section2.addWidget(nextButton);
  card.addSection(section2);

  return CardService.newNavigation().pushCard(card.build());

}


function updateEmailText(e) {
  var documentType = e && e.formInput && e.formInput.documentType;

  if (documentType === 'letters') {
    var section1 = CardService.newCardSection().setHeader('Select Document Type:');
    var radioButtonGroup = CardService.newSelectionInput()
      .setType(CardService.SelectionInputType.RADIO_BUTTON)
      .setFieldName('documentType')
      .addItem('Letters', 'letters', true)
      .addItem('E-mail Messages', 'emailMessages', false)
      .setOnChangeAction(CardService.newAction().setFunctionName('updateEmailText'))

    section1.addWidget(radioButtonGroup);
    card.addSection(section1);

    var section3 = CardService.newCardSection().setHeader('E-mail Merges');
    const emailText = CardService.newTextParagraph()
      .setText("Send letters to a group of people. You can personalize the letter that each person receives. Click Next to continue")

    section3.addWidget(emailText)
    card.addSection(section3)

    var section2 = CardService.newCardSection().setHeader('Step 1 of 6');
    var nextButton = CardService.newTextButton().setText("Next").setOnClickAction(
      CardService.newAction().setFunctionName('updateStartingDocText')
      //.setFunctionName('startingDocument')
    )
    section2.addWidget(nextButton);
    card.addSection(section2);

    return CardService.newNavigation().updateCard(card.build())
  } else if (documentType === 'emailMessages') {
    var section1 = CardService.newCardSection().setHeader('Select Document Type:');
    var radioButtonGroup = CardService.newSelectionInput()
      .setType(CardService.SelectionInputType.RADIO_BUTTON)
      .setFieldName('documentType')
      .addItem('Letters', 'letters', false)
      .addItem('E-mail Messages', 'emailMessages', true)
      .setOnChangeAction(CardService.newAction().setFunctionName('updateEmailText'))

    section1.addWidget(radioButtonGroup);
    card.addSection(section1);
   // var section2 = CardService.newCardSection().setHeader('Step 1 of 6');
    var nextButton = CardService.newTextButton().setText("Next").setOnClickAction(
      CardService.newAction().setFunctionName('updateStartingDocText')
      //.setFunctionName('startingDocument')
    )
    section2.addWidget(nextButton);
   // card.addSection(section2);

    var section3 = CardService.newCardSection().setHeader('Letters');
    const emailText = CardService.newTextParagraph()
      .setText("Send letters to a group of people. You can personalize the letter that each person receives.")

    section3.addWidget(emailText)
    card.addSection(section3)

    return CardService.newNavigation().updateCard(card.build());
  }

}

