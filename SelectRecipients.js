var card2 = CardService.newCardBuilder();
function startingDocument(e) {
  var documentType = e.formInput.documentType || e.parameters.documentType

  switch (documentType) {
    case 'letters':
      var section1 = CardService.newCardSection().setHeader('Select Starting document');
      var radioButtonGroup = CardService.newSelectionInput()
        .setType(CardService.SelectionInputType.RADIO_BUTTON)
        .setFieldName('startingDoc')
        .addItem('Use the current document', 'currentDoc', true)
        .addItem('Start from a template', 'fromTemplate', false)
      //  .addItem('Start from existing Document', 'fromExistingDoc', false)
        .setOnChangeAction(CardService.newAction().setFunctionName('updateStartingDocText'))
      section1.addWidget(radioButtonGroup);
      card2.addSection(section1);

      var section3 = CardService.newCardSection().setHeader('Use the current document');
      const emailText = CardService.newTextParagraph()
        .setText("Start from the document shown here and use the Mail Merge wizard to add recipient information")

      section3.addWidget(emailText)
      card2.addSection(section3)

      var section2 = CardService.newCardSection().setHeader('Step 2 of 6');
      var nextButton = CardService.newTextButton().setText("Next").setOnClickAction(
        CardService.newAction().setFunctionName('selectRecipients')
      )
      section2.addWidget(nextButton);
      card2.addSection(section2);



      return CardService.newNavigation().pushCard(card2.build());
      break;
  }
}


function updateStartingDocText(e) {
  var documentType = e && e.formInput && e.formInput.documentType;

  if (documentType === 'fromTemplate') {
     var section1 = CardService.newCardSection().setHeader('Select Starting document');
      var radioButtonGroup = CardService.newSelectionInput()
        .setType(CardService.SelectionInputType.RADIO_BUTTON)
        .setFieldName('startingDoc')
        .addItem('Use the current document', 'currentDoc', false)
        .addItem('Start from a teample', 'fromTemplate', true)
        .addItem('Start from existing Document', 'fromExistingDoc', false)
        .setOnChangeAction(CardService.newAction().setFunctionName('updateStartingDocText'))
      section1.addWidget(radioButtonGroup);
      card2.addSection(section1);

      var section3 = CardService.newCardSection().setHeader('Use the current document');
      const emailText = CardService.newTextParagraph()
        .setText("Start from a ready-to-use mail merge emplate that can be customized to suit your needs.")

      section3.addWidget(emailText)
      card2.addSection(section3)

      var section2 = CardService.newCardSection().setHeader('Step 2 of 6');
      var nextButton = CardService.newTextButton().setText("Next").setOnClickAction(
        CardService.newAction().setFunctionName('selectRecipients')
      )
      section2.addWidget(nextButton);
      card2.addSection(section2);

      return CardService.newNavigation().updateCard(card2.build());
  } else if (documentType === 'currentDoc'){
     var section1 = CardService.newCardSection().setHeader('Select Starting document');
      var radioButtonGroup = CardService.newSelectionInput()
        .setType(CardService.SelectionInputType.RADIO_BUTTON)
        .setFieldName('startingDoc')
        .addItem('Use the current document', 'currentDoc', true)
        .addItem('Start from a teample', 'fromTemplate', false)
        .addItem('Start from existing Document', 'fromExistingDoc', false)
        .setOnChangeAction(CardService.newAction().setFunctionName('updateStartingDocText'))
      section1.addWidget(radioButtonGroup);
      card2.addSection(section1);

      var section3 = CardService.newCardSection().setHeader('Use the current document');
      const emailText = CardService.newTextParagraph()
        .setText("Start from the document shown here and use the Mail Merge wizard to add recipient information")

      section3.addWidget(emailText)
      card2.addSection(section3)

      var section2 = CardService.newCardSection().setHeader('Step 2 of 6');
      var nextButton = CardService.newTextButton().setText("Next").setOnClickAction(
        CardService.newAction().setFunctionName('selectRecipients')
      )
      section2.addWidget(nextButton);
      card2.addSection(section2);



      return CardService.newNavigation().updateCard(card2.build());
  }
}

