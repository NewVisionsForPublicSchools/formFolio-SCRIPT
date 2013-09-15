function setFormEditTrigger() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var formUrl = ss.getFormUrl();
  var form = FormApp.openByUrl(formUrl);
  ScriptApp.newTrigger('formEditMenus').forForm(form).onOpen().create();
}

function formEditMenus() {
  var properties = ScriptProperties.getProperties();
  var ss = SpreadsheetApp.openById(properties.ssId);
  var formUrl = ss.getFormUrl();
  var form = FormApp.openByUrl(formUrl);
  var app = UiApp.createApplication().setTitle("Things to watch out for in editing this form...");
  var panel = app.createVerticalPanel();
  var html = "<ol>";
  if (properties.folderMode == "0") {
    html += "<li>If you add options to one of your form selection questions, don't forget to add corresponding folders using the \"formFolio\" menu to the left of \"Help\" above.</li>";
  }
  html += "<li>Because this form feeds a sheet with custom headers to the right of your form data, use the options below to insert new questions to avoid overwriting the custom headers in your destination sheet</li>";
  panel.add(app.createHTML(html));
  var numQsLabel = app.createLabel("Optional: Number of new questions to add to this form").setStyleAttribute('marginTop', '10px').setStyleAttribute('fontSize', '14px');
  var numQsSelect = app.createListBox().setName('numQsAdded');
  numQsSelect.addItem('0');
  numQsSelect.addItem('1');
  numQsSelect.addItem('2');
  numQsSelect.addItem('3');
  numQsSelect.addItem('4');
  panel.add(numQsLabel).add(numQsSelect);
  var OkHandler = app.createServerHandler('closeSaveFormUi');
  var addQsHandler = app.createServerHandler('insertNewFormCols').addCallbackElement(panel);
  var OkButton = app.createButton("Close", OkHandler);
  var addColsButton = app.createButton("Add number of new questions selected above", addQsHandler);
  var buttonPanel = app.createHorizontalPanel().setStyleAttribute('marginTop', '15px');
  var waitingImage = app.createImage(this.waitingImageUrl);
  waitingImage.setStyleAttribute('position','absolute')
  .setStyleAttribute('left','35%')
  .setStyleAttribute('top','20%')
  .setStyleAttribute('width','150px')
  .setStyleAttribute('height','150px')
  .setVisible(false);  
  var waitingHandler = app.createClientHandler().forTargets(waitingImage).setVisible(true).forTargets(panel).setStyleAttribute('opacity', '0.5').forTargets(OkButton).setEnabled(false).forTargets(addColsButton).setEnabled(false);
  OkButton.addClickHandler(waitingHandler);
  addColsButton.addClickHandler(waitingHandler);
  buttonPanel.add(OkButton).add(addColsButton);
  panel.add(buttonPanel);
  app.add(panel);
  app.add(waitingImage);
  FormApp.getUi().showDialog(app);
  var menu = FormApp.getUi().createMenu("formFolio");
  menu.addItem('Add question(s)', 'formEditMenus');
  if (properties.folderMode == "0") {
    menu.addItem('Refresh folders', 'formFolio_refreshMappingTabs');
  }
  menu.addToUi();
  return app;
}

function closeSaveFormUi() {
  var app = UiApp.getActiveApplication();
  app.close();
  return app;
}

function insertNewFormCols(e) {
  var app = UiApp.getActiveApplication();
  var properties = ScriptProperties.getProperties();
  var ss = SpreadsheetApp.openById(properties.ssId);
  var formUrl = ss.getFormUrl();
  var form = FormApp.openByUrl(formUrl);
  var formSheet = getSheetById(ss, properties.formSheetId);
  var headerColors = formSheet.getRange(1, 1, 1, formSheet.getLastColumn()).getBackgrounds()[0];
  var resourceTypeCol = headerColors.indexOf("#000000");
  if (e.parameter.numQsAdded) {
    var numQsAdded = parseInt(e.parameter.numQsAdded);
    if (numQsAdded > 0) {
      formSheet.insertColumnsAfter(resourceTypeCol, numQsAdded);
    }
    SpreadsheetApp.flush();
    for (var i=0; i<numQsAdded; i++) {
      form.addTextItem().setTitle("Untitled question" + (i+1));
    }
  }
  app.close();
  return app;
}
