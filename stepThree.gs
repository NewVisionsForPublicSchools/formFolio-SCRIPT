function formFolio_setRunMode() {
  var properties = ScriptProperties.getProperties();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var formUrl = ss.getFormUrl();
  var form = FormApp.openByUrl(formUrl);
  var formSheetId = properties.formSheetId
  var sheet = getSheetById(ss, formSheetId);
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var app = UiApp.createApplication().setTitle("Step 3: Set run mode").setHeight(400);
  var outerScrollPanel = app.createScrollPanel().setHeight('390px');
  var panel = app.createVerticalPanel();
  var label = app.createLabel("Indicate how you want to handle submitted Drive resources...");
  var listBox = app.createListBox().setName("mode");
  listBox.addItem("use original Drive resource in auto-foldering", "original");
  listBox.addItem("make copies of submitted Drive resources and auto-folder the copies", "copy"); 
  panel.add(label).add(listBox);
  var descriptionField = app.createLabel("").setId("modeDescription").setStyleAttribute('margin', '15px').setStyleAttribute('padding', '5px');
  panel.add(descriptionField);
  var copyModePanel = app.createVerticalPanel().setStyleAttribute('marginTop', '10px').setId('copyModePanel').setVisible(false);
  var copyShareLabel = app.createLabel("How do you want the submitter shared on the copied resource?")
  var copyShareListBox = app.createListBox().setName("copyShareMode");
  copyShareListBox.addItem("not at all", "none");
  copyShareListBox.addItem("as an editor", "edit");
  copyShareListBox.addItem("as a viewer", "view");
  copyShareListBox.addItem("as a commenter", "comment");
  if (properties.copyShareMode == "none") {
    copyShareListBox.setSelectedIndex(0);
  }
  if (properties.copyShareMode == "edit") {
    copyShareListBox.setSelectedIndex(1);
  }
  if (properties.copyShareMode == "view") {
    copyShareListBox.setSelectedIndex(2);
  }
  if (properties.copyShareMode == "comment") {
    copyShareListBox.setSelectedIndex(3);
  }
  copyModePanel.add(copyShareLabel).add(copyShareListBox);
  var copyFileNamingLabel = app.createLabel("Set the naming convention for the copied file").setStyleAttribute('marginTop', '10px');
  var copyFileNamingBox = app.createTextBox().setName('copyName').setWidth("100%");
  if ((properties.copyName)&&(properties.copyName!="")) {
    copyFileNamingBox.setValue(properties.copyName);
  } else {
    copyFileNamingBox.setValue("%originalFileName");
  }
  var variables = normalizeHeaders(headers);
  for (var i=0; i<variables.length; i++) {
    variables[i] = "$" + variables[i];
  }
  variables.push("<strong>%originalFileName</strong>");
  variables = variables.join(", ");
  var variablesLabel = app.createHTML("Available variables: " + variables).setStyleAttribute('fontSize','11px').setStyleAttribute('color', '#363636');
  
  var descFieldLabel = app.createLabel("Optional: Choose a form question field that will correspond to the \"File Description\" in Drive.  FYI: Descriptions are searchable and show up in the \"Details\" view for a Drive file or folder.").setStyleAttribute('marginTop', '10px');
  var descFieldList = app.createListBox().setName("descQId");
  var formQs = form.getItems();
  var unAllowedTypes = ['PAGE_BREAK','TIME','IMAGE','SECTION_HEADER','DATE','DATETIME','TIME','DURATION','GRID'];
  descFieldList.addItem("None","");
  var formQIds = [];
  for (var i=0; i<formQs.length; i++) {
    var thisType = formQs[i].getType().toString();
    if (unAllowedTypes.indexOf(thisType)==-1) {
      var thisFormQId = formQs[i].getId();
      descFieldList.addItem(formQs[i].getTitle(), thisFormQId);
      formQIds.push(thisFormQId);
    }
  }
  if ((properties.descQId)&&(properties.descQId!="")) {
    var index = formQIds.indexOf(parseInt(properties.descQId));
    descFieldList.setSelectedIndex(index+1);
  }
  
  
  var collabFieldLabel = app.createLabel("Optional: Choose a form question that will correspond to additional collaborator email addresses (comma separated) to be shared on the copied file with the same settings selected above. Must be a text or paragraph question.").setStyleAttribute('marginTop', '10px');
  var collabFieldList = app.createListBox().setName("collabQId");
  var allowedTypes = ['TEXT','PARAGRAPH_TEXT'];
  collabFieldList.addItem("None","");
  var formQIds = [];
  for (var i=0; i<formQs.length; i++) {
    var thisType = formQs[i].getType().toString();
    var testFormQId = formQs[i].getId();
    if ((allowedTypes.indexOf(thisType)!=-1)&&(testFormQId!=properties.urlQId)&&(testFormQId!=properties.userNameId)) {
      var thisFormQId = formQs[i].getId();
      collabFieldList.addItem(formQs[i].getTitle(), thisFormQId);
      formQIds.push(thisFormQId);
    }
  }
  if ((properties.descQId)&&(properties.descQId!="")) {
    var index = formQIds.indexOf(parseInt(properties.collabQId));
    collabFieldList.setSelectedIndex(index+1);
  }
  copyModePanel.add(copyShareLabel).add(copyShareListBox).add(copyFileNamingLabel).add(copyFileNamingBox).add(variablesLabel);
  copyModePanel.add(descFieldLabel).add(descFieldList);
  copyModePanel.add(collabFieldLabel).add(collabFieldList);
  panel.add(copyModePanel);
  copyModePanel.setVisible(false);
  var listBoxChangeHandler = app.createServerHandler("formFolio_modeChangeRefresh").addCallbackElement(panel);
  listBox.addChangeHandler(listBoxChangeHandler);
  app.add(panel);
  if (properties.mode == "copy"){
    copyModePanel.setVisible(true);
    listBox.setSelectedIndex(1);
    var e = new Object();
    e.parameter = new Object;
    e.parameter.mode = "copy";
    formFolio_modeChangeRefresh(e);
  } else {
    listBox.setSelectedIndex(0);
    var e = new Object();
    e.parameter = new Object;
    e.parameter.mode = "original";
    formFolio_modeChangeRefresh(e)
  } 
  var saveHandler = app.createServerHandler('formFolio_saveMode').addCallbackElement(panel);
  var button = app.createButton("Save settings", saveHandler);
  var waitingImage = app.createImage(this.waitingImageUrl);
  waitingImage.setStyleAttribute('position','absolute')
  .setStyleAttribute('left','35%')
  .setStyleAttribute('top','20%')
  .setStyleAttribute('width','150px')
  .setStyleAttribute('height','150px')
  .setVisible(false);  
  var waitingHandler = app.createClientHandler().forTargets(waitingImage).setVisible(true).forTargets(panel).setStyleAttribute('opacity', '0.5').forTargets(button).setEnabled(false);
  button.addClickHandler(waitingHandler);
  panel.add(button);
  outerScrollPanel.add(panel);
  app.add(outerScrollPanel);
  app.add(waitingImage);
  ss.show(app);
  return app;
}


function formFolio_saveMode(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var app = UiApp.getActiveApplication();
  var properties = ScriptProperties.getProperties();
  properties.mode = e.parameter.mode;
  if (properties.mode == "copy") {
    properties.copyShareMode = e.parameter.copyShareMode;
    properties.copyName = e.parameter.copyName;
    properties.descQId = e.parameter.descQId;
    properties.collabQId = e.parameter.collabQId;
    if ((properties.copyName == "")||(!properties.copyName)) {
      Browser.msgBox("You must enter a file naming convention...");
      ScriptProperties.setProperties(properties);
      formFolio_setRunMode();
    }
  } else {
    properties.copyShareMode = '';
    properties.copyName = '';
  }
  ScriptProperties.setProperties(properties);
  formFolio_setTriggers(properties.mode);
  if (!properties.allSaved) {
    formFolio_addFolderingStatusCol(ss, properties.formSheetId);
    setFormEditTrigger();
    Browser.msgBox("Congrats! You should now be ready to test a Drive resource submission via your form...");
    properties.allSaved = true;
    ScriptProperties.setProperties(properties);
  }
  app.close();
  return app;
}


function formFolio_setTriggers(mode) {
  var ssKey = SpreadsheetApp.getActiveSpreadsheet().getId();
  var triggers = ScriptApp.getProjectTriggers();
  if (mode == "copy") {
    var alreadySet = false;
    for (var i=0; i<triggers.length; i++) {
      if (triggers[i].getHandlerFunction() == "formFolio_copyToFolders") {
        alreadySet = true;
        break;
      }
      if (triggers[i].getHandlerFunction() == "formFolio_addToFolders") {
        ScriptApp.deleteTrigger(triggers[i]);
      }
    }
    if (!alreadySet) {
      ScriptApp.newTrigger('formFolio_copyToFolders').forSpreadsheet(ssKey).onFormSubmit().create();
    }
  }
  if (mode == "original") {
    var alreadySet = false;
    for (var i=0; i<triggers.length; i++) {
      if (triggers[i].getHandlerFunction() == "formFolio_addToFolders") {
        alreadySet = true;
        break;
      }
      if (triggers[i].getHandlerFunction() == "formFolio_copyToFolders") {
        ScriptApp.deleteTrigger(triggers[i]);
      }
    }
    if (!alreadySet) {
      ScriptApp.newTrigger('formFolio_addToFolders').forSpreadsheet(ssKey).onFormSubmit().create();
    }
  }
}


function formFolio_modeChangeRefresh(e) {
  var app = UiApp.getActiveApplication();
  var mode = e.parameter.mode;
  var descriptionField = app.getElementById('modeDescription');
  var copyModePanel = app.getElementById('copyModePanel');
  if (mode == "original") {
    copyModePanel.setVisible(false);
    descriptionField.setText("This mode will preserve resource authors as owners, and will leave the file names unchanged.").setStyleAttribute('backgroundColor', 'pink');
  } 
  if (mode == "copy") {
     copyModePanel.setVisible(true);
     descriptionField.setText("This mode will create a copy of the original resource, making you the owner.  This mode also allows you to automatically rename the resource and include the author as a collaborator.").setStyleAttribute('backgroundColor', '#d2f6c1');
  }
  return app;
}



function formFolio_addFolderingStatusCol(ss, sheetId) {
  var sheet = getSheetById(ss, sheetId);
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (headers.indexOf("Link to Resource")==-1) {
    sheet.insertColumnAfter(sheet.getLastColumn());
    sheet.getRange(1, sheet.getLastColumn()+1).setValue("Link to Resource").setFontColor("white").setFontWeight("bold").setBackground('black');
  }
  if (headers.indexOf("Resource URL")==-1) {
    sheet.insertColumnAfter(sheet.getLastColumn());
    sheet.getRange(1, sheet.getLastColumn()+1).setValue("Resource URL").setFontColor("white").setFontWeight("bold").setBackground('black');
  }
  if (headers.indexOf("Resource Type")==-1) {
    sheet.insertColumnAfter(sheet.getLastColumn());
    sheet.getRange(1, sheet.getLastColumn()+1).setValue("Resource Type").setFontColor("white").setFontWeight("bold").setBackground('black');
  }
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (headers.indexOf("Drive ID")==-1) {
    sheet.insertColumnAfter(sheet.getLastColumn());
    sheet.getRange(1, sheet.getLastColumn()+1).setValue("Drive ID").setFontColor("white").setFontWeight("bold").setBackground('black');
  }
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (headers.indexOf("Foldering Status")==-1) {
    sheet.insertColumnAfter(sheet.getLastColumn());
    sheet.getRange(1, sheet.getLastColumn()+1).setValue("Foldering Status").setFontColor("white").setFontWeight("bold").setBackground('black');
  }
  return;
}
