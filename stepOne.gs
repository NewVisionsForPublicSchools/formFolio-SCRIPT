var scriptTitle = "formFolio Script V1.0.6 (9/15/13)";
var scriptName = 'formFolio';
var scriptTrackingId = 'UA-43639576-1';
var waitingIconId = '0B7-FEGXAo-DGalczbTk3UEtWdlk';
var waitingImageUrl = 'https://drive.google.com/uc?export=download&id='+this.waitingIconId;

// Written by Andrew Stillman for New Visions for Public Schools
// Published under GNU General Public License, version 3 (GPL-3.0)
// See restrictions at http://www.opensource.org/licenses/gpl-3.0.html


function onInstall() {
  onOpen();
}


function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var properties = ScriptProperties.getProperties();
  if (!ss) {
    ss = SpreadsheetApp.openById(properties.ssId);
  }
  var menuItems = [];
  menuItems[0] = {name: "What is formFolio?", functionName: "formFolio_whatIs"};
  menuItems[1] = null;
  if (!properties.installed) {
    menuItems.push({name: "Run initial installation", functionName: "formFolio_runInstallation"});
  }
  if (properties.installed=="true") {
    menuItems.push({name: "Step 1: Help formFolio understand your form", functionName: "formFolio_settingsUi"});
    if (properties.openStep2=="true") {
      menuItems.push({name: "Step 2: Configure, Create, Refresh Folders", functionName: "formFolio_folderKeys"}); 
      if (properties.openStep3=="true") {
        menuItems.push({name: "Step 3: Set run mode", functionName: "formFolio_setRunMode"}); 
        menuItems.push(null);
        menuItems.push({name: "Package this system for others to copy", functionName: "formFolio_extractorWindow"});  
      }
    }
  }
  ss.addMenu("formFolio", menuItems);
}


function formFolio_runInstallation() {
  formFolio_preconfig();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var properties = ScriptProperties.getProperties();
  if (!ss) {
    ss = SpreadsheetApp.openById(properties.ssId);
  }
  var domain = Session.getActiveUser().getEmail().split("@")[1];
  var formUrl = ss.getFormUrl();
  var msg = "";
  if (formUrl) {
    var form = FormApp.openByUrl(formUrl);
    if (!properties.urlQId) {
      msg = "The script discovered an existing form, added a text question for the Drive resource URL";
      formFolio_addDriveUrlQuestion(form);   
    }
    if (!properties.userNameQId) {
      if (domain!="gmail.com") {
        form.setRequireLogin(true);
        form.setCollectEmail(true);
        msg = ", set your form to auto-collect username";
      } else {
        formFolio_addEmailQuestion(form);
        msg = " and added a question to collect email.";
      }
    }
  } else {
    msg = "The script attached a new form to this Spreadsheet";
    var form = FormApp.create(ss.getName() + "- Form");
    var formId = form.getId();
    var ssFile = DriveApp.getFileById(ss.getId());
    var ssFileParents = ssFile.getParents();
    var formFile = DriveApp.getFileById(formId);
    while (ssFileParents.hasNext()) {
      var thisParent = ssFileParents.next();
      thisParent.addFile(formFile);
      DriveApp.getRootFolder().removeFile(formFile);
    }
    form.setTitle(ss.getName() + "- Form");
    if (domain!="gmail.com") {
      form.setRequireLogin(true);
      form.setCollectEmail(true);
      msg += "set the form to collect username, ";
    } else {
      formFolio_addEmailQuestion(form);
      msg += ", added a question to collect email";
    }
    var listItem = form.addListItem();
    listItem.setTitle("Select Folder").setHelpText("This question was added as a default and can be modified to suit your need...").setChoiceValues(['Folder 1','Folder 2', 'Folder 3']);
    listItem.setRequired(true);
    ScriptProperties.setProperty('formQidsSelected', listItem.getId().toString());
    msg += ", added a default list type question entitled \"Folder\"";
    formFolio_addDriveUrlQuestion(form);
    msg += " and added a text type question to enable users to submit Drive resource URLs."; 
    form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
    SpreadsheetApp.flush();
    var sheets = ss.getSheets();
    var formSheetId = sheets[sheets.length-1].getSheetId();
    ScriptProperties.setProperty('formSheetId', formSheetId);
  }
  ScriptProperties.setProperty('installed', 'true');
  Browser.msgBox("formFolio successfully initialized. " + msg);
  onOpen();
}


function formFolio_addDriveUrlQuestion(form) {
  var textQ = form.addTextItem();
  textQ.setTitle("Paste the URL of the Google Drive resource you are submitting");
  textQ.setHelpText("This Drive resource may be any file type or folder in Drive, but it must be shared (at least view-only) with " + Session.getEffectiveUser().getEmail() + " in order for this form to place it in the correct folder(s).");
  textQ.setRequired(true);
  ScriptProperties.setProperty('urlQId', textQ.getId().toString());
  return;
}

function formFolio_addEmailQuestion(form) {
  var textQ = form.addTextItem();
  textQ.setTitle("Please provide your Google email address");
  textQ.setHelpText("This will be used to notify you in case this form is unable to access your Drive submission.");
  textQ.setRequired(true);
  ScriptProperties.setProperty('userNameId', textQ.getId().toString());
  return;
}


function formFolio_settingsUi() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var properties = ScriptProperties.getProperties();
  if (!ss) {
    ss = SpreadsheetApp.openById(properties.ssId);
  }
  var formUrl = ss.getFormUrl();
  if (!formUrl) {
    Browser.msgBox("This script requires a spreadsheet that already has a Google Form attached to it. Please come back to this step once you have a form.");
    return;
  }
  var form = FormApp.openByUrl(formUrl);
  var app = UiApp.createApplication().setTitle("Step 1: Help formFolio understand your form...");
  var panel = app.createVerticalPanel().setId('panel');
  var outerScrollPanel = app.createScrollPanel().setHeight("250px").setStyleAttribute('backgroundColor', 'whiteSmoke');
  var innerpanel = app.createVerticalPanel().setId('innerpanel').setStyleAttribute('margin', '10px');
  var formSheetLabel = app.createLabel("Select the sheet containing your form responses");
  var sheets = ss.getSheets();
  var formSheetSelect = app.createListBox().setName('formSheetId').setId('formSheetSelect');
  var sheetIds = [];
  for (var i=0; i<sheets.length; i++) {
    formSheetSelect.addItem(sheets[i].getName(), sheets[i].getSheetId());
    sheetIds.push(sheets[i].getSheetId());
  }
  if (properties.formSheetId) {
    var selectedId = parseInt(properties.formSheetId);
    var index = sheetIds.indexOf(selectedId);
    formSheetSelect.setSelectedIndex(index);
  }
  innerpanel.add(formSheetLabel).add(formSheetSelect);
  var formQsScrollPanel = app.createScrollPanel().setHeight("150px").setWidth("300px").setStyleAttribute('backgroundColor', 'whiteSmoke');
  var formQsInnerPanel = app.createVerticalPanel();
  var formQsLabel = app.createLabel("Select the form question(s) whose options will correspond to Drive folders (only multiple choice, list, checkbox questions, and username (for Apps domain users collecting username) should appear below)").setStyleAttribute('marginTop', '15px');
  var formItems = form.getItems();
  var numItems = 0;
  if (properties.formQidsSelected) {
    var selectedQids = properties.formQidsSelected.split(","); 
  }
  for (var i=0; i<formItems.length; i++) {
    if ((formItems[i].getType() == "MULTIPLE_CHOICE")||(formItems[i].getType() == "CHECKBOX")||(formItems[i].getType() == "LIST")) {
      numItems++;
      var thisCheckBox = app.createCheckBox(formItems[i].getTitle()).setName('checkBox-'+numItems);
      if (selectedQids) {
        if (selectedQids.indexOf(formItems[i].getId().toString()) != -1) {
          thisCheckBox.setValue(true);
        }
      }
      var thisHidden = app.createHidden('hidden-'+numItems, formItems[i].getId());
      formQsInnerPanel.add(thisCheckBox).add(thisHidden);
    }
  }
  var isUsernameCollected = form.collectsEmail();
  if (isUsernameCollected) {
    numItems++;
    var thisCheckBox = app.createCheckBox('Username').setName('checkBox-'+numItems);
    if (selectedQids) {
      if (selectedQids.indexOf('Username') != -1) {
        thisCheckBox.setValue(true);
      }
    }
    var thisHidden = app.createHidden('hidden-'+numItems, "Username");
    formQsInnerPanel.add(thisCheckBox).add(thisHidden);
  }
  if (numItems=="0") {
    Browser.msgBox("Your form must have at least one multiple choice, checkbox, or list item question, or be set to collect username in order to work with this script.");
    app.close;
    return app;
  }
  var numHidden = app.createHidden('numItems', numItems);
  innerpanel.add(formQsLabel).add(formQsInnerPanel).add(numHidden);
  
  var textQIds = [];
  for (var i=0; i<formItems.length; i++) {
    if (formItems[i].getType() == "TEXT") {
      textQIds.push(formItems[i].getId());
    }
  }
  if (form.collectsEmail() == false) {
    var usernameQsLabel = app.createLabel("Select the form question that collects the user's Google Email address. (only text questions will appear in the listbox below)").setStyleAttribute('marginTop', '15px');
    var usernameHelpLabel = app.createLabel("Why is this necessary? Because your form is not set to collect username, you must ask the user to self-report their Google email address.  This will be used to email them a notification and link to share in the event the Drive resource is not shared properly upon initial submission.").setStyleAttribute('fontSize', '10px').setStyleAttribute('color', 'grey');
    var usernameListbox = app.createListBox().setName('usernameQId').setStyleAttribute('width', '300px');
    for (var i=0; i<formItems.length; i++) {
      if (formItems[i].getType() == "TEXT") {  
        usernameListbox.addItem(formItems[i].getTitle(), formItems[i].getId());
      }
      if (properties.userNameId) {
        var index = textQIds.indexOf(parseInt(properties.userNameId));
        usernameListbox.setSelectedIndex(index);
      }
      innerpanel.add(usernameQsLabel).add(usernameHelpLabel).add(usernameListbox);
    } 
  }
  var urlQsLabel = app.createLabel("Select the form question that collects the URL of the submitted Drive resource. (only text questions will appear in the listbox below)").setStyleAttribute('marginTop', '15px');
  var urlQsListbox = app.createListBox().setName('urlQId').setStyleAttribute('width', '400px');
  for (var i=0; i<formItems.length; i++) {
    if (formItems[i].getType() == "TEXT") {
      urlQsListbox.addItem(formItems[i].getTitle(), formItems[i].getId());
    }
  }
  if (properties.urlQId) {
    var index = textQIds.indexOf(parseInt(properties.urlQId));
    urlQsListbox.setSelectedIndex(index);
  }
  var nonDriveUrlOption = app.createCheckBox("Email users warning when they try to submit a non-Drive URL").setName('nonDriveWarning').setValue(true);
  if (properties.nonDriveWarning == "false") {
    nonDriveUrlOption.setValue(false);
  }
  innerpanel.add(urlQsLabel).add(urlQsListbox);
  innerpanel.add(nonDriveUrlOption);
  outerScrollPanel.add(innerpanel);
  panel.add(outerScrollPanel);
  var saveHandler = app.createServerHandler('formFolio_saveSettings').addCallbackElement(panel);
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
  app.add(panel);
  app.add(waitingImage);
  ss.show(app);
  return app;
}


function formFolio_saveSettings(e) {
  var app = UiApp.getActiveApplication();
  var properties = ScriptProperties.getProperties();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    ss = SpreadsheetApp.openById(properties.ssId);
  }
  var formUrl = ss.getFormUrl();
  var form = FormApp.openByUrl(formUrl);
  properties.formSheetId = e.parameter.formSheetId;
  var numItems = parseInt(e.parameter.numItems);
  var formQidsSelected = [];
  var sheetIdMappings = properties.sheetIdMappings;
  if (sheetIdMappings) {
    sheetIdMappings = Utilities.jsonParse(sheetIdMappings);
  } else {
    sheetIdMappings = '';
  }
  for (var i=0; i<numItems; i++) {
    var isSelected = e.parameter['checkBox-'+(i+1)]
    if (isSelected == "true") {
      formQidsSelected.push(e.parameter['hidden-'+(i+1)]);
    } else if ((sheetIdMappings != "")&&(isSelected == "false")) {
      sheetIdMappings = removeSheetMapping(e.parameter['hidden-'+(i+1)], sheetIdMappings)
    }
  }
  if (sheetIdMappings!='') {
    properties.sheetIdMappings = Utilities.jsonStringify(sheetIdMappings);
  }
  properties.formQidsSelected = formQidsSelected.join(",");
  if (form.collectsEmail()) {
    properties.userNameId = "Username";
  } else {
    properties.userNameId = e.parameter.usernameQId;
  }
  properties.urlQId = e.parameter.urlQId;
  properties.nonDriveWarning = e.parameter.nonDriveWarning;
  ScriptProperties.setProperties(properties);
  var existingHelpText = form.getItemById(e.parameter.urlQId).asTextItem().getHelpText();
  if (existingHelpText.indexOf("This Drive resource")==-1) {
    var newHelpText = existingHelpText + " This Drive resource must be shared (at least view-only) with " + Session.getEffectiveUser().getEmail() + " in order for this form to place it in the correct folder(s).";
  }
  form.getItemById(e.parameter.urlQId).setHelpText(newHelpText);
  app.close();
  if (!properties.openStep2) {
    properties.openStep2 = "true";
    ScriptProperties.setProperties(properties);
    onOpen();
    formFolio_folderKeys();
  }
  return app;
}
