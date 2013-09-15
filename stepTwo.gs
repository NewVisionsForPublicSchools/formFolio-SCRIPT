function formFolio_folderKeys() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var formUrl = ss.getFormUrl();
  var form = FormApp.openByUrl(formUrl);
  var properties = ScriptProperties.getProperties();
  if (!ss) {
    ss = SpreadsheetApp.openById(properties.ssId);
  }
  var app = UiApp.createApplication().setTitle("Step 2: Configure, Create, Refresh Folders").setHeight(340);
  var panel = app.createVerticalPanel();
  var helpLabel = app.createLabel("In Step 1 you indicated that the following form item(s) will serve as Drive folder selectors. See multiple items listed below? This means each Drive resource submission will be added to multiple folders.");
  var innerPanel = app.createHorizontalPanel().setStyleAttribute('margin', '3px');
  var selectedQs = properties.formQidsSelected;
  if (selectedQs) {
    selectedQs = selectedQs.split(",");
    for (var i=0; i<selectedQs.length; i++) {
      var labelText = '';
      if (selectedQs[i]=="Username") {
        labelText = "Username";
      } else {
        labelText = form.getItemById(selectedQs[i]).getTitle();
      }
      var label = app.createLabel(labelText);
      label.setStyleAttribute('margin', '3px').setStyleAttribute('padding','5px').setStyleAttribute('backgroundColor', '#FFFFCC').setStyleAttribute('fontSize', '14px');
      if (labelText == "Username") {
        label.setStyleAttribute('backgroundColor','#FFA07A');
      }
      innerPanel.add(label);
    }
  } else {
    Browser.msgBox("You must select at least one form item to serve as a Drive folder selector. Please return to Step 1 to correct this...");
    app.close;
    return app;
  }
  panel.add(helpLabel);
  panel.add(innerPanel);
  if (((selectedQs.indexOf("Username")!=-1)&&(selectedQs.length>1))||((selectedQs.indexOf("Username")==-1)&&(selectedQs.length>0))) {
    var generalHeader = app.createLabel("General folder options").setStyleAttribute('width','100%').setStyleAttribute('color','black').setStyleAttribute('backgroundColor','#FFFFCC').setStyleAttribute('padding','3px').setStyleAttribute('margin', '3px').setStyleAttribute('fontSize', '14px');
    var folderModeLabel = app.createLabel("How do you want to map question options to Drive folders?").setStyleAttribute('margin', '3px');
    var folderModeBox = app.createListBox().setName("folderMode").setStyleAttribute('margin', '3px');
    folderModeBox.addItem("Automatically create new Drive folders for all question options in my form","0");
    folderModeBox.addItem("I have existing Drive folders that I will manually map to my question options", "1");
    if (properties.folderMode) {
      var index = parseInt(properties.folderMode);
      folderModeBox.setSelectedIndex(index);
    }
    panel.add(generalHeader).add(folderModeLabel).add(folderModeBox);
  }
  if (selectedQs.indexOf("Username")!=-1) {
    var usernameHeader = app.createLabel("User folder options").setStyleAttribute('width','100%').setStyleAttribute('color','black').setStyleAttribute('backgroundColor','#FFA07A').setStyleAttribute('padding','3px').setStyleAttribute('margin', '3px').setStyleAttribute('fontSize', '14px');
    var userAccessLabel = app.createLabel("Select the level of access you want users to have to their own \"User Folder\"").setStyleAttribute('margin', '3px');
    var userAccessBox = app.createListBox().setName("userFolderAccess").setStyleAttribute('margin', '3px');
    userAccessBox.addItem('No access', '0');
    userAccessBox.addItem('View own folder', '1');
    userAccessBox.addItem('Edit own folder', '2');
    if (properties.userFolderAccess) {
      var index = parseInt(properties.userFolderAccess);
      userAccessBox.setSelectedIndex(index);
    }
    var usernameModeLabel = app.createLabel("How do you want user folders managed?").setStyleAttribute('margin', '3px');
    var usernameModeBox = app.createListBox().setName("userFolderMode").setStyleAttribute('margin', '3px');
    usernameModeBox.addItem("Automatically create user folders as new users submit to the form","0");
    usernameModeBox.addItem("Do not create folders for users not listed in \"User Folders\" sheet", "1");
    if (properties.userFolderMode) {
      var index = parseInt(properties.userFolderMode);
      usernameModeBox.setSelectedIndex(index);
    }
    panel.add(usernameHeader).add(userAccessLabel).add(userAccessBox).add(usernameModeLabel).add(usernameModeBox);
  }
  // var helpLabel2 = app.createLabel("Clicking the button below will create/update tabs in your spreadsheet for each question. You will use these tabs to maintain folder key mappings...").setStyleAttribute('marginTop', '10px');
 // var helpLabel3 = app.createLabel("What is a folder key? It's the long, unique string found in the URL while looking at any Drive folder in your browser. See image below:").setStyleAttribute("color","grey").setStyleAttribute("fontSize","10px").setStyleAttribute('marginTop', '5px');
 // var imageId = '0B2vrNcqyzernQklCd0lsQm5iNlk';
 // var url = 'https://drive.google.com/uc?export=download&id='+imageId;
 // var helpImage = app.createImage(url).setWidth('500px');
 // panel.add(helpLabel2).add(helpLabel3).add(helpImage);
  var tabRefreshHandler = app.createServerHandler('formFolio_refreshMappingTabs').addCallbackElement(panel);
  var button = app.createButton('Save Settings and Create / Refresh Form Folders', tabRefreshHandler).setStyleAttribute('marginTop','8px');
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


function formFolio_refreshMappingTabs(e) {
  var app = UiApp.getActiveApplication();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var properties = ScriptProperties.getProperties();
  if (!ss) {
    ss = SpreadsheetApp.openById(properties.ssId);
  }
  var formUrl = ss.getFormUrl();
  var form = FormApp.openByUrl(formUrl);
  var sheets = ss.getSheets();
  var sheetNames = [];
  var message = '';
  if (properties.sheetIdMappings) {
    var sheetIdMappings = Utilities.jsonParse(properties.sheetIdMappings);
  } else {
    sheetIdMappings = [];
  }
  var selectedQs = properties.formQidsSelected;
  if (selectedQs) {
    selectedQs = selectedQs.split(",");
  }
  if (selectedQs.indexOf("Username")!=-1) {
    var userFolderMode = e.parameter.userFolderMode;
    var userFolderAccess = e.parameter.userFolderAccess;
    ScriptProperties.setProperty('userFolderMode', userFolderMode);
    ScriptProperties.setProperty('userFolderAccess', userFolderAccess);
  }
  if (((selectedQs.indexOf("Username")!=-1)&&(selectedQs.length>1))||((selectedQs.indexOf("Username")==-1)&&(selectedQs.length>0))) {
    if (e) {
      if (e.parameter) {
        var folderMode = e.parameter.folderMode;
        if (folderMode) {
          ScriptProperties.setProperty('folderMode', folderMode);
        } else {
          folderMode = properties.folderMode;
        }
      } else {
        folderMode = properties.folderMode;
      }
    } else {
      folderMode = properties.folderMode;
    }
  } 
  for (var i=0; i<selectedQs.length; i++) {
    var mappedSheet = returnMappedSheet(ss, selectedQs[i], sheetIdMappings);
    if (mappedSheet == "not found") {
        sheetIdMappings = removeSheetMapping(selectedQs[i], sheetIdMappings);
      }
    if (selectedQs[i]!="Username") {
      var thisQ = form.getItemById(parseInt(selectedQs[i]));
      var theseOptions = formFolio_getOptionsList(thisQ);
      if ((!mappedSheet)||(mappedSheet=="not found")) {
        try {
          var newSheet = ss.insertSheet("folderKeyMappings - " + thisQ.getTitle());
        } catch(err) {
          newSheet = ss.getSheetByName("folderKeyMappings - " + thisQ.getTitle());
        }
        newSheet.getRange(1, 1, 1, 2).setValues([["Question Item","Folder Key"]]);
        newSheet.getRange(2, 1, theseOptions.length, 1).setValues(theseOptions);
        newSheet.setFrozenRows(1);
        sheetIdMappings.push({sheetId: newSheet.getSheetId(), formQId: selectedQs[i]})
      } else {
        var oldOptions = mappedSheet.getRange(2, 1, mappedSheet.getLastRow()-1, 1).getValues();
        var old1D = [];
        for (var j=0; j<oldOptions.length; j++) {
          old1D.push(oldOptions[j][0]);
        }
        for (var j=0; j<theseOptions.length; j++) {
          var thisOption = theseOptions[j][0];
          var alreadyExists = old1D.indexOf(thisOption);
          if (alreadyExists==-1) {
            mappedSheet.insertRowAfter(mappedSheet.getLastRow());
            mappedSheet.getRange(mappedSheet.getLastRow()+1, 1).setValue(theseOptions[j][0])
          }
        }
      }
    } else {  //this is the username option
      if (!mappedSheet) {
        var userSheet = formFolio_createUserFolderSheet();
        var results = formFolio_createRefreshUserFolders();
        sheetIdMappings.push({sheetId: userSheet.getSheetId(), formQId: selectedQs[i], rootFolderId: results.userRootFolderId})
      } else {
        var results = formFolio_createRefreshUserFolders();
      }
    }
  }
  sheetIdMappings = Utilities.jsonStringify(sheetIdMappings);
  ScriptProperties.setProperty('sheetIdMappings', sheetIdMappings);
  if (folderMode=="1") {
    message += "Before moving on to step 3, you will want to take some time to populate the folder keys that corresponds to each question choice in the folder mapping tab(s).";
  }
  if (folderMode=="0") {
    var generalResults = formFolio_createRefreshGeneralFolders();
    if ((generalResults.newCount + generalResults.errorCount)>0) {
      message += " " + generalResults.newCount +  " new General folder(s) created, with " + generalResults.errorCount + " errors.";
    }
  }
  ScriptProperties.setProperty('openStep3','true');
  if (results) {
    if ((results.newCount + results.errorCount)>0) {
      message += " " + results.newCount +  " new User folder(s) created, with " + results.errorCount + " errors.";
    }
  }
  onOpen();
  if (app) {
    if (message!='') {
      try {
        Browser.msgBox(message);
      } catch(err) {
      }
    }
    app.close();
    return app;
  }
}

function removeSheetMapping(selectedQId, sheetIdMappings) {
  var index = sheetIdMappings.indexOf(selectedQId); 
  sheetIdMappings.splice(index, 1); 
  return sheetIdMappings;
}


function returnMappedSheet(ss, selectedQId, sheetIdMappings) {
  for (var i=0; i<sheetIdMappings.length; i++) {
    if (sheetIdMappings[i].formQId == selectedQId) {
      var sheetId = sheetIdMappings[i].sheetId;
      var sheets = ss.getSheets();
      var found = false;
      for (var j=0; j<sheets.length; j++) {
        if (sheets[j].getSheetId() == sheetId) {
          found = true;
          var foundSheet = sheets[j];
        }
      }
      if (found == true) {
        return foundSheet;
      } else {
        return "not found";
      }
    }
  }
  return;
}

function formFolio_addUserToFolderSheet(emailAddress) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('User Folders');
  if (!sheet) {
    sheet = formFolio_createUserFolderSheet();
  }
  var lastRow = sheet.getLastRow();
  if (lastRow>1) {
    var dataRange = sheet.getRange(2, 1, lastRow-1, sheet.getLastColumn());
    var data = getRowsData(sheet, dataRange);
    var formulas = getRowsFormulas(sheet, dataRange);
  } else {
    var data = [];
    var formulas = [];
  }
  if (emailAddress) {
    var found = false;
    for (var i=0; i<data.length; i++) {
      if ((data[i].emailAddress)&&(data[i].emailAddress!='')) {
        if (data[i].emailAddress == emailAddress) {
          found = true;
          break;
        }  
      }
    }
    if (found == false) {
      var thisData = new Object();
      thisData.username = emailAddress;
      thisData.folderTitle = emailAddress;
      data.push(thisData);
      setRowsData(sheet, data);
      setRowsFormulas(sheet, formulas, '', 2, 1);
    } else {
      return sheet;
    }
  } else {
    return sheet;
  }
}


function formFolio_createRefreshGeneralFolders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var properties = ScriptProperties.getProperties();
  if (!ss) {
    ss = SpreadsheetApp.openById(properties.ssId);
  }
  var formUrl = ss.getFormUrl();
  var form = FormApp.openByUrl(formUrl);
  var sheetIdMappings = properties.sheetIdMappings;
  var returnObj = new Object();
  returnObj.newCount = 0;
  returnObj.errorCount = 0;
  if (sheetIdMappings) {
    sheetIdMappings = Utilities.jsonParse(sheetIdMappings);
    for (var i=0; i<sheetIdMappings.length; i++) {
      var sheet = returnMappedSheet(ss, sheetIdMappings[i].formQId, sheetIdMappings);
      var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      var folderKeyCol = headers.indexOf("Folder Key")+1;
      if (folderKeyCol == 0) {
        Browser.msgBox("Error: You are missing a \"Folder Key\" column in " + sheet.getName());
      }
      if ((!sheet)||(sheet=="not found")) {
        Browser.msgBox("Error: One of your folder sheets has been deleted.  Return to step 2.");
      }
      var lastRow = sheet.getLastRow();
      if (lastRow>1) {
        var dataRange = sheet.getRange(2, 1, lastRow-1, sheet.getLastColumn());
        var data = getRowsData(sheet, dataRange);
      } else {
        continue;
      }
      var rootFolder = null;
      if (sheetIdMappings[i].rootFolderId) {
         rootFolder = DriveApp.getFolderById(sheetIdMappings[i].rootFolderId);
      } 
      if ((!rootFolder)||(rootFolder.isTrashed())) {
        var parentFolder = DriveApp.getFileById(ss.getId()).getParents();
        if (parentFolder.hasNext()) {
          parentFolder = parentFolder.next();
        } else {
          parentFolder = DriveApp.getRootFolder();
        }
        rootFolder = parentFolder.createFolder('formFolio ' + form.getItemById(sheetIdMappings[i].formQId).getTitle() + " Folders");
        sheetIdMappings[i].rootFolderId = rootFolder.getId();
      }
      for (var j=0; j<data.length; j++) {
        if ((data[j].questionItem)&&(!data[j].folderKey)) {
          var folderTitle = data[j].questionItem;
          try {
            var newFolder =  rootFolder.createFolder(folderTitle);
            sheet.getRange(j+2, folderKeyCol).setValue(newFolder.getId());
            returnObj.newCount++;
          } catch (err) {
            returnObj.errorCount++;
          }
        }
      }
    }
    sheetIdMappings = Utilities.jsonStringify(sheetIdMappings);
    ScriptProperties.setProperty('sheetIdMappings', sheetIdMappings);
  }
  return returnObj;
}





function formFolio_createRefreshUserFolders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var properties = ScriptProperties.getProperties();
  if (!ss) {
    ss = SpreadsheetApp.openById(properties.ssId);
  }
  var sheet = ss.getSheetByName('User Folders');
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var folderKeyCol = headers.indexOf("Folder Key")+1;
  var returnObj = new Object();
  returnObj.newCount = 0;
  returnObj.errorCount = 0;
  if (folderKeyCol == 0) {
    Browser.msgBox("Error: You are missing a \"Folder Key\" column in your \"User Folders\" sheet");
  }
  if (!sheet) {
    sheet = formFolio_createUserFolderSheet();
  }
  var lastRow = sheet.getLastRow();
  if (lastRow>1) {
    var dataRange = sheet.getRange(2, 1, lastRow-1, sheet.getLastColumn());
    var data = getRowsData(sheet, dataRange);
  } else {
    return returnObj;
  }
  if (properties.userRootFolderId) {
    var userRootFolder = DriveApp.getFolderById(properties.userRootFolderId);
  }
   
  if ((!userRootFolder)||(userRootFolder.isTrashed())) {
    var parentFolder = DriveApp.getFileById(ss.getId()).getParents();
    if (parentFolder.hasNext()) {
      parentFolder = parentFolder.next();
    } else {
      parentFolder = DriveApp.getRootFolder();
    }
    userRootFolder = parentFolder.createFolder('formFolio User Folders');
    ScriptProperties.setProperty('userRootFolderId', userRootFolder.getId());
  }
  for (var i=0; i<data.length; i++) {
    if ((data[i].username)&&(!data[i].folderKey)) {
      var folderTitle = data[i].folderTitle;
      if (!folderTitle) {
        folderTitle = data[i].username;
      }
      try {
        var newFolder = formFolio_createUserFolder(data[i].username, folderTitle, userRootFolder, properties);
        sheet.getRange(i+2, folderKeyCol).setValue(newFolder.getId());
        returnObj.newCount++;
      } catch (err) {
        returnObj.errorCount++;
      }
    }
  }
  returnObj.userRootFolderId = userRootFolder.getId();
  return returnObj;
}


function formFolio_createUserFolderSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var properties = ScriptProperties.getProperties();
  if (!ss) {
    ss = SpreadsheetApp.openById(properties.ssId);
  }
  try {
    var sheet = ss.insertSheet("User Folders");
    var headerRange = sheet.getRange(1, 1, 1, 5);
    headerRange.setValues([["Username","First Name","Last Name", "Folder Title","Folder Key"]]);
    headerRange.setNotes([['Full email address, must include @myappsdomain.com','Optional','Optional','Can be customized via a spreadsheet formula. Default formula uses username or \"First Name Last Name\" if available','Generated for you by the script']]);
    sheet.getRange(2, 4).setFormula('=if(and(B2="",C2=""),A2,CONCATENATE(B2," ",C2))');
    sheet.setFrozenRows(1);
  } catch(err) {
    sheet = ss.getSheetByName('User Folders');
  }
  return sheet;
}



function formFolio_createUserFolder(emailAddress, folderTitle, parentFolder, properties) {
  var createdFolder = parentFolder.createFolder(folderTitle);
  if (properties.userFolderMode == "0") {
    return createdFolder;
  }
  if (properties.userFolderMode == "1") {
    createdFolder.addViewer(emailAddress);
  }
  if (properties.userFolderMode == "2") {
    createdFolder.addEditor(emailAddress);
  }
  return createdFolder;
}


function formFolio_getOptionsList(formItem) {
  var type = formItem.getType();
  var options = [];
  if (type == "MULTIPLE_CHOICE") {
    var choices = formItem.asMultipleChoiceItem().getChoices();
  }
  if (type == "CHECKBOX") {
    var choices = formItem.asCheckboxItem().getChoices()
    }
  if (type == "LIST") {
    var choices = formItem.asListItem().getChoices();
  }
  for (var i=0; i<choices.length; i++) {
    options.push([choices[i].getValue()]);
  }
  return options;
}
