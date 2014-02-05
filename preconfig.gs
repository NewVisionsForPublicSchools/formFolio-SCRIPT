function formFolio_preconfig() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ssId = ss.getId();
  ScriptProperties.setProperty('ssId', ssId);
  // if you are interested in sharing your complete workflow system for others to copy (with script settings)
  // Select the "Generate preconfig()" option in the menu and
  //#######Paste preconfiguration code below before sharing your system for copy#######
 

  
  
  //#######End preconfiguration code#######
  //remember to clear out all folder keys in folder key sheets if you are making a copy of this system for others.
  //Fetch system name, if this script is part of a New Visions system
  var systemName = NVSL.getSystemName();
  if (systemName) {
    ScriptProperties.setProperty('systemName', systemName)
  }
  //Fetch institutional tracking code.  If it exists, launch initialize function (autolaunch step 1 for repeat users)
  //If it doesn't exist, the checkInstitutionalTrackingCode() will launch the tracking settings UI.
  var institutionalTrackingString = NVSL.checkInstitutionalTrackingCode();
}


function formFolio_extractorWindow () {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var properties = ScriptProperties.getProperties();
  var sheetIdMappings = properties.sheetIdMappings
  if (sheetIdMappings) {
    sheetIdMappings = Utilities.jsonParse(sheetIdMappings);
    for (var i=0; i<sheetIdMappings.length; i++) {
      delete sheetIdMappings[i].rootFolderId;
    }
    properties.sheetIdMappings = Utilities.jsonStringify(sheetIdMappings);
  }
  var excludedProperties = ['formfolio_sid','installed','openStep2','openStep3','allSaved','ssId'];
  var propertyString = '';
  for (var key in properties) {
    if ((properties[key]!='')&&(excludedProperties.indexOf(key)==-1)) {
      var keyProperty = properties[key].replace(/[/\\*]/g, "\\\\");                                     
      propertyString += "   ScriptProperties.setProperty('" + key + "','" + keyProperty + "');\n";
    }
  }
  var app = UiApp.createApplication().setHeight(500).setTitle("Export preconfig() settings");
  var panel = app.createVerticalPanel().setWidth("100%").setHeight("100%");
  var labelText = "Copying a Google Spreadsheet copies scripts along with it, but without any of the script settings saved.  This normally makes it hard to share full, script-enabled Spreadsheet systems. ";
  labelText += " You can solve this problem by pasting the code below into a script function called \"formFolio_preconfig\" (go to formFolio in the Script Editor and select \"preconfig.gs\" in the left sidebar) prior to publishing your Spreadsheet for others to copy. \n";
  labelText += " After a user copies your spreadsheet, they will select \"Run initial installation.\"  This will preconfigure all needed script settings.  If you got this workflow from someone as a copy of a spreadsheet, this has probably already been done for you.";
  var label = app.createLabel(labelText);
  var window = app.createTextArea();
  var codeString = "//This section sets all script properties associated with this formFolio profile \n";
  codeString += "var preconfigStatus = ScriptProperties.getProperty('installed');\n";
  codeString += "if (preconfigStatus!='true') {\n";
  codeString += propertyString; 
  codeString += "};\n";
  codeString += "formFolio_setTriggers('" + properties.mode + "');\n";
  codeString += "ss.toast('Custom formFolio preconfiguration ran successfully. Please step through the formFolio menu options to confirm system settings.');\n";
  window.setText(codeString).setWidth("100%").setHeight("350px");
  app.add(label);
  panel.add(window);
  app.add(panel);
  ss.show(app);
  return app;
}
