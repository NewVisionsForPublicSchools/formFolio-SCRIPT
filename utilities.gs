function call(func, optLoggerFunction) {
  for (var n=0; n<6; n++) {
    try {
      return func();
    } catch(e) {
      if (optLoggerFunction) {optLoggerFunction("GASRetry " + n + ": " + e)}
      if (n == 5) {
        throw e;
      } 
      Utilities.sleep((Math.pow(2,n)*1000) + (Math.round(Math.random() * 1000)));
    }    
  }
}



// This function subs in row values for $variables
function replaceStringFields(headers, string, rowData, originalFileName) {
  var newString = string;
  var normalizedHeaders = normalizeHeaders(headers);
  var mergeTags = "$"+normalizedHeaders.join(",$");
  mergeTags = mergeTags.split(",");
  for (var i=0; i< mergeTags.length; i++) {
    var key = normalizedHeaders[i];
    var replacementValue = rowData[key];
    var replaceTag = mergeTags[i];
    replaceTag = replaceTag.replace("$","\\$") + "\\b";
    var find = new RegExp(replaceTag, "g");
    newString = newString.replace(find, replacementValue);
    newString = newString.replace(/(\r\n|\n|\r)/gm,"<br>");
  }
  newString = newString.replace("%originalFileName", originalFileName);
  return newString;
}


function getSheetById(ss, id) {
  var sheets = ss.getSheets();
  for (var i=0; i<sheets.length; i++) {
    if (sheets[i].getSheetId() == id) {
      return sheets[i];
    }
  }
  return;
}




function copyFolder(e) {
  var app = UiApp.getActiveApplication();
  var successLabel = app.getElementById('successLabel');
  var button = app.getElementById('button');
  successLabel.setVisible(false);
  successLabel.setHref('');
  var driveRoot = DocsList.getRootFolder();
  var rootFolderId = e.parameter.subFolder;
  var rootFolder = DocsList.getFolderById(rootFolderId);
  var rootName = rootFolder.getName();
  var copyRoot = DocsList.createFolder("Copy of " + rootName);
  var subFiles = rootFolder.getFiles();
  var subFolders = rootFolder.getFolders();
  for (var i=0; i<subFiles.length; i++) {
    var isSsConnectedForm = ssConnectedForm(subFiles[i]);
    if (!isSsConnectedForm) {
      var copy = subFiles[i].makeCopy(subFiles[i].getName());   
      copy.removeFromFolder(driveRoot);
      copy.addToFolder(copyRoot); 
      var connectedForm = returnConnectedForm(copy);
      if (connectedForm) {
        connectedForm.removeFromFolder(driveRoot);
        connectedForm.addToFolder(copyRoot);
      }
    }
  }
  for (var j=0; j<subFolders.length; j++) {
    var subFolderName = subFolders[j].getName();
    var subFolderCopy = DocsList.createFolder(subFolderName); 
    subFolderCopy.removeFromFolder(driveRoot);
    subFolderCopy.addToFolder(copyRoot);
    var files = subFolders[j].getFiles();
    for (var k=0; k<files.length; k++) {
      var isSsConnectedForm = ssConnectedForm(files[k]);
      if (!isSsConnectedForm) {  //only copy freestanding forms... ss connected forms will copy automatically
        var copy = files[k].makeCopy(files[k].getName());
        copy.removeFromFolder(driveRoot);
        copy.addToFolder(subFolderCopy);
        var connectedForm = returnConnectedForm(copy);
        if (connectedForm) {
          connectedForm.removeFromFolder(driveRoot);
          connectedForm.addToFolder(subFolderCopy);
        }
      }
    }
  }
  successLabel.setVisible(true).setHref(copyRoot.getUrl());
  var refreshPanel = app.getElementById('refreshPanel');
  var panel = app.getElementById('panel');
  refreshPanel.setVisible(false);
  panel.setStyleAttribute('opacity','1');
  app.close();
  return app;
}



function formFolio_copyFolder(rootFolder, folderCopyName) {
  var driveRoot = DocsList.getRootFolder();
  var rootName = rootFolder.getName();
  var copyRoot = DocsList.createFolder(folderCopyName);
  var subFiles = rootFolder.getFiles();
  var subFolders = rootFolder.getFolders();
  for (var i=0; i<subFiles.length; i++) {
    var isSsConnectedForm = ssConnectedForm(subFiles[i]);
    if (!isSsConnectedForm) {
      var copy = subFiles[i].makeCopy(subFiles[i].getName());   
      copy.removeFromFolder(driveRoot);
      copy.addToFolder(copyRoot); 
      var connectedForm = returnConnectedForm(copy);
      if (connectedForm) {
        connectedForm.removeFromFolder(driveRoot);
        connectedForm.addToFolder(copyRoot);
      }
    }
  }
  for (var j=0; j<subFolders.length; j++) {
    var subFolderName = subFolders[j].getName();
    var subFolderCopy = DocsList.createFolder(subFolderName); 
    subFolderCopy.removeFromFolder(driveRoot);
    subFolderCopy.addToFolder(copyRoot);
    var files = subFolders[j].getFiles();
    for (var k=0; k<files.length; k++) {
      var isSsConnectedForm = ssConnectedForm(files[k]);
      if (!isSsConnectedForm) {  //only copy freestanding forms... ss connected forms will copy automatically
        var copy = files[k].makeCopy(files[k].getName());
        copy.removeFromFolder(driveRoot);
        copy.addToFolder(subFolderCopy);
        var connectedForm = returnConnectedForm(copy);
        if (connectedForm) {
          connectedForm.removeFromFolder(driveRoot);
          connectedForm.addToFolder(subFolderCopy);
        }
      }
    }
  }
  return copyRoot;
}


function returnConnectedForm(file) {
  if ((file.getFileType().toString() == "spreadsheet")||(file.getFileType().toString() == "SPREADSHEET")) {
    var fileId = file.getId();
    var ss = SpreadsheetApp.openById(fileId);
    var formUrl = ss.getFormUrl();
    if (formUrl) {
      var form = FormApp.openByUrl(formUrl);
      var formId = form.getId();
      var docsListForm = DocsList.getFileById(formId);
      return docsListForm;
    } else {
      return false;
    }
  } else {
    return false;
  }
  
}



function ssConnectedForm(file) {
  var type = file.getFileType().toString();
  if ((type == "form")||(type == "FORM")) {
    var formId = file.getId();
    var form = FormApp.openById(formId);
    var destType = form.getDestinationType().toString();
    if ((destType == "SPREADSHEET")||(destType == "spreadsheet")) {
      return true;
    } else {
      return false;
    }
  } else {
    return false;
  }
}



function addCommentersToFolder(rootFolder, commenterEmail) {
  var commentersToAdd = [];
  commentersToAdd.push(commenterEmail);
  var subFiles = rootFolder.getFiles();
  var subFolders = rootFolder.getFolders();
  for (var i=0; i<subFiles.length; i++) {
    try {
      var driveFile = call(function() {return DriveApp.getFileById(subFiles[i].getId());});
      call(function() {driveFile.addCommenters(commentersToAdd);});
    } catch(err) {
      Logger.log(err.message);
    }
  }
  for (var j=0; j<subFolders.length; j++) {
    var files = subFolders[j].getFiles();
    for (var k=0; k<files.length; k++) {      
      try {
        var driveFile = call(function() {return DriveApp.getFileById(files[k].getId());});
        call(function() {driveFile.addCommenters(commentersToAdd);});
      } catch(err) {
        Logger.log(err.message);
        continue;
      }   
    }
    try {
      call(function() {subFolders[j].addViewers(commentersToAdd);});
    } catch(err) {
      Logger.log(err.message);
    } 
  }
  try {
    call(function() {rootFolder.addViewers(commentersToAdd);});
  } catch(err) {
    Logger.log(err.message);
  }
  return;
}



function formFolio_sendShareNotification(email, url, formResponseUrl) {
  try {
    var properties = ScriptProperties.getProperties();
    var userNameId = properties.userNameId;
    var urlQid = properties.urlQId;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sharelink = '';
    var scriptRunnerEmail = Session.getEffectiveUser().getEmail();
    if (url.indexOf("?")==-1) {
      sharelink = url + "?userstoinvite=" + scriptRunnerEmail;
    } else {
      sharelink = url + "&userstoinvite=" + scriptRunnerEmail;
    }
    var htmlBody = "Dear " + email + ",<br><br>";
    htmlBody += "You recently submitted a Drive resource to a form owned by " + scriptRunnerEmail + " that requires at least view privileges in order to submit to the correct Drive folders. <strong>To rectify, please follow both steps below:</strong><br>";
    htmlBody += '<br><strong>Step 1:</strong> Please use <a href="' + sharelink + '">this link</a> to add ' + scriptRunnerEmail + ' to the Drive resource.<br>';
    htmlBody += '<br><strong>Step 2:</strong> Resubmit your resource to this Google Form link:';
    //https://docs.google.com/forms/d/1eqsZDVYRplWbbscr0By9KuyfRJCGt87UPBVY_F-XxWE/viewform?entry.497590868=url&entry.345610109=email
    htmlBody += formResponseUrl;
    MailApp.sendEmail(email, "You submitted a Drive resource that wasn't shared properly...", "", {htmlBody: htmlBody});
    return;
  } catch(err) {
    Logger.log(err.message);
    return err.message;
  }
}



function formFolio_sendWrongResourceTypeNotification(email, url, formUrl) {
  try {
    var formTitle = FormApp.openByUrl(formUrl).getTitle();
    var properties = ScriptProperties.getProperties();
    var userNameId = properties.userNameId;
    var urlQid = properties.urlQId;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sharelink = '';
    var scriptRunnerEmail = Session.getEffectiveUser().getEmail();
    var htmlBody = "Dear " + email + ",<br><br>";
    htmlBody += 'You recently submitted an invalid URL to the form <a href = "' + formUrl + '">' + formTitle + '</a> owned by ' + scriptRunnerEmail + '.';
    htmlBody += '<br>Only Google Drive hosted resources may be submitted to the form.  Please try again after uploading your resource to Drive</br>';
    MailApp.sendEmail(email, "You submitted an invalid URL...", "", {htmlBody: htmlBody});
    return;
  } catch(err) {
    Logger.log(err.message);
    return err.message;
  }
}



function returnType(url, urlArgs) {
  var type = '';
  if (url.indexOf("folders")!=-1) {
    type = "FOLDER";
  }
  if (url.indexOf("folderview")!=-1) {
    type = "FOLDERVIEW";
  }
  if (urlArgs.indexOf("document")!=-1) {
    type = "DOCUMENT";
  }
  if (urlArgs.indexOf("spreadsheet")!=-1) {
    type = "SPREADSHEET";
  }
  if (urlArgs.indexOf("presentation")!=-1) {
    type = "PRESENTATION";
  }
  if (urlArgs.indexOf("drawings")!=-1) {
    type = "DRAWING";
  }
  if (urlArgs.indexOf("forms")!=-1) {
    type = "FORM";
  }
  if (urlArgs.indexOf("file")!=-1) {
    type = "OTHER";
  }
  return type;
}


function returnDriveResource(urlArgs, type, domain) {
  try {
    if (!domain) {
      switch(type)
      {
        case "DOCUMENT":
          var resource = DriveApp.getFileById(urlArgs[5])   
          break;
        case "SPREADSHEET":
          var key = urlArgs[4].split("key=")[1].split("#")[0];
          var resource = DriveApp.getFileById(key);
          break;
        case "PRESENTATION":
          var resource = DriveApp.getFileById(urlArgs[5])   
          break;
        case "DRAWING":
          var resource = DriveApp.getFileById(urlArgs[5])   
          break;
        case "FORM":
          var resource = DriveApp.getFileById(urlArgs[5])   
          break;
        case "FOLDER":
          var resource = DriveApp.getFileById(urlArgs[4])   
          break;
        case "FOLDERVIEW":
          var resource = DriveApp.getFileById(urlArgs[3].split("id=")[1].split("&")[0]);   
          break;
        case "OTHER":
          var resource = DriveApp.getFileById(urlArgs[5])   
          break;
        default:
          return;
      }
    } else {
      switch(type)
      {
        case "DOCUMENT":
          var resource = DriveApp.getFileById(urlArgs[7])
          break;
        case "SPREADSHEET":
          var key = urlArgs[6].split("key=")[1].split("#")[0];
          var resource = DriveApp.getFileById(key);
          break;
        case "PRESENTATION":
          var resource = DriveApp.getFileById(urlArgs[7])   
          break;
        case "DRAWING":
          var resource = DriveApp.getFileById(urlArgs[7])   
          break;
        case "FORM":
          var resource = DriveApp.getFileById(urlArgs[7])   
          break;
        case "FOLDER":
          var resource = DriveApp.getFileById(urlArgs[6])   
          break;
        case "FOLDERVIEW":
          var resource = DriveApp.getFileById(urlArgs[5].split("id=")[1].split("&")[0]);   
          break;
        case "OTHER":
          var resource = DriveApp.getFileById(urlArgs[7])   
          break;
        default:
           return;
      }
    }
    return resource; 
  } catch(err) {
    return err.message;
  }
}




function returnDriveIdFromURL(url) {
  var urlArgs = url.split("/");
  var type = returnType(url, urlArgs);
  var domain = false;
  var id = ''
  if (urlArgs.indexOf("a")!=-1) {
    domain = true;
  }
  var resource = returnDriveResource(urlArgs, type, domain);
  if (resource.toString().indexOf("No item with the given ID could be found, or you do not have permission to access it.")!=-1) {
    var resourceId = "Can't access";
  } else {
    var resourceId = resource.getId(); 
  }
  return resourceId;
}



function formFolio_logAddSubmission()
{
  var systemName = ScriptProperties.getProperty("systemName")
  NVSL.log("Drive%20Resource%20Added%20To%20Folders", scriptName, scriptTrackingId, systemName)
}


function formFolio_logCopySubmission()
{
  var systemName = ScriptProperties.getProperty("systemName")
  NVSL.log("Drive%20Resource%20Copied%20To%20Folders", scriptName, scriptTrackingId, systemName)
}

function formFolio_logNotSharedProperly()
{
  var systemName = ScriptProperties.getProperty("systemName")
  NVSL.log("Drive%20Resource%20Not%20Shared%20Prior%20To%20Submission", scriptName, scriptTrackingId, systemName)
}

//This function makes a call to the correct installation function.
//Embed this in the function that creates first actively loaded UI panel within the script
function setSid() { 
  var scriptNameLower = scriptName.toLowerCase();
  var sid = ScriptProperties.getProperty(scriptNameLower + "_sid");
  if (sid == null || sid == "")
  {
    var dt = new Date();
    var ms = dt.getTime();
    var ms_str = ms.toString();
    ScriptProperties.setProperty(scriptNameLower + "_sid", ms_str);
    var uid = UserProperties.getProperty(scriptNameLower + "_uid");
    if (uid) {
      logRepeatInstall();
    } else {
      logFirstInstall();
      UserProperties.setProperty(scriptNameLower + "_uid", ms_str);
    }      
  }
}

function logRepeatInstall() {
  var systemName = ScriptProperties.getProperty("systemName")
  NVSL.log("Repeat%20Install", scriptName, scriptTrackingId, systemName)
}

function logFirstInstall() {
  var systemName = ScriptProperties.getProperty("systemName")
  NVSL.log("First%20Install", scriptName, scriptTrackingId, systemName)
}
