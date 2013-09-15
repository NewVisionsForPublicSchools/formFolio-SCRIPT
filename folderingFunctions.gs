

function formFolio_addToFolders(recursed) {
  try {
    var properties = ScriptProperties.getProperties();
    var sheetIdMappings = Utilities.jsonParse(properties.sheetIdMappings);
    var urlQid = properties.urlQId;
    var userNameId = properties.userNameId;
    var formSheetId = properties.formSheetId;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      ss = SpreadsheetApp.openById(properties.ssId);
    }
    var sheet = getSheetById(ss, formSheetId);
    var thisSubmission = sheet.getActiveRange().getValues();
    var thisRow = sheet.getActiveRange().getRow();
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var formUrl = ss.getFormUrl();
    var form = FormApp.openByUrl(formUrl);
    var responses = form.getResponses();
    var thisResponseUrl = responses[responses.length-1].toPrefilledUrl();
    var urlHeader = form.getItemById(urlQid).getTitle();
    if (userNameId!="Username") {
      var usernameHeader = form.getItemById(userNameId).getTitle();
      var usernameIndex = headers.indexOf(usernameHeader);
    } else {
      var usernameIndex = headers.indexOf("Username");
    }
    var urlIndex = headers.indexOf(urlHeader);  
    var linkCol = headers.indexOf("Link to Resource")+1;
    var newUrlCol = headers.indexOf("Resource URL")+1;
    var statusCol = headers.indexOf("Foldering Status")+1;
    var typeCol = headers.indexOf("Resource Type")+1;
    var idCol = headers.indexOf("Drive ID")+1;
    var url = thisSubmission[0][urlIndex];
    var email = thisSubmission[0][usernameIndex];
    var items = form.getItems();
    var itemTitles = [];
    var hasFolders = [];
    var folderIds = [];
    var sheets = ss.getSheets();
    for (var j=0; j<sheetIdMappings.length; j++) {
      var thisSheet = returnMappedSheet(ss, sheetIdMappings[j].formQId, sheetIdMappings);
      var theseHeaders = thisSheet.getRange(1, 1, 1, thisSheet.getLastColumn()).getValues()[0];
      var thisFolderKeyIndex = theseHeaders.indexOf("Folder Key");
      if ((thisSheet)&&(thisSheet!="not found")) {
        if (sheetIdMappings[j].formQId!="Username") {
          var thisQuestionTitle = form.getItemById(parseInt(sheetIdMappings[j].formQId)).getTitle();
          var colNum = headers.indexOf(thisQuestionTitle);
        } else {
          var colNum = headers.indexOf("Username");
        }
        if (colNum==-1) {
          continue;
        } else {
          var thisResponse = thisSubmission[0][colNum].split(", ");
        }
        var thisFolderLookupData = thisSheet.getRange(2, 1, thisSheet.getLastRow()-1, thisSheet.getLastColumn()).getValues();
        var found = false;
        for (var r = 0; r<thisResponse.length; r++) {
          for (var k=0; k<thisFolderLookupData.length; k++) {
            if ((thisFolderLookupData[k][0] == thisResponse[r])&&(thisFolderLookupData[k][thisFolderKeyIndex]!='')) {
              var thisFolderKey = thisFolderLookupData[k][thisFolderKeyIndex];
              folderIds.push(thisFolderKey);
              found = true;
            }
          }
          if ((found == false)&&(sheetIdMappings[j].formQId=="Username")&&(properties.userFolderMode == "0")&&(!recursed)) {
            formFolio_addUserToFolderSheet(thisResponse[0]);
            SpreadsheetApp.flush();
            formFolio_createRefreshUserFolders();
            SpreadsheetApp.flush();
            formFolio_addToFolders(true);
            return;
          }
        }
      }
    }
    var urlArgs = url.split("/");
    var type = returnType(url, urlArgs);
    if (type != "") {
      var resourceId = returnDriveIdFromURL(url);
      if (resourceId != "Can't access") {
        if ((type=="FOLDER")||(type=="FOLDERVIEW")) {
          var docsListResource = DocsList.getFolderById(resourceId)
          } else { 
            var docsListResource = DocsList.getFileById(resourceId);
          }
        var successFolders = [];
        for (var f=0; f<folderIds.length; f++) {
          var folder = DocsList.getFolderById(folderIds[f]);
          docsListResource.addToFolder(folder);
          successFolders.push(folder.getName());
        }
        sheet.getRange(thisRow, newUrlCol).setValue(url);
        sheet.getRange(thisRow, linkCol).setValue('=hyperlink("'+url+'", "' + docsListResource.getName() + '")');
        sheet.getRange(thisRow, statusCol).setValue("Successfully added to folder(s): " + successFolders.join(", "));
        sheet.getRange(thisRow, idCol).setValue(resourceId);
        if (type!="FOLDERVIEW") {
          sheet.getRange(thisRow, typeCol).setValue(type);
        } 
        if (type=="FOLDERVIEW") {
          sheet.getRange(thisRow, typeCol).setValue('FOLDER');
        }
        formFolio_logAddSubmission();
      } else {
        formFolio_sendShareNotification(email, url, thisResponseUrl);
        sheet.getRange(thisRow, statusCol).setValue("Resource wan't shared properly. Notification sent to user: " + email);
        formFolio_logNotSharedProperly();
      }
    } else {
      if (properties.nonDriveWarning=="true") {
        formFolio_sendWrongResourceTypeNotification(email, url, formUrl);
        sheet.getRange(thisRow, statusCol).setValue("Non Drive resource URL. No notifications sent");
      } else {
         formFolio_sendWrongResourceTypeNotification(email, url, formUrl);
        sheet.getRange(thisRow, statusCol).setValue("Invalid resource URL. Notification sent to user:" + email);
      }
    }
  } catch(err) {
    sheet.getRange(thisRow, statusCol).setValue("There was a problem adding this resource to folders: " + err.message);
  }
}


function formFolio_copyToFolders(recursed) {
  try {
    var properties = ScriptProperties.getProperties();
    var sheetIdMappings = Utilities.jsonParse(properties.sheetIdMappings);
    var urlQid = properties.urlQId;
    var userNameId = properties.userNameId;
    var formSheetId = properties.formSheetId;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      ss = SpreadsheetApp.openById(properties.ssId);
    }
    var sheet = getSheetById(ss, formSheetId);
    var thisSubmissionRange = sheet.getActiveRange();
    var thisSubmission = thisSubmissionRange.getValues();
    var thisSubmissionObject = getRowsData(sheet, thisSubmissionRange, 1)[0];
    var thisRow = thisSubmissionRange.getRow();
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var formUrl = ss.getFormUrl();
    var form = FormApp.openByUrl(formUrl);
    var responses = form.getResponses();
    var thisResponseUrl = responses[responses.length-1].toPrefilledUrl();
    var urlHeader = form.getItemById(urlQid).getTitle();
    if (userNameId!="Username") {
      var usernameHeader = form.getItemById(userNameId).getTitle();
      var usernameIndex = headers.indexOf(usernameHeader);
    } else {
      var usernameIndex = headers.indexOf("Username");
    }
    var urlIndex = headers.indexOf(urlHeader);  
    var linkCol = headers.indexOf("Link to Resource")+1;
    var newUrlCol = headers.indexOf("Resource URL")+1;
    var statusCol = headers.indexOf("Foldering Status")+1;
    var typeCol = headers.indexOf("Resource Type")+1;
    var idCol = headers.indexOf("Drive ID")+1;
    var url = thisSubmission[0][urlIndex];
    var emails = [];
    emails[0] = thisSubmission[0][usernameIndex];
    var items = form.getItems();
    var itemTitles = [];
    var hasFolders = [];
    var folderIds = [];
    var sheets = ss.getSheets();
    for (var j=0; j<sheetIdMappings.length; j++) {
      var thisSheet = returnMappedSheet(ss, sheetIdMappings[j].formQId, sheetIdMappings);
      var theseHeaders = thisSheet.getRange(1, 1, 1, thisSheet.getLastColumn()).getValues()[0];
      var thisFolderKeyIndex = theseHeaders.indexOf("Folder Key");
      if ((thisSheet)&&(thisSheet!="not found")) {
        if (sheetIdMappings[j].formQId!="Username") {
          var thisQuestionTitle = form.getItemById(parseInt(sheetIdMappings[j].formQId)).getTitle();
          var colNum = headers.indexOf(thisQuestionTitle);
        } else {
          var colNum = headers.indexOf("Username");
        }
        if (colNum==-1) {
          continue;
        } else {
          var thisResponse = thisSubmission[0][colNum].split(", ");
        }
        var found = false;
        var thisFolderLookupData = thisSheet.getRange(2, 1, thisSheet.getLastRow()-1, thisSheet.getLastColumn()).getValues();
        for (var r = 0; r<thisResponse.length; r++) {
          for (var k=0; k<thisFolderLookupData.length; k++) {
            if (thisFolderLookupData[k][0] == thisResponse[r]) {
              var thisFolderKey = thisFolderLookupData[k][thisFolderKeyIndex];
              folderIds.push(thisFolderKey);
              found = true;
            }
          }
        }
        if ((found == false)&&(sheetIdMappings[j].formQId=="Username")&&(properties.userFolderMode == "0")&&(!recursed)) {
          formFolio_addUserToFolderSheet(thisResponse[0]);
          SpreadsheetApp.flush();
          formFolio_createRefreshUserFolders();
          SpreadsheetApp.flush();
          formFolio_copyToFolders(true);
          return;
        }
      }
    }
    var urlArgs = url.split("/");
    var type = returnType(url, urlArgs);
    if (type != "") {
      var resourceId = returnDriveIdFromURL(url);
      if (resourceId != "Can't access") {
        if ((type=="FOLDER")||(type=="FOLDERVIEW")) {
          var docsListResource = DocsList.getFolderById(resourceId);
          var originalFileName = docsListResource.getName();
          var newName = replaceStringFields(headers, properties.copyName, thisSubmissionObject, originalFileName);
          var copyResource = formFolio_copyFolder(docsListResource, newName);
          var copyResourceDrive = DriveApp.getFolderById(copyResource.getId());
        } else { 
          var docsListResource = DocsList.getFileById(resourceId);
          var originalFileName = docsListResource.getName();
          var newName = replaceStringFields(headers, properties.copyName, thisSubmissionObject, originalFileName);
          var copyResource = docsListResource.makeCopy(newName);
          var copyResourceDrive = DriveApp.getFileById(copyResource.getId());
        }
        if ((properties.descQId)&&(properties.descQId != "")) {
          var thisQuestionTitle = form.getItemById(parseInt(properties.descQId)).getTitle();
          var colNum = headers.indexOf(thisQuestionTitle);
          var thisDescription = thisSubmission[0][colNum];
          if ((thisDescription)&&(thisDescription!='')) {
            copyResourceDrive.setDescription(thisDescription.toString());
          }
        }
        var successFolders = [];
        for (var f=0; f<folderIds.length; f++) {
          var folder = DocsList.getFolderById(folderIds[f]);
          copyResource.addToFolder(folder);
          successFolders.push(folder.getName());
        }
        copyResource.removeFromFolder(DocsList.getRootFolder());
        if (type!="FOLDERVIEW") {
          sheet.getRange(thisRow, typeCol).setValue(type);
        } 
        if (type=="FOLDERVIEW") {
          sheet.getRange(thisRow, typeCol).setValue('FOLDER');
        }
        var copyUrl = copyResource.getUrl();
        sheet.getRange(thisRow, newUrlCol).setValue(copyUrl);
        sheet.getRange(thisRow, linkCol).setValue('=hyperlink("'+copyUrl+'", "' + newName + '")');
        sheet.getRange(thisRow, idCol).setValue(copyResource.getId());
        var shareType = properties.copyShareMode;
        if ((properties.collabQId)&&(properties.collabQId!='')) {
          var collabQuestionTitle = form.getItemById(parseInt(properties.collabQId)).getTitle();
          var collabColNum = headers.indexOf(collabQuestionTitle);
          var theseCollaborators = thisSubmission[0][collabColNum];
          if ((theseCollaborators)&&(theseCollaborators!='')) {
            theseCollaborators = theseCollaborators.replace(/\s+/g, '').split(",");
            for (var k=0; k<theseCollaborators.length; k++) {
              emails = emails.concat(theseCollaborators);
            }
          }
        }
        var badEmails = [];
        var goodEmails = [];
        for (var k=0; k<emails.length; k++) {
          if (shareType == "view") {
            try {
              copyResourceDrive.addViewer(emails[k]);
              goodEmails.push(emails[k]);
            } catch(err) {
              badEmails.push(emails[k]);
            }
          }
          if (shareType == "edit") {
            try {
              copyResourceDrive.addEditor(emails[k]);
              goodEmails.push(emails[k]);
            } catch(err) {
              badEmails.push(emails[k]);
            }
          }
          if ((shareType == "comment")&&(type!="FOLDER")&&(type!="FOLDERVIEW")) {
            try {
              copyResourceDrive.addCommenter(emails[k]);
              goodEmails.push(emails[k]);
            } catch(err) {
              badEmails.push(emails[k]);
            }
          }
          if ((shareType == "comment")&&((type=="FOLDER")||(type=="FOLDERVIEW"))) {
            try {
              addCommentersToFolder(copyResource, emails[k]);
              goodEmails.push(emails[k]);
            } catch(err) {
              badEmails.push(emails[k]);
            }
          }
        }
        var statusMessage = "Successfully copied, renamed, and added to folder(s): " + successFolders.join(", ");
        
        if (shareType!="none") {
          statusMessage += ". Shared with " + goodEmails.join(", ") + " with " + shareType + " privileges."; 
          if (badEmails.length>0) {
            statusMessage += " Trouble sharing with " + badEmails.join(", "); 
          }                                                   
        }
        sheet.getRange(thisRow, statusCol).setValue(statusMessage);
        formFolio_logCopySubmission();
      } else {
        formFolio_sendShareNotification(emails[0], url, thisResponseUrl);
        sheet.getRange(thisRow, statusCol).setValue("Resource wan't shared properly. Notification sent to user: " + emails[0]);
        formFolio_logNotSharedProperly();
      }
    } else {
      if (properties.nonDriveWarning=="true") {
        formFolio_sendWrongResourceTypeNotification(emails[0], url, formUrl);
        sheet.getRange(thisRow, statusCol).setValue("Non Drive resource URL. No notifications sent");
      } else {
         formFolio_sendWrongResourceTypeNotification(emails[0], url, formUrl);
        sheet.getRange(thisRow, statusCol).setValue("Invalid resource URL. Notification sent to user:" + emails[0]);
      }
    }
  } catch(err) {
    sheet.getRange(thisRow, statusCol).setValue("There was a problem copying or adding this resource to folders: " + err.message);
  }
}
