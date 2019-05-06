function onOpen() {
    var menu = SpreadsheetApp.getUi();
    menu.createMenu('Functions')
    .addItem('Compose Letters', 'createLetters')
    .addToUi();
    SpreadsheetApp.flush();
  }
  
  /*
  * Function: createLettersFromSheet
  * Purpose: Loop through spreadsheet data and merge with a template document.
  */
  function createLetters() {
    // config variables are found on named sheets, "Data", "app_vars", and "template_vars"
    var dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("parsed_data");
    var responseSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses");
    var startRowCell;
    var appVarsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("app_vars");
    var appVars = appVarsSheet
    .getRange(3, 1, appVarsSheet.getLastRow(), 2)
    .getValues()
    .reduce(function(obj, row, i) { // transform 2D array of sheet values into { key: value } hash
      if (row[0] === 'startRow') {
        // hold onto rowId cell so it can be set later
        startRowCell = appVarsSheet.getRange('B' + (i + 3));
      }
      obj[row[0]] = row[1];
      return obj;
    }, {});
    var templateVarsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("template_vars");
    var templateFields = templateVarsSheet
    .getRange(3, 1, templateVarsSheet.getLastRow(), 3)
    .getValues()
    .reduce(function(obj, row) {
      if (row[0].trim())
        obj[row[0]] = [row[1], row[2]];
      return obj;
    }, {});
    var transitionFormsFolderId = appVars.folderID;
    var transitionFormsFolder = DriveApp.getFolderById(transitionFormsFolderId);
    var rowId = appVars.startRow || 2;
    var templateDocId = appVars.templateDocId;
    var endRow = appVars.maxRowsOverride
    || responseSheet
    .getRange('A:A')
    .getValues()
    .filter(function(val) { return val[0] })
    .length;
    var newDocNameCol = appVars.newDocNameCol;
    var date = new Date();
    if (rowId > endRow) {
      return Browser.msgBox("No rows to write.\\nStart line = " + rowId + "\\nEnd line = " + endRow);
    }
    // add date to templateFields object
    // Begin letter creation.
    try {
      while (rowId <= endRow) {
        // map template fields to student data from data sheet
        var templateData = Object.keys(templateFields).reduce(function(obj, col) {
          var cell = col + rowId;
          var sheet = col.slice(0, 11) === 'parsed_data' ? dataSheet : responseSheet;
          var field = templateFields[col][1];
          var val = sheet.getRange(cell).getValue();
          if (field === '%dateline%') {
            var d = new Date(val);
            val = d.getMonth() + 1 + '/' + d.getDate() + '/' + d.getFullYear();
          }
          obj[field] = val;
          return obj;
        }, {});
        
        var schoolName = dataSheet.getRange(appVars.newFolderCol + rowId).getValue();
        
        // Put together newDocName
        var newDocName = newDocNameCol && dataSheet.getRange(newDocNameCol + rowId).getValue();
        newDocName = newDocName || 'newDoc_' + Date.now();
        // get destination folder for New School by its unique ID
        // or create it if not found
        if (!schoolName) throw 'No receiving school found for student, ' + newDocName;
        var schoolFolder = getSchoolFolder(schoolName, transitionFormsFolder);
        
        // Write data to template.
        var newDoc = writeToTemplate(templateDocId, schoolFolder, newDocName, templateData);
        
        SpreadsheetApp.flush();
        rowId++;
      };
      startRowCell.setValue(rowId);
      SpreadsheetApp.flush();
      Browser.msgBox("All documents written.");
    } catch (er) {
      Browser.msgBox(er);
    }
  }
  
  
  /***
  * Function: writeToTemplate
  * Purpose: This function takes a document template ID, a folder Object to create the new document in,
  * a new document name, a date, and an object mapping template fields to values.
  * It will merge the data into a copy of the template document, which
  * is renamed to the new document name, and it returns the new document.
  */
  function writeToTemplate(templateDocId, folder, newDocName, data) {
    // find duplicate documents and delete them
    var duplicateDocs = folder.getFilesByName(newDocName);
    while (duplicateDocs.hasNext()) {
      var doc = duplicateDocs.next();
      doc.setTrashed(true);
      folder.removeFile(doc);
    }
    
    var newDocId = DriveApp
    .getFileById(templateDocId)
    .makeCopy(folder)
    .getId();
    
    var newDoc = DocumentApp.openById(newDocId);
    
    newDoc.setName(newDocName);
    
    var newBody = newDoc.getBody();
    
    // Replace variables with spreadsheet data.
    Object.keys(data).forEach(function(key) {
      newBody.replaceText(key, data[key]);
    });
    // newBody.replaceText("%dateline%", dateline);
    SpreadsheetApp.flush();
  
    return newDoc;
  }
  
  function getSchoolFolder(schoolName, parentFolder) {
    var schoolFoldersByName = parentFolder.getFoldersByName(schoolName);
    if (schoolFoldersByName.hasNext()) {
      return schoolFoldersByName.next();
    } else {
      return parentFolder.createFolder(schoolName)
    }
  }
  