function createSpreadsheetWithParams(folderId, sourceSpreadsheet, sheetNames, newSpreadsheetNameBase, headers, copyToSingleSheetFlag = true, frequency = "m") {
  //var sourceSpreadsheet = SpreadsheetApp.openById(spreadsheetId);

  // Format the date to "year_month" pattern
  var namePrefix = getPrefix();  // defaults to monthly

  // Create a new Spreadsheet monthly
  var newSpreadsheetName = newSpreadsheetNameBase + "-" + namePrefix;
  var newSpreadsheet = SpreadsheetApp.create(newSpreadsheetName);
  var sheet = newSpreadsheet.getSheets()[0];
  sheet.setName(newSpreadsheetName);

  if (headers != null && headers.length > 0){
    headers.forEach(function(header, index) {
      sheet.getRange(index + 1, 1).setValue(header);
    });
  }

  // Loop through each sheet name provided in the parameter
  if (copyToSingleSheetFlag) {
    generateSingleSheetFromMany(sourceSpreadsheet, sheet, sheetNames, headers);
  } else {
    copySheetsToNewSpreadsheet(sourceSpreadsheet, newSpreadsheet, sheetNames);
  }

  moveToFolder(newSpreadsheet.getId(), folderId);
}

function getPrefix(frequency = 'm'){
  var prefix = "";
  var date = new Date();
  var year = date.getFullYear();
  var month = date.getMonth() + 1; // Months are 0-indexed
  prefix = year + "_" + (month < 10 ? "0" + month : month) ; // Ensures month is two digits  

  if (frequency == "w"){
    prefix += "_" + Math.floor((day - 1) / 7); // get the week of the month
  } else if (frequency == "d") {
    prefix += "_" + date.getDate(); 
  }

  return prefix;
}

// Loop through each sheet name provided in the parameter and combine into 1 sheet
function generateSingleSheetFromMany(sourceSpreadsheet, workingSheet, sheetNames, headers){
    sheetNames.forEach(function(sheetName) {
      var dataSheet = sourceSpreadsheet.getSheetByName(sheetName);
      var dataRange = dataSheet.getDataRange();
      var dataValues = dataRange.getValues();
      var startRow = 0;
      if (headers != null && headers.length > 0) {
        startRow = headers.length + 2; // Assuming headers take up the first 4 rows
      }
      workingSheet.getRange(startRow, 1, dataRange.getNumRows(), dataRange.getNumColumns()).setValues(dataValues);
      startRow += dataRange.getNumRows() + 2; // Adjust for a two-line gap between data sets
    }
  );
}

function copySheetsToNewSpreadsheet(sourceSpreadsheet, newSpreadsheet, sheetNames) {
  sheetNames.forEach(function(sheetName) {
    var sourceSheet = sourceSpreadsheet.getSheetByName(sheetName);

    if (sourceSheet) {
      // Copy the sheet to the new spreadsheet
      sourceSheet.copyTo(newSpreadsheet);

      // Optionally, rename the copied sheet in the new spreadsheet
      var lastSheet = newSpreadsheet.getSheets()[newSpreadsheet.getNumSheets() - 1];
      lastSheet.setName(sheetName);
    }
  });
}
 

function moveToFolder(spreadsheetId, folderId) {
  // Get the file and the destination folder
  var file = DriveApp.getFileById(spreadsheetId);
  var folder = DriveApp.getFolderById(folderId);

  // Move the file to the destination folder
  file.moveTo(folder);
}

