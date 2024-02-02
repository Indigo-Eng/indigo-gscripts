function createSpreadsheetWithParams(folderId, sourceSpreadsheet, sheetNames, newSpreadsheetNameBase, 
                                      headers, copyToSingleSheetFlag, frequency = "m") {  

  // Format the date to "year_month" pattern
  var postfix = getPrefix(frequency);  // defaults to monthly

  // Create a new Spreadsheet monthly
  var newSpreadsheetName = newSpreadsheetNameBase + "-" + postfix;

  var newSpreadsheet = SpreadsheetApp.create(newSpreadsheetName);  

  // Loop through each sheet name provided in the parameter
  if (copyToSingleSheetFlag) {    
    console.log("creating single-sheet spreadsheet " + newSpreadsheetName);    
    var sheet = newSpreadsheet.getSheets()[0];
    sheet.setName(postfix);

    if (headers != null && headers.length > 0){
      headers.forEach(function(header, index) {
        sheet.getRange(index + 1, 1).setValue(header);
      });
    }
    generateSingleSheetFromMany(sourceSpreadsheet, sheet, sheetNames, headers);
  } else {
    console.log("copying spreadsheet with " + sourceSpreadsheet.getNumSheets() + " tabs");    
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
  var startRow = 0;
  if (headers != null && headers.length > 0) {
    startRow = headers.length + 2; // Assuming headers take up the first 4 rows
  }

  sheetNames.forEach(function(sheetName) {
    var dataSheet = sourceSpreadsheet.getSheetByName(sheetName);
    var dataRange = dataSheet.getDataRange();
    var dataValues = dataRange.getValues();    
    
    // Get the total number of columns in workingSheet
    var totalColumnsWorkingSheet = dataSheet.getMaxColumns();

    // Calculate the starting column index to center the data
    
    workingSheet.getRange(startRow, 1, dataRange.getNumRows(), dataRange.getNumColumns()).setValues(dataValues);
    startRow += dataRange.getNumRows() + 2; // Adjust for a two-line gap between data sets
  });
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
  removeBlankSheets(newSpreadsheet);
}
 
function removeBlankSheets(spreadsheet) {
  var sheets = spreadsheet.getSheets();
  sheets.forEach(function(sheet) {
    if (isSheetBlank(sheet)) {
      spreadsheet.deleteSheet(sheet);
    }
  });
}

function isSheetBlank(sheet) {
  // Check if the sheet is blank
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();
  return values.length === 0 || (values.length === 1 && values[0].every(function(cell) { return cell === ''; }));
}

function moveToFolder(spreadsheetId, folderId) {
  // Get the file and the destination folder
  var file = DriveApp.getFileById(spreadsheetId);
  var folder = DriveApp.getFolderById(folderId);

  var fileName = file.getName();

  // Search for files in the destination folder with the same name
  var existingFiles = folder.getFilesByName(fileName);
  while (existingFiles.hasNext()) {
    var existingFile = existingFiles.next();
    // Delete each file with the same name
    existingFile.setTrashed(true);
  }

  // Move the file to the destination folder
  file.moveTo(folder);
}

function copyColumnsByNames(sourceSpreadsheet, targetSpreadsheet, sheetName, columnNames) {
  var sourceSheet = sourceSpreadsheet.getSheetByName(sheetName);
  var targetSheet = targetSpreadsheet.getSheetByName(sheetName) || 
                    targetSpreadsheet.insertSheet(sheetName);

  // Fetch all data from the source sheet
  var data = sourceSheet.getDataRange().getValues();
  var headers = data[0]; // Assuming the first row contains headers
  
  // If columnNames is empty, copy all columns in the original order
  if (columnNames.length === 0) {
    columnNames = headers;
  }

  // Map column names to indices
  var colIndices = columnNames.map(function(name) {
    return headers.indexOf(name);
  }).filter(function(index) {
    // Filter out columns that were not found
    return index >= 0;
  });
  
  // Filter and reorder the data based on the specified column names
  var filteredData = data.map(function(row) {
    return colIndices.map(function(index) {
      return row[index];
    });
  });
  
  // Clear the target sheet and set the new values
  targetSheet.clear();
  targetSheet.getRange(1, 1, filteredData.length, columnNames.length).setValues(filteredData);
}

function copyMultipleSheetsColumns(folderId, sourceSpreadsheetId, targetSpreadsheetId, sheetsAndColumns) {
  var sourceSpreadsheet = SpreadsheetApp.openById(sourceSpreadsheetId);
  var targetSpreadsheet = SpreadsheetApp.openById(targetSpreadsheetId);

  sheetsAndColumns.forEach(function(sheetInfo) {
    copyColumnsByNames(sourceSpreadsheet, targetSpreadsheet, sheetInfo.sheetName, sheetInfo.columnNames);
  });

  moveToFolder(folderId, targetSpreadsheet);
}

