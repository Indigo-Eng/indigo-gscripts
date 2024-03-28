function createMonthlySpreadsheet() {
  // folder ID passed into function so as to find later
  var folderId = "19Wa5jCEgG0-Ej441BfxrYJFUYMsWcOm0";
  // this spreasheet 
  var sourceSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  // sheets to be generated
  // for this to work, the AirByte automation for Tokio Marina has to be in place (it is, currently scheduled to run at 1 am on the first day of every month)
  var sheetNames = ["TOKIO_CYBER_LAST_MONTH_VIEW", "TOKIO_CYBER_SUMMARY"];  

  var newSpreadsheetName = "Tokio_Cyber_report_Test"
  var destinationSpreadsheet =  SpreadsheetApp.create(newSpreadsheetName);
  
  // Add Headers
  var headers = [
    "Tokio Marine HCC - Cyber & Professional Lines Group",
    "Sample Reporting Template",
    "Indigo Risk Retention Group, Inc",
    ""
  ];
  // this call works for combining in single sheet
  //createSpreadsheetWithParams(folderId, sourceSpreadsheet, sheetNames, newSpreadsheetName, headers, true); 
    var sheetsAndColumns = [
    {sheetName: 'TOKIO_CYBER_LAST_MONTH_VIEW', columnNames: []},
    {sheetName: 'TOKIO_CYBER_SUMMARY', columnNames: []} // Empty array implies all columns should be copied
  ];

  copyMultipleSheetsColumns(folderId, sourceSpreadsheet.getId(), destinationSpreadsheet.getId(), sheetsAndColumns);
}

function createTimeDrivenTriggers() {  
}
