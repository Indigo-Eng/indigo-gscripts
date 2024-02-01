function createMonthlySpreadsheet() {
  // folder ID passed into function so as to find later
  var folderId = "19Wa5jCEgG0-Ej441BfxrYJFUYMsWcOm0";
  // this spreasheet 
  var sourceSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  // sheets to be generated
  // for this to work, the AirByte automation for Tokio Marina has to be in place (it is, currently scheduled to run at 1 am on the first day of every month)
  var sheetNames = ["TOKIO_CYBER_LAST_MONTH_VIEW", "TOKIO_CYBER_SUMMARY"];

  var newSpreadsheetName = "Tokio_Cyber_report"
  
  // Add Headers
  var headers = [
    "Tokio Marine HCC - Cyber & Professional Lines Group",
    "Sample Reporting Template",
    "Indigo Risk Retention Group, Inc",
    ""
  ];
  // this call works for combining in single sheet
  createSpreadsheetWithParams(folderId, sourceSpreadsheet, sheetNames, newSpreadsheetName, headers, true); 
}

function createTimeDrivenTriggers() {
  // Trigger every 1st of the month
    ScriptApp.newTrigger('createMonthlySpreadsheet')
    .timeBased()
    .everyDays(1) // Sets the trigger to run every day
    .atHour(17) // Sets the hour for the trigger (5 PM in 24-hour format)
    .nearMinute(30) // Targets the trigger to run as close as possible to 5:30 PM
    .create();
}
