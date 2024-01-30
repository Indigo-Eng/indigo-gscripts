function createMonthlySpreadsheet() {
  // folder ID passed into function so as to find later
  var folderId = "1xnXA-sVytv3YKL-x1Kghc8_dm3orPQNz";
  // this spreasheet 
  var sourceSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  // sheets to be generated
  // for this to work, the AirByte automation for Tokio Marina has to be in place (it is, currently scheduled to run at 1 am on the first day of every month)
  var sheetNames = ["TOKIO_CYBER_LAST_MONTH_VIEW", "TOKIO_CYBER_SUMMARY"];

  var newSpreadsheetName = "Tokio Cyber"
  
  // Add Headers
  var headers = [
    "Tokio Marine HCC - Cyber & Professional Lines Group",
    "Sample Reporting Template",
    "Indigo Risk Retention Group, Inc",
    ""
  ];

  createMonthlySpreadsheetWithParams(folderId, sourceSpreadsheet, sheetNames, newSpreadsheetName, headers) 
}

function createTimeDrivenTriggers() {
  // Trigger every 1st of the month
  ScriptApp.newTrigger('createMonthlySpreadsheet')
          .timeBased()
         .onMonthDay(1)
         .atHour(8) // Sets the hour at which the trigger should start considering execution
         .nearMinute(0) // Sets the minute as close as possible to the beginning of the hour
         .create();
  }
