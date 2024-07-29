function insertDropdown(sheet, range, values) {
  sheet.getRange(range).setDataValidation(
    SpreadsheetApp.newDataValidation()
    .requireValueInList(values)
    .build()
  );
}

function createDropDown(source, target, sRange, tRange) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet(); // Get the active spreadsheet

  // Replace 'source' with the name of the sheet containing the source data
  var sourceSheet = spreadsheet.getSheetByName(source); 

  // Fetch the values from the source sheet
  var sourceRange = sourceSheet.getRange(sRange); 
  var sourceValues = sourceRange.getValues().flat();
  var uniqueValues = [...new Set(sourceValues)].filter(String);
  
  // Replace 'target with the name of the sheet where the drop-down will be created
  var targetSheet = spreadsheet.getSheetByName(target); 
  
  // Create data validation rule and set it as a dropdown
  insertDropdown(targetSheet, tRange, uniqueValues);
}

function setSharedDropdownTrigger(func) {
  // Get the active spreadsheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Remove any existing triggers of the same function
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === func) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Create a new onEdit trigger for the function
  ScriptApp.newTrigger(func)
    .forSpreadsheet(spreadsheet)
    .onEdit()
    .create();
}

function insertSaleSharedDropdowns() {
  if (SpreadsheetApp.getActiveSheet().getName() !== "Settings") {
    return; 
  }
  var sheetName = 'Sales';
  // Insert priority dropdown:
  createDropDown('Settings', sheetName, getSettingsValueRangeByRow(2), getRangeByHeaderName(sheetName, SALES_HEADERS.PRIORITY));

  // Insert branch dropdown:
  createDropDown('Settings', sheetName, getSettingsValueRangeByRow(3), getRangeByHeaderName(sheetName, SALES_HEADERS.BRANCH));
}

function insertRestockSharedDropdowns() {
  if (SpreadsheetApp.getActiveSheet().getName() !== "Settings") {
    return; 
  }
  var sheetName = 'Restock';
  // Insert priority dropdown:
  createDropDown('Settings', sheetName, getSettingsValueRangeByRow(2), getRangeByHeaderName(sheetName, RESTOCK_HEADERS.PRIORITY));

  // Insert branch dropdown:
  createDropDown('Settings', sheetName, getSettingsValueRangeByRow(3), getRangeByHeaderName(sheetName, RESTOCK_HEADERS.BRANCH));
}

function insertEmployeeSharedDropdowns() {
  if (SpreadsheetApp.getActiveSheet().getName() !== "Settings") {
    return; 
  }
  var sheetName = 'Employee';
  // Insert branch dropdown:
  createDropDown('Settings', sheetName, getSettingsValueRangeByRow(3), getRangeByHeaderName(sheetName, EMPLOYEE_HEADERS.BRANCH));
}
