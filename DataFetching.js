function getSettingsValueRangeByRow(startRow) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings");
  
  // Get the last column with data in the specified row
  var lastColumn = sheet.getLastColumn();
  var range = sheet.getRange(startRow, 1, 1, lastColumn);
  var values = range.getValues();
  
  // Find the last column with data in row 2
  var lastDataColumn = values[0].findIndex(value => value === "") + 1; // Find the first empty cell
  
  // If all cells are filled, lastDataColumn will be -1, so adjust to lastColumn
  if (lastDataColumn === 0) {
    lastDataColumn = lastColumn;
  }
  
  // Define the dynamic range
  var endColumn = String.fromCharCode(64 + lastDataColumn); // Convert column number to letter
  var dynamicRange = 'B' + startRow + ':' + endColumn + startRow;
  
  return dynamicRange;
}

function getRangeByHeaderName(sheetName, headerName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]; // Assuming headers are in the first row

  if (headers.length === 0) {
    throw new Error('Header not found: ' + headerName);
  }  

  var columnIndex = headers.indexOf(headerName) + 1; // Column index is 1-based
  var columnLetter = columnToLetter(columnIndex);

  return columnLetter + '2:' + columnLetter;
}

function getRangeByColumnIndex(column) {
  var columnLetter = columnToLetter(column);
  return columnLetter + '2:' + columnLetter;
}

function columnToLetter(column) {
  var temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}



function getDropDownValues(sheetName, columnName) {
  // Open the spreadsheet and the specified sheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  
  // Get the specified range
  var range = sheet.getRange(getRangeByHeaderName(sheetName, columnName));
  
  // Get the data validation rules for the range
  var rules = range.getDataValidations().flat();
  
  // Use a Set to store unique dropdown values
  var dropDownValuesSet = new Set();
  
  // Extract the dropdown values from the rules
  rules.forEach(rule => {
    if (rule && rule.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
      var values = rule.getCriteriaValues();
      values[0].forEach(value => dropDownValuesSet.add(value));
    }
  })

  // Convert the Set to an array and return
  return Array.from(dropDownValuesSet);
}

function getUniqueValues(sheetName, rangeNotation) {
  // Open the spreadsheet and the specified sheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  
  // Get all values in the specified range
  var range = sheet.getRange(rangeNotation);
  var values = range.getValues();
  
  // Use a Set to store unique values
  var uniqueValuesSet = new Set(values.flatMap(row => row.map(value => value.toString().trim())).filter(value => value !== ''));
  
  // Convert the Set to an array
  return Array.from(uniqueValuesSet);
}



// FETCH DATA FROM 'Settings' Sheet
function getSSTRate() {
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  var sstSettings = ss.getRange(getSettingsValueRangeByRow(4)).getValues().flat().filter(String);
  
  if (sstSettings[1] == true) {
    return sstSettings[0];
  }
  return 0;
}

function getGSTRate() {
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  var sstSettings = ss.getRange(getSettingsValueRangeByRow(5)).getValues().flat().filter(String);
  
  if (sstSettings[1] == true) {
    return sstSettings[0];
  }
  return 0;
}

function fetchBranches() {
  // Get the active spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get source data sheet ("Settings")
  var sourceSheet = ss.getSheetByName('Settings'); 
  
  // Fetch the data from "Settings" sheet
  var sourceRange = sourceSheet.getRange(getSettingsValueRangeByRow(3)); 
  var sourceValues = sourceRange.getValues().flat();
  
  // Remove duplicates and empty values
  var data = [...new Set(sourceValues)].filter(String);

  // Return the array
  return data;
}

function fetchPriorities() {
  // Get the active spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get source data sheet ("Settings")
  var sourceSheet = ss.getSheetByName('Settings'); 
  
  // Fetch the data from "Settings" sheet
  var sourceRange = sourceSheet.getRange(getSettingsValueRangeByRow(2)); 
  var sourceValues = sourceRange.getValues().flat();
  
  // Remove duplicates and empty values
  var data = [...new Set(sourceValues)].filter(String);

  // Return the array
  return data;
}



// FETCH DATA FROM DROPDOWN
function fetchSalesStatuses() {
  return getDropDownValues(SALES_HEADERS.SHEET_NAME, SALES_HEADERS.STATUS);
}

function fetchSalesChannels() {
  return getDropDownValues(SALES_HEADERS.SHEET_NAME, SALES_HEADERS.SALES_CHANNEL);
}

function fetchPaymentMethods() {
  return getDropDownValues(SALES_HEADERS.SHEET_NAME, SALES_HEADERS.PAYMENT_METHOD);
}

function fetchRestockStatuses() {
  return getDropDownValues(RESTOCK_HEADERS.SHEET_NAME, RESTOCK_HEADERS.STATUS);
}

function fetchRestockPaymentMethods() {
  return getDropDownValues(RESTOCK_HEADERS.SHEET_NAME, RESTOCK_HEADERS.PAYMENT_METHOD);
}

function fetchProductCategories() {
  return getDropDownValues(INVENTORY_HEADERS.SHEET_NAME, INVENTORY_HEADERS.CATEGORY);
}

function fetchEmployeeStatuses() {
  return getDropDownValues(EMPLOYEE_HEADERS.SHEET_NAME, EMPLOYEE_HEADERS.STATUS);
}

function fetchJobTitles() {
  return getDropDownValues(EMPLOYEE_HEADERS.SHEET_NAME, EMPLOYEE_HEADERS.JOB_TITLE);
}

function fetchDepartments() {
  return getDropDownValues(EMPLOYEE_HEADERS.SHEET_NAME, EMPLOYEE_HEADERS.DEPARTMENT);
}

function fetchSupplierCategories() {
  return getDropDownValues(SUPPLIER_HEADERS.SHEET_NAME, SUPPLIER_HEADERS.CATEGORY);
}

function fetchLiabilityCategories() {
  return getDropDownValues(LIABILITY_HEADERS.SHEET_NAME, LIABILITY_HEADERS.CATEGORY);
}

function fetchLiabilityFrequency() {
  return getDropDownValues(LIABILITY_HEADERS.SHEET_NAME, LIABILITY_HEADERS.FREQUENCY);
}

function fetchLiabilityStatuses() {
  return getDropDownValues(LIABILITY_HEADERS.SHEET_NAME, LIABILITY_HEADERS.STATUS);
}

function fetchAssetCategories() {
  return getDropDownValues(ASSET_HEADERS.SHEET_NAME, ASSET_HEADERS.CATEGORY);
}

function fetchExpensesCategories() {
  return getDropDownValues(EXPENSE_HEADERS.SHEET_NAME, EXPENSE_HEADERS.CATEGORY);
}

function fetchIncomeCategories() {
  return getDropDownValues(INCOME_HEADERS.SHEET_NAME, INCOME_HEADERS.CATEGORY);
}



// FETCH UNIQUE DATA IN COLUMN
function fetchCustomers() {
  var rangeNotation = getRangeByHeaderName(CUSTOMER_HEADERS.SHEET_NAME, CUSTOMER_HEADERS.CONTACT_NUMBER);
  return getUniqueValues(CUSTOMER_HEADERS.SHEET_NAME, rangeNotation);
}

function fetchProducts() {
  var rangeNotation = getRangeByHeaderName(INVENTORY_HEADERS.SHEET_NAME, INVENTORY_HEADERS.PRODUCT_ID);
  return getUniqueValues(INVENTORY_HEADERS.SHEET_NAME, rangeNotation);
}

function fetchSuppliers() {
  return getUniqueValues('Supplier', 'A2:A');
}



// FETCH AND SET BASED ON DATA
function getProductPrice(productID) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(INVENTORY_HEADERS.SHEET_NAME);

  var range = sheet.getRange('A2:E');
  var values = range.getValues();
  
  // Iterate through the rows to find the matching ProductID
  for (var i = 0; i < values.length; i++) {
    if (values[i][0] === productID) {
      return values[i][4];
    }
  }
  return null;
}

function getProductCost(productID) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(INVENTORY_HEADERS.SHEET_NAME);

  var range = sheet.getRange('A2:F');
  var values = range.getValues();
  
  // Iterate through the rows to find the matching ProductID
  for (var i = 0; i < values.length; i++) {
    if (values[i][0] === productID) {
      return values[i][5];
    }
  }
  return null;
}

function getProductStock(productID) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(INVENTORY_HEADERS.SHEET_NAME);

  var range = sheet.getRange('A2:D'); // Adjust the range based on where your ProductID and Price are located
  var values = range.getValues();
  
  // Iterate through the rows to find the matching ProductID
  for (var i = 0; i < values.length; i++) {
    if (values[i][0] === productID) { // Assuming ProductID is in the first column
      return values[i][3];
    }
  }
  return null;
}

function setProductStock(productID, stock) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(INVENTORY_HEADERS.SHEET_NAME);

  var range = sheet.getRange('A2:D');
  var values = range.getValues();
  
  // Iterate through the rows to find the matching ProductID
  for (var i = 0; i < values.length; i++) {
    if (values[i][0] === productID) { // Assuming ProductID is in the first column
      values[i][3] = stock;
    }
  }
  range.setValues(values);
}
