// HEADER CONSTANTS
const SALES_HEADERS = Object.freeze({
  SHEET_NAME: 'Sales',
  SALES_ID: 'Sales ID',
  PRIORITY: 'Priority',
  BRANCH: 'Branch',
  CUSTOMER: 'Customer',
  DATE: 'Date',
  TIME: 'Time',
  PAYMENT_DATE: 'Payment Date',
  STATUS: 'Status',
  SALES_CHANNEL: 'Sales Channel',
  PRODUCT: 'Product',
  UNIT_COST_RM: 'Unit Cost (RM)',
  UNIT_PRICE_RM: 'Unit Price (RM)',
  QUANTITY: 'Quantity',
  STOCK_LEVEL: 'Stock Level',
  PAYMENT_METHOD: 'Payment method',
  SUB_TOTAL_RM: 'Sub-total (RM)',
  TOTAL_RM: 'Total (RM)',
  NOTES: 'Notes'
});

const RESTOCK_HEADERS = Object.freeze({
  SHEET_NAME: 'Restock',
  RESTOCK_ID: 'Restock ID',
  PRIORITY: 'Priority',
  BRANCH: 'Branch',
  SUPPLIER: 'Supplier',
  PRODUCT: 'Product',
  QUANTITY: 'Quantity',
  STATUS: 'Status',
  DATE: 'Date',
  ARRIVAL_DATE: 'Arrival Date',
  PAYMENT_DATE: 'Payment Date',
  TOTAL_RM: 'Total (RM)',
  PAYMENT_METHOD: 'Payment Method',
  NOTES: 'Notes'
});

const EMPLOYEE_HEADERS = Object.freeze({
  SHEET_NAME: 'Employee',
  EMPLOYEE_ID: 'Employee ID',
  NAME: 'Name',
  BRANCH: 'Branch',
  STATUS: 'Status',
  JOB_TITLE: 'Job Title',
  DEPARTMENT: 'Department',
  SALARY_RM: 'Salary (RM)',
  ADDRESS: 'Address',
  CONTACT_NUMBER: 'Contact Number',
  EMAIL: 'Email',
  BIRTH_DATE: 'Birth Date',
  JOINING_DATE: 'Joining Date',
  EXIT_DATE: 'Exit Date',
  BANK_ACCOUNT: 'Bank Account',
  EPF: 'EPF',
  SOCSO: 'SOCSO'
});

const INVENTORY_HEADERS = Object.freeze({
  SHEET_NAME: 'Inventory',
  PRODUCT_ID: 'Product ID',
  NAME: 'Name',
  CATEGORY: 'Category',
  QUANTITY: 'Quantity',
  SELLING_PRICE_RM: 'Selling Price (RM)',
  UNIT_COST_RM: 'Unit Cost (RM)',
  TOTAL_SOLD: 'Total Sold',
  LAST_RESTOCK: 'Last Restock',
  NOTES: 'Notes'
});

const SUPPLIER_HEADERS = Object.freeze({
  SHEET_NAME: 'Supplier',
  SUPPLIER_ID: 'Supplier ID',
  NAME: 'Name',
  CATEGORY: 'Category',
  CONTACT_NUMBER: 'Contact Number',
  EMAIL: 'Email',
  WEBSITE: 'Website',
  RATING: 'Rating',
  NOTES: 'Notes'
});

const CUSTOMER_HEADERS = Object.freeze({
  SHEET_NAME: 'Customer',
  CUSTOMER_ID: 'Customer ID',
  NAME: 'Name',
  CONTACT_NUMBER: 'Contact Number',
  EMAIL: 'Email',
  RECENT_ORDER: 'Recent Order',
  SALE_COUNT: 'Sale Count',
  TOTAL_SALES_RM: 'Total Sales (RM)',
  DAYS_SINCE_LAST_ORDER: 'Days Since Last Order',
  R_SCORE: 'R Score',
  F_SCORE: 'F Score',
  M_SCORE: 'M Score',
  SEGMENT: 'Segment'
});

const LIABILITY_HEADERS = Object.freeze({
  SHEET_NAME: 'Liabilities',
  LIABILITY_ID: 'Liability ID',
  NAME: 'Name',
  ACCOUNT: 'Account',
  CREDITOR: 'Creditor',
  CATEGORY: 'Category',
  INITIAL_AMOUNT_RM: 'Initial Amount (RM)',
  CURRENT_BALANCE_RM: 'Current Balance (RM)',
  INTEREST_RATE_PERCENT: 'Interest Rate (%)',
  FREQUENCY: 'Frequency',
  MIN_PAYMENT_RM: 'Min Payment (RM)',
  DUE_DATE: 'Due Date',
  LAST_PAYMENT: 'Last Payment',
  LAST_PAYMENT_AMOUNT_RM: 'Last Payment Amount (RM)',
  NEXT_PAYMENT_AMOUNT_RM: 'Next Payment Amount (RM)',
  START_DATE: 'Start Date',
  MATURITY_DATE: 'Maturity Date',
  CURRENCY: 'Currency',
  TAXABLE: 'Taxable',
  NOTES: 'Notes',
  STATUS: 'Status'
});

const ASSET_HEADERS = Object.freeze({
  SHEET_NAME: 'Assets',
  ASSET_ID: 'Asset ID',
  NAME: 'Name',
  CATEGORY: 'Category',
  PURCHASE_DATE: 'Purchase Date',
  COST_RM: 'Cost (RM)',
  CURRENT_VALUE_RM: 'Current Value (RM)',
  ANNUAL_DEPRECIATION_RATE_PERCENT: 'Annual Depreciation Rate (%)',
  SALVAGE_VALUE: 'Salvage Value (RM)',
  USEFUL_YEARS: 'Useful Years'
});

const EXPENSE_HEADERS = Object.freeze({
  SHEET_NAME: 'Expenses',
  EXPENSE_ID: 'Expenses ID',
  NAME: 'Name',
  DATE: 'Date',
  CATEGORY: 'Category',
  DESCRIPTION: 'Description',
  OPERATING_EXPENSE: 'Operating Expense',
  CAPITAL_EXPENDITURE: 'Capital Expenditure',
  TOTAL_RM: 'Total (RM)'
});

const INCOME_HEADERS = Object.freeze({
  SHEET_NAME: 'Income',
  INCOME_ID: 'Income ID',
  NAME: 'Name',
  DATE: 'Date',
  CATEGORY: 'Category',
  DESCRIPTION: 'Description',
  TOTAL_RM: 'Total (RM)'
});

const CASH_FLOW_HEADERS = Object.freeze({
  SHEET_NAME: 'Cashflow Dashboard',
  MONTH: 'Month',
  OPERATING_CASH_FLOW_RM: 'Operating Cash Flow (RM)',
  FREE_CASH_FLOW_RM: 'Free Cash Flow (RM)',
  DAYS_PAYABLE_OUTSTANDING: 'Days Payable Outstanding (DPO)',
  DAYS_SALES_OUTSTANDING: 'Days Sales Outstanding (DSO)',
  DAYS_INVENTORY_OUTSTANDING: 'Days Inventory Outstanding (DIO)',
  CASH_CONVERSION_CYCLE: 'Cash Conversion Cycle'
});





// HELPER FUNCTIONS
function applyDateValidation(sheet, headers, headerName) {
  const columnIndex = headers.indexOf(headerName) + 1;
  if (columnIndex > 0) {
    const range = getRangeByColumnIndex(columnIndex);
    sheet.getRange(range).setDataValidation(
      SpreadsheetApp.newDataValidation()
      .requireDate()
      .setAllowInvalid(false)
      .build()
    );
  }
}

function setNumberFormats(sheet, headers, formats) {
  Object.keys(formats).forEach(headerName => {
    const columnIndex = headers.indexOf(headerName) + 1;
    if (columnIndex > 0) {
      const range = getRangeByColumnIndex(columnIndex);
      sheet.getRange(range).setNumberFormat(formats[headerName]);
    }
  });
}

function setDropdowns(sheet, headers, dropdowns) {
  Object.keys(dropdowns).forEach(headerName => {
    var rangeNotation = getRangeByColumnIndex(headers.indexOf(headerName)+1);
    insertDropdown(sheet, rangeNotation, dropdowns[headerName]);
  });
}



// TABLE GENERATION FUNCTIONS
function createSettingsTable() {
  // Get the active spreadsheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet();

  // Rename sheet
  sheet.setName("Settings");

  // Define the table headers and data
  var headers = ['Priorities', 'Branches', 'SST Rate (%)', 'GST Rate (%)'];

  // Data
  var data = [
    ['Low', 'Medium', 'High'], 
    ['Main', 'All'],
    [6],
    [10]
  ];

  sheet.getRange(1, 1, 1, 4).merge().setValue("Settings for dropdown options:").setFontWeight('bold').setFontSize(18);

  // Set the headers in the first column
  sheet.getRange(2, 1, headers.length, 1).setValues(headers.map(header => [header]));

  // Set the data in the subsequent columns
  for (var i = 0; i < data.length; i++) {
    sheet.getRange(i + 2, 2, 1, data[i].length).setValues([data[i]]);
  }

  // Format the headers
  var headerRange = sheet.getRange(2, 1, headers.length, 1);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#d9d9d9');

  // Format columns
  sheet.autoResizeColumn(1);

  // Add SST and GST checkbox
  sheet.getRange('C4:C5').setDataValidation(
    SpreadsheetApp.newDataValidation().
    requireCheckbox().
    setAllowInvalid(false).
    build()
  );
}

function createSalesTable() {
  // Create a new sheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet();
  var sheetName = SALES_HEADERS.SHEET_NAME;

  // Rename the new sheet
  sheet.setName(sheetName);

  // Define the sales table headers
  var headers = [
    SALES_HEADERS.SALES_ID,
    SALES_HEADERS.PRIORITY,
    SALES_HEADERS.BRANCH,
    SALES_HEADERS.CUSTOMER,
    SALES_HEADERS.DATE,
    SALES_HEADERS.TIME,
    SALES_HEADERS.PAYMENT_DATE,
    SALES_HEADERS.STATUS,
    SALES_HEADERS.SALES_CHANNEL,
    SALES_HEADERS.PRODUCT,
    SALES_HEADERS.UNIT_COST_RM,
    SALES_HEADERS.UNIT_PRICE_RM,
    SALES_HEADERS.QUANTITY,
    SALES_HEADERS.STOCK_LEVEL,
    SALES_HEADERS.PAYMENT_METHOD,
    SALES_HEADERS.SUB_TOTAL_RM,
    SALES_HEADERS.TOTAL_RM,
    SALES_HEADERS.NOTES
  ];

  // Set the headers in the first row
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Format the headers
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#d9d9d9');

  // Set data types and formulas
  const formats = {
    [SALES_HEADERS.DATE]: 'yyyy-mm-dd',
    [SALES_HEADERS.TIME]: 'h:mm',
    [SALES_HEADERS.PAYMENT_DATE]: 'yyyy-mm-dd',
    [SALES_HEADERS.QUANTITY]: '##0',
    [SALES_HEADERS.STOCK_LEVEL]: '##0',
    [SALES_HEADERS.UNIT_COST_RM]: '#,##0.00',
    [SALES_HEADERS.UNIT_PRICE_RM]: '#,##0.00',
    [SALES_HEADERS.SUB_TOTAL_RM]: '#,##0.00',
    [SALES_HEADERS.TOTAL_RM]: '#,##0.00'
  };
  setNumberFormats(sheet, headers, formats);

  // Styling the table
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length).createFilter();
  
  // Applying date picker
  applyDateValidation(sheet, headers, SALES_HEADERS.DATE);
  applyDateValidation(sheet, headers, SALES_HEADERS.PAYMENT_DATE);

  // Insert dropdowns
  const dropdowns = {
    [SALES_HEADERS.STATUS]: ['Pending', 'Completed', 'Cancelled'],
    [SALES_HEADERS.SALES_CHANNEL]: ['Online', 'In-store', 'Phone', 'Email'],
    [SALES_HEADERS.PAYMENT_METHOD]: ['Online Banking', 'Cash', 'Credit/Debit Card', 'E-Wallet']
  };
  setDropdowns(sheet, headers, dropdowns);
  
  
  // Insert shared dropdowns
  SpreadsheetApp.setActiveSheet(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings"));
  insertSaleSharedDropdowns();
  setSharedDropdownTrigger('insertSaleSharedDropdowns');
  SpreadsheetApp.setActiveSheet(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName));
}

function createRestockTable() {
  // Insert new sheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet();
  var sheetName = RESTOCK_HEADERS.SHEET_NAME;

  // Rename sheet
  sheet.setName(sheetName);

  // Define the table headers
  var headers = [
    RESTOCK_HEADERS.RESTOCK_ID,
    RESTOCK_HEADERS.PRIORITY,
    RESTOCK_HEADERS.BRANCH,
    RESTOCK_HEADERS.SUPPLIER,
    RESTOCK_HEADERS.PRODUCT,
    RESTOCK_HEADERS.QUANTITY,
    RESTOCK_HEADERS.STATUS,
    RESTOCK_HEADERS.DATE,
    RESTOCK_HEADERS.ARRIVAL_DATE,
    RESTOCK_HEADERS.PAYMENT_DATE,
    RESTOCK_HEADERS.TOTAL_RM,
    RESTOCK_HEADERS.PAYMENT_METHOD,
    RESTOCK_HEADERS.NOTES
  ];

  // Set the headers in the first row
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Format the headers
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#d9d9d9');

  // Set data types
  const formats = {
    [RESTOCK_HEADERS.DATE]: 'yyyy-mm-dd',
    [RESTOCK_HEADERS.ARRIVAL_DATE]: 'yyyy-mm-dd',
    [RESTOCK_HEADERS.PAYMENT_DATE]: 'yyyy-mm-dd',
    [RESTOCK_HEADERS.QUANTITY]: '##0',
    [RESTOCK_HEADERS.TOTAL_RM]: '#,##0.00'
  };
  setNumberFormats(sheet, headers, formats);

  // Styling the table
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length).createFilter();

  // Insert dropdowns
  const dropdowns = {
    [RESTOCK_HEADERS.STATUS]: ['In-stock', 'Out-of-stock'],
    [RESTOCK_HEADERS.PAYMENT_METHOD]: ['Online Banking', 'Cash', 'Credit/Debit Card', 'E-Wallet']
  };
  setDropdowns(sheet, header, dropdowns);

  // Apply date picker
  applyDateValidation(sheet, headers, RESTOCK_HEADERS.ARRIVAL_DATE);
  applyDateValidation(sheet, headers, RESTOCK_HEADERS.DATE);
  applyDateValidation(sheet, headers, RESTOCK_HEADERS.PAYMENT_DATE);

  // Insert shared dropdowns
  SpreadsheetApp.setActiveSheet(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings"));
  insertRestockSharedDropdowns();
  setSharedDropdownTrigger('insertRestockSharedDropdowns');
  SpreadsheetApp.setActiveSheet(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName));
}

function createEmployeeTable() {
  // Insert new sheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet();
  var sheetName = EMPLOYEE_HEADERS.SHEET_NAME;

  // Rename sheet
  sheet.setName(sheetName);

  // Define the table headers
  var headers = [
    EMPLOYEE_HEADERS.EMPLOYEE_ID,
    EMPLOYEE_HEADERS.NAME,
    EMPLOYEE_HEADERS.BRANCH,
    EMPLOYEE_HEADERS.STATUS,
    EMPLOYEE_HEADERS.JOB_TITLE,
    EMPLOYEE_HEADERS.DEPARTMENT,
    EMPLOYEE_HEADERS.SALARY_RM,
    EMPLOYEE_HEADERS.ADDRESS,
    EMPLOYEE_HEADERS.CONTACT_NUMBER,
    EMPLOYEE_HEADERS.EMAIL,
    EMPLOYEE_HEADERS.BIRTH_DATE,
    EMPLOYEE_HEADERS.JOINING_DATE,
    EMPLOYEE_HEADERS.EXIT_DATE,
    EMPLOYEE_HEADERS.BANK_ACCOUNT,
    EMPLOYEE_HEADERS.EPF,
    EMPLOYEE_HEADERS.SOCSO
  ];

  // Set the headers in the first row
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Format the headers
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#d9d9d9');

  // Set data types
  const formats = {
    [EMPLOYEE_HEADERS.EXIT_DATE]: 'yyyy-mm-dd',
    [EMPLOYEE_HEADERS.JOINING_DATE]: 'yyyy-mm-dd',
    [EMPLOYEE_HEADERS.BIRTH_DATE]: 'yyyy-mm-dd',
    [EMPLOYEE_HEADERS.SALARY_RM]: '#,##0.00',
    [EMPLOYEE_HEADERS.BANK_ACCOUNT]: '##0',
    [EMPLOYEE_HEADERS.SOCSO]: '##0',
    [EMPLOYEE_HEADERS.EPF]: '##0',
  };
  setNumberFormats(sheet, headers, formats);
  
  // Style the table
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length).createFilter();

  // Apply date picker
  applyDateValidation(sheet, headers, EMPLOYEE_HEADERS.EXIT_DATE);
  applyDateValidation(sheet, headers, EMPLOYEE_HEADERS.JOINING_DATE);
  applyDateValidation(sheet, headers, EMPLOYEE_HEADERS.BIRTH_DATE);

  // Insert dropdowns
  const dropdowns = {
    [EMPLOYEE_HEADERS.STATUS]: ['Active', 'Inactive'],
    [EMPLOYEE_HEADERS.JOB_TITLE]: ['Manager', 'Assistant', 'Sales'],
    [EMPLOYEE_HEADERS.DEPARTMENT]: ['HR', 'Finance', 'Marketing']
  };
  setDropdowns(sheet, headers, dropdowns);

  // Insert shared dropdown
  SpreadsheetApp.setActiveSheet(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings"));
  insertEmployeeSharedDropdowns();
  setSharedDropdownTrigger('insertEmployeeSharedDropdowns');
  SpreadsheetApp.setActiveSheet(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName));
}

function createInventoryTable() {
  // Insert new sheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet();
  var sheetName = INVENTORY_HEADERS.SHEET_NAME;

  // Rename sheet
  sheet.setName(sheetName);

  // Define the table headers
  var headers = [
    INVENTORY_HEADERS.PRODUCT_ID,
    INVENTORY_HEADERS.NAME,
    INVENTORY_HEADERS.CATEGORY,
    INVENTORY_HEADERS.QUANTITY,
    INVENTORY_HEADERS.SELLING_PRICE_RM,
    INVENTORY_HEADERS.UNIT_COST_RM,
    INVENTORY_HEADERS.TOTAL_SOLD,
    INVENTORY_HEADERS.LAST_RESTOCK,
    INVENTORY_HEADERS.NOTES
  ];

  // Set the headers in the first row
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Format the headers
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#d9d9d9');

  // Set data types
  const formats = {
    [INVENTORY_HEADERS.LAST_RESTOCK]: 'yyyy-mm-dd',
    [INVENTORY_HEADERS.UNIT_COST_RM]: '#,##0.00',
    [INVENTORY_HEADERS.SELLING_PRICE_RM]: '#,##0.00',
    [INVENTORY_HEADERS.QUANTITY]: '##0',
    [INVENTORY_HEADERS.TOTAL_SOLD]: '##0',
  };
  setNumberFormats(sheet, headers, formats);

  // Style the table
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length).createFilter();

  // Insert categories dropdown
  const dropdowns = {
    [INVENTORY_HEADERS.CATEGORY]: ['Electronics', 'Furniture', 'Clothing']
  };
  setDropdowns(sheet, headers, dropdowns)

  // Apply date picker
  applyDateValidation(sheet, headers, INVENTORY_HEADERS.LAST_RESTOCK);
}

function createSupplierTable() {
  // Insert new sheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet();
  var sheetName = SUPPLIER_HEADERS.SHEET_NAME;

  // Rename sheet
  sheet.setName(sheetName);

  // Define the table headers
  var headers = [
    SUPPLIER_HEADERS.SUPPLIER_ID,
    SUPPLIER_HEADERS.NAME,
    SUPPLIER_HEADERS.CATEGORY,
    SUPPLIER_HEADERS.CONTACT_NUMBER,
    SUPPLIER_HEADERS.EMAIL,
    SUPPLIER_HEADERS.WEBSITE,
    SUPPLIER_HEADERS.RATING,
    SUPPLIER_HEADERS.NOTES
  ];

  // Set the headers in the first row
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Format the headers
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#d9d9d9');

  // Style the table
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length).createFilter();

  // Insert categories dropdown
  const dropdowns = {
    [SUPPLIER_HEADERS.CATEGORY]: ['Local', 'International'],
    [SUPPLIER_HEADERS.RATING]: ['Excellent', 'Good', 'Average', 'Poor', 'N/A'],
  };
  setDropdowns(sheet, headers, dropdowns);
}

function createCustomerTable() {
  // Insert new sheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet();
  var sheetName = CUSTOMER_HEADERS.SHEET_NAME;

  // Rename sheet
  sheet.setName(sheetName);

  // Define the table headers
  var headers = [
    CUSTOMER_HEADERS.CUSTOMER_ID,
    CUSTOMER_HEADERS.NAME,
    CUSTOMER_HEADERS.CONTACT_NUMBER,
    CUSTOMER_HEADERS.EMAIL,
    CUSTOMER_HEADERS.RECENT_ORDER,
    CUSTOMER_HEADERS.SALE_COUNT,
    CUSTOMER_HEADERS.TOTAL_SALES_RM,
    CUSTOMER_HEADERS.DAYS_SINCE_LAST_ORDER,
    CUSTOMER_HEADERS.R_SCORE,
    CUSTOMER_HEADERS.F_SCORE,
    CUSTOMER_HEADERS.M_SCORE,
    CUSTOMER_HEADERS.SEGMENT
  ];

  // Set the headers in the first row
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Format the headers
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#d9d9d9');

  // Set data types
  const formats = {
    [CUSTOMER_HEADERS.RECENT_ORDER]: 'yyyy-mm-dd',
    [CUSTOMER_HEADERS.TOTAL_SALES_RM]: '#,##0.00',
    [CUSTOMER_HEADERS.SALE_COUNT]: '##0',
  }
  setNumberFormats(sheet, headers, formats)
  
  // Apply date picker
  applyDateValidation(sheet, headers, CUSTOMER_HEADERS.RECENT_ORDER);
  
  // Style the table
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length).createFilter();

  // Apply conditional formatting
  var rColumn = columnToLetter(headers.indexOf(CUSTOMER_HEADERS.R_SCORE)+1);
  var mColumn = columnToLetter(headers.indexOf(CUSTOMER_HEADERS.M_SCORE)+1);
  var conditionalFormatRuleRange = rColumn + '2:' + mColumn;

  conditionalFormatRules = sheet.getConditionalFormatRules();
  conditionalFormatRules.splice(conditionalFormatRules.length - 1, 1, SpreadsheetApp.newConditionalFormatRule()
    .setRanges([sheet.getRange(conditionalFormatRuleRange)])
    .setGradientMinpoint('#FAD2CF')
    .setGradientMidpointWithValue('#FEEFC3', SpreadsheetApp.InterpolationType.PERCENTILE, '50')
    .setGradientMaxpoint('#CEEAD6')
    .build()
  );
  sheet.setConditionalFormatRules(conditionalFormatRules);
}

function createLiabilitiesTable() {
  // Insert new sheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet();
  var sheetName = LIABILITY_HEADERS.SHEET_NAME;

  // Rename sheet
  sheet.setName(sheetName);

  // Define the table headers
  var headers = [
    LIABILITY_HEADERS.LIABILITY_ID,
    LIABILITY_HEADERS.NAME,
    LIABILITY_HEADERS.ACCOUNT,
    LIABILITY_HEADERS.CREDITOR,
    LIABILITY_HEADERS.CATEGORY,
    LIABILITY_HEADERS.INITIAL_AMOUNT_RM,
    LIABILITY_HEADERS.CURRENT_BALANCE_RM,
    LIABILITY_HEADERS.INTEREST_RATE_PERCENT,
    LIABILITY_HEADERS.FREQUENCY,
    LIABILITY_HEADERS.MIN_PAYMENT_RM,
    LIABILITY_HEADERS.DUE_DATE,
    LIABILITY_HEADERS.LAST_PAYMENT,
    LIABILITY_HEADERS.LAST_PAYMENT_AMOUNT_RM,
    LIABILITY_HEADERS.NEXT_PAYMENT_AMOUNT_RM,
    LIABILITY_HEADERS.START_DATE,
    LIABILITY_HEADERS.MATURITY_DATE,
    LIABILITY_HEADERS.CURRENCY,
    LIABILITY_HEADERS.TAXABLE,
    LIABILITY_HEADERS.NOTES,
    LIABILITY_HEADERS.STATUS
  ];

  // Set the headers in the first row
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Format the headers
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#d9d9d9');

  // Set data types
  const formats = {
    [LIABILITY_HEADERS.DUE_DATE]: 'yyyy-mm-dd',
    [LIABILITY_HEADERS.LAST_PAYMENT]: 'yyyy-mm-dd',
    [LIABILITY_HEADERS.START_DATE]: 'yyyy-mm-dd',
    [LIABILITY_HEADERS.MATURITY_DATE]: 'yyyy-mm-dd',
    [LIABILITY_HEADERS.CURRENT_BALANCE_RM]: '#,##0.00',
    [LIABILITY_HEADERS.INITIAL_AMOUNT_RM]: '#,##0.00',
    [LIABILITY_HEADERS.NEXT_PAYMENT_AMOUNT_RM]: '#,##0.00',
    [LIABILITY_HEADERS.LAST_PAYMENT_AMOUNT_RM]: '#,##0.00',
    [LIABILITY_HEADERS.MIN_PAYMENT_RM]: '#,##0.00',
    [LIABILITY_HEADERS.INTEREST_RATE_PERCENT]: '0.00%',
  }
  setNumberFormats(sheet, headers, formats);

  // Style the table
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length).createFilter();

  // Insert categories dropdown
  const dropdowns = {
    [LIABILITY_HEADERS.CATEGORY]: ['Borrowings', 'Lease liabilities', 'Payables and accruals', 'Contract', 'Tax', 'Others'],
    [LIABILITY_HEADERS.FREQUENCY]: ['Daily', 'Weekly', 'Monthly', 'Quaterly', 'Yearly'],
    [LIABILITY_HEADERS.STATUS]: ['Pending', 'On-going', 'Completed']
  }
  setDropdowns(sheet, headers, dropdowns);

  // Apply date picker
  applyDateValidation(sheet, header, LIABILITY_HEADERS.DUE_DATE);
  applyDateValidation(sheet, header, LIABILITY_HEADERS.MATURITY_DATE);
  applyDateValidation(sheet, header, LIABILITY_HEADERS.START_DATE);
  applyDateValidation(sheet, header, LIABILITY_HEADERS.LAST_PAYMENT);
}

function createAssetsTable() {
  // Insert new sheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet();
  sheetName = ASSET_HEADERS.SHEET_NAME;

  // Rename sheet
  sheet.setName(sheetName);

  // Define the table headers
  var headers = [
    ASSET_HEADERS.ASSET_ID,
    ASSET_HEADERS.NAME,
    ASSET_HEADERS.CATEGORY,
    ASSET_HEADERS.PURCHASE_DATE,
    ASSET_HEADERS.COST_RM,
    ASSET_HEADERS.CURRENT_VALUE_RM,
    ASSET_HEADERS.ANNUAL_DEPRECIATION_RATE_PERCENT,
    ASSET_HEADERS.SALVAGE_VALUE,
    ASSET_HEADERS.USEFUL_YEARS
  ];

  // Set the headers in the first row
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Format the headers
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#d9d9d9');

  // Set data types
  const formats = {
    [ASSET_HEADERS.PURCHASE_DATE]: 'yyyy-mm-dd',
    [ASSET_HEADERS.COST_RM]: '#,##0.00',
    [ASSET_HEADERS.CURRENT_VALUE_RM]: '#,##0.00',
    [ASSET_HEADERS.SALVAGE_VALUE]: '#,##0.00',
    [ASSET_HEADERS.ANNUAL_DEPRECIATION_RATE_PERCENT]: '0.00%',
  };
  setNumberFormats(sheet, headers, formats);

  // Style the table
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length).createFilter();

  // Insert category dropdown
  const dropdowns = {
    [ASSET_HEADERS.CATEGORY]: ['Property, plant and equipment', 'Right-of-use assets', 'Intangible assets', 'Investments', 'Others', 'Deferred tax assets', 'Inventory', 'Contract', 'Receivables, deposits and prepayments', 'Cash and cash equivalents']
  };
  setDropdowns(sheet, headers, dropdowns);

  // Apply date picker
  applyDateValidation(sheet, headers, ASSET_HEADERS.PURCHASE_DATE);
}

function createExpensesTable() {
  // Insert new sheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet();
  var sheetName = EXPENSE_HEADERS.SHEET_NAME;

  // Rename sheet
  sheet.setName(sheetName);

  // Define the table headers
  var headers = [
    EXPENSE_HEADERS.EXPENSE_ID,
    EXPENSE_HEADERS.NAME,
    EXPENSE_HEADERS.DATE,
    EXPENSE_HEADERS.CATEGORY,
    EXPENSE_HEADERS.DESCRIPTION,
    EXPENSE_HEADERS.OPERATING_EXPENSE,
    EXPENSE_HEADERS.CAPITAL_EXPENDITURE,
    EXPENSE_HEADERS.TOTAL_RM
  ];

  // Set the headers in the first row
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Format the headers
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#d9d9d9');

  // Set data types
  const formats = {
    [EXPENSE_HEADERS.DATE]: 'yyyy-mm-dd',
    [EXPENSE_HEADERS.TOTAL_RM]: '#,##0.00',
  };
  setNumberFormats(sheet, headers, formats);

  // Style the table
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length).createFilter();

  // Apply date picker
  applyDateValidation(sheet, headers, EXPENSE_HEADERS.DATE);

  // Insert dropdowns
  const dropdowns = {
    [EXPENSE_HEADERS.CATEGORY]: ['Loan Interests', 'Salaries', 'Rentals', 'Contracts', 'Commisions', 'Bad Debts', 'Transports', 'Repairs/Maintenances', 'Promotions/Advertisements', 'Tax', 'Cost of Sales', 'Fulfillment', 'Technology', 'Administrative', 'Others'],
    [EXPENSE_HEADERS.CAPITAL_EXPENDITURE]: ['Yes', 'No'],
    [EXPENSE_HEADERS.OPERATING_EXPENSE]: ['Yes', 'No'],
  };
  setDropdowns(sheet, headers, dropdowns);
}

function createIncomeTable() {
  // Insert new sheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet();
  var sheetName = INCOME_HEADERS.SHEET_NAME;

  // Rename sheet
  sheet.setName(sheetName);

  // Define the table headers
  var headers = [
    INCOME_HEADERS.INCOME_ID,
    INCOME_HEADERS.NAME,
    INCOME_HEADERS.DATE,
    INCOME_HEADERS.CATEGORY,
    INCOME_HEADERS.DESCRIPTION,
    INCOME_HEADERS.TOTAL_RM
  ];

  // Set the headers in the first row
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Format the headers
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#d9d9d9');

  // Set data types
  const formats = {
    [INCOME_HEADERS.DATE]: 'yyyy-mm-dd',
    [INCOME_HEADERS.TOTAL_RM]: '#,##0.00',
  };
  setNumberFormats(sheet, headers, formats);
  
  // Style the table
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length).createFilter();

  // Apply date picker
  applyDateValidation(sheet, headers, INCOME_HEADERS.DATE);

  // Insert dropdowns
  const dropdowns = {
    [INCOME_HEADERS.CATEGORY]: ['Sales', 'Opening Stock', 'Purchases', 'Closing Stock', 'Cost of Sales', 'Other Businesses', 'Dividends', 'Interests and Discounts', 'Rents, Royalties, and Premiums', 'Other Income', 'Depreciation']
  };
  setDropdowns(sheet, headers, dropdowns);
}

function createCashflowMetricsTable() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetCount = ss.getSheets().length;

  let sheet = ss.getSheetByName(CASH_FLOW_HEADERS.SHEET_NAME);
  if (sheet) {
    sheet.clear();
    removeAllCharts(sheet);
  } else {
    sheet = ss.insertSheet();
  }
  var sheetName = CASH_FLOW_HEADERS.SHEET_NAME;

  // Rename sheet
  sheet.setName(sheetName);

  ss.setActiveSheet(sheet);

  // Style sheet
  sheet.setHiddenGridlines(true);

  sheet.getRange('A1').setValue('Year: ').setFontWeight('italic');

  // Assume monthly data is available in these sheets and metrics are calculated using the necessary columns.
  var currentDate = new Date();
  var year = currentDate.getFullYear();
  
  var yearRange = [];
  for (var y = 2014; y <= year; y++) {
    yearRange.push(y);
  }
  insertDropdown(sheet, 'B1', yearRange);
  sheet.getRange('B1').setValue(year);

  // Define the table headers
  var headers = [
    CASH_FLOW_HEADERS.MONTH,
    CASH_FLOW_HEADERS.OPERATING_CASH_FLOW_RM,
    CASH_FLOW_HEADERS.FREE_CASH_FLOW_RM,
    CASH_FLOW_HEADERS.DAYS_PAYABLE_OUTSTANDING,
    CASH_FLOW_HEADERS.DAYS_SALES_OUTSTANDING,
    CASH_FLOW_HEADERS.DAYS_INVENTORY_OUTSTANDING,
    CASH_FLOW_HEADERS.CASH_CONVERSION_CYCLE
  ];

  // Set the headers in the first row
  sheet.getRange(2, 1, 1, headers.length).setValues([headers]);
  
  // Format the headers
  var headerRange = sheet.getRange(2, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#d9d9d9');
  
  // Set data types for numerical columns
  var secondColumn = columnToLetter(headers.indexOf(CASH_FLOW_HEADERS.OPERATING_CASH_FLOW_RM)+1);
  var lastColumn = columnToLetter(headers.indexOf(CASH_FLOW_HEADERS.CASH_CONVERSION_CYCLE)+1);
  var numberFormatRange = secondColumn + '3:' + lastColumn + '14';
  sheet.getRange(numberFormatRange).setNumberFormat('#,##0.00');

  var months = [];
  for (var month = 0; month < 12; month++) {
    var monthName = new Date(year, month, 1).toLocaleString('default', { month: 'long' });
    months.push([monthName]);
  }
  sheet.getRange("A3:A14").setValues(months);
  
  var operatingCashflowRange = sheet.getRange('B3:B14');
  var freeCashflowRange = sheet.getRange('C3:C14');
  var dpoRange = sheet.getRange('D3:D14');
  var dsoRange = sheet.getRange('E3:E14');
  var dioRange = sheet.getRange('F3:F14');
  var cccRange = sheet.getRange('G3');

  operatingCashflowRange.setFormula(
    "=ARRAYFORMULA(calculateOperatingCashFlow($B$1, ROW()-2))"
  );

  freeCashflowRange.setFormula(
    "=ARRAYFORMULA(calculateFreeCashFlowForMonth($B$1, ROW()-2))"
  );

  dpoRange.setFormula(
    "=ARRAYFORMULA(calculateDPO($B$1, ROW()-2))"
  );

  dsoRange.setFormula(
    "=ARRAYFORMULA(calculateDSO($B$1, ROW()-2))"
  );

  dioRange.setFormula(
    "=ARRAYFORMULA(calculateDIO($B$1, ROW()-2))"
  );

  cccRange.setFormula(
    "=ARRAYFORMULA(F3:F14 + E3:E14 - D3:D14)"
  );

  sheet.autoResizeColumns(1, headers.length);
}
