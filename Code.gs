function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Manage Data').addSubMenu(
      ui.createMenu('Generate Tables')
      .addItem('Generate All Tables', 'generateTables')
      .addSeparator()
      .addItem('Generate Sales Table', 'createSalesTable')
      .addItem('Generate Restock Table', 'createRestockTable')
      .addItem('Generate Employee Table', 'createEmployeeTable')
      .addItem('Generate Inventory Table', 'createInventoryTable')
      .addItem('Generate Supplier Table', 'createSupplierTable')
      .addItem('Generate Customer Table', 'createCustomerTable')
      .addItem('Generate Liabilities Table', 'createLiabilitiesTable')
      .addItem('Generate Assets Table', 'createAssetsTable')
      .addItem('Generate Expenses Table', 'createExpensesTable')
      .addItem('Generate Income Table', 'createIncomeTable')
    ).addSubMenu(
      ui.createMenu('Add Data')
      .addItem('Add Sales Data', 'showSalesForm')
      .addItem('Add Restock Data', 'showRestockForm')
      .addItem('Add Inventory Data', 'showInventoryForm')
      .addItem('Add Employee Data', 'showEmployeeForm')
      .addItem('Add Supplier Data', 'showSupplierForm')
      .addItem('Add Customer Data', 'showCustomerForm')
      .addItem('Add Liabilities Data', 'showLiabilitiesForm')
      .addItem('Add Assets Data', 'showAssetsForm')
      .addItem('Add Income Data', 'showIncomeForm')
      .addItem('Add Expenses Data', 'showExpensesForm') 
    ).addSubMenu(
      ui.createMenu('Generate Report')
      .addItem('Generate Sales Report', 'generateSalesReport')
    )
    .addToUi();
  ui
    .createMenu('Dashboard')
    .addItem('Generate All Dashboards', 'generateDashboards')
    .addSeparator()
    .addItem('Generate Cashflow Dashboard', 'generateCashFlowDashboard')
    .addItem('Generate Customer Dashboard', 'generateCustomerDashboard')
    .addItem('Generate Sales Dashboard', 'generateSalesDashboard')
    .addItem('Generate Inventory Dashboard', 'generateInventoryDashboard')
    .addToUi();
}

function generateTables() {
  createSettingsTable(); // Generate options table
  createSalesTable(); // Generate sales table
  createRestockTable(); // Generate restock table
  createEmployeeTable(); // Generate employee table
  createInventoryTable(); // Generate inventory table
  createSupplierTable(); // Generate supplier table
  createCustomerTable(); // Generate customer table
  createLiabilitiesTable(); // Generate liabilities table
  createAssetsTable(); // Generate assets table
  createExpensesTable(); // Generate expenses table
  createIncomeTable(); // Generate income table
}

function showSalesForm() {
  var html = HtmlService.createHtmlOutputFromFile("SaleForm").setTitle("Sales Form");
  SpreadsheetApp.getUi().showSidebar(html);
}

function showRestockForm() {
  var html = HtmlService.createHtmlOutputFromFile('RestockForm').setTitle("Restock Form");
  SpreadsheetApp.getUi().showSidebar(html);
}

function showInventoryForm() {
  var html = HtmlService.createHtmlOutputFromFile('InventoryForm').setTitle("Inventory Form");
  SpreadsheetApp.getUi().showSidebar(html);
}

function showEmployeeForm() {
  var html = HtmlService.createHtmlOutputFromFile('EmployeeForm').setTitle("Emoloyee Form");
  SpreadsheetApp.getUi().showSidebar(html);
}

function showSupplierForm() {
  var html = HtmlService.createHtmlOutputFromFile('SupplierForm').setTitle("Supplier Form");
  SpreadsheetApp.getUi().showSidebar(html);
}

function showCustomerForm() {
  var html = HtmlService.createHtmlOutputFromFile('CustomerForm').setTitle("Customer Form");
  SpreadsheetApp.getUi().showSidebar(html);
}

function showLiabilitiesForm() {
  var html = HtmlService.createHtmlOutputFromFile('LiabilitiesForm').setTitle("Liabilities Form");
  SpreadsheetApp.getUi().showSidebar(html);
}

function showAssetsForm() {
  var html = HtmlService.createHtmlOutputFromFile('AssetsForm').setTitle("Assets Form");
  SpreadsheetApp.getUi().showSidebar(html);
}

function showExpensesForm() {
  var html = HtmlService.createHtmlOutputFromFile('ExpensesForm').setTitle("Expenses Form");
  SpreadsheetApp.getUi().showSidebar(html);
}

function showIncomeForm() {
  var html = HtmlService.createHtmlOutputFromFile('IncomeForm').setTitle("Income Form");
  SpreadsheetApp.getUi().showSidebar(html);
}

function generateDashboards() {
  generateCashFlowDashboard();
  generateCustomerDashboard();
  generateSalesDashboard();
  generateInventoryDashboard();
}

