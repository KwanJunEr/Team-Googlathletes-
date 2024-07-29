function generateCashFlowDashboard() {
  createCashflowMetricsTable();
  generateOperatingCashflowLineChart();
  generateFreeCashflowLineChart();
  generateCCCLineChart();
  generateDPOColumnChart();
  generateDSOColumnChart();
  generateDIOColumnChart();
}

// OPERATING CASH FLOW
function calculateOperatingCashFlow(year, month) {
  var operatingIncome = getOperatingIncome(year, month);
  var depreciation = getDepreciation(year, month);
  var taxes = getTaxes(year, month);
  var changeInWorkingCapital = getChangeInWorkingCapitalForMonth(year, month);

  return operatingIncome - depreciation - taxes + changeInWorkingCapital;
}

function getOperatingIncome(year, month) {
  // Calculate Total Revenue from the sales table
  var totalRevenue = getTotalRevenueForMonth(year, month);
  
  // Calculate Total Operating Expenses from the 'expenses' table
  var totalOperatingExpenses = getTotalOperatingExpensesForMonth(year, month);

  // Calculate Operating Income
  return totalRevenue - totalOperatingExpenses;
}

function getTotalRevenueForMonth(year, month) {
  var salesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SALES_HEADERS.SHEET_NAME);
  
  var range = salesSheet.getRange('G2:Q');
  var values = range.getValues();
  var totalRevenue = 0;

  for (var i = 0; i < values.length; i++) {
    var date = new Date(values[i][0]);
    if (date.getFullYear() == year && date.getMonth() == month-1) {
      totalRevenue += parseFloat(values[i][10]) || 0;
    }
  }
  return totalRevenue;
}

function getTotalOperatingExpensesForMonth(year, month) {
  var expensesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(EXPENSE_HEADERS.SHEET_NAME);
  var range = expensesSheet.getRange('C2:H');
  var values = range.getValues();
  var totalOperatingExpenses = 0;
  
  for (var i = 0; i < values.length; i++) {
    var date = new Date(values[i][0]);
    var isOperatingExpense = values[i][3];
    
    if (date.getFullYear() == year && date.getMonth() == month-1 && isOperatingExpense == 'Yes') {
      totalOperatingExpenses += parseFloat(values[i][5]) || 0;
    }
  }
  return totalOperatingExpenses;
}

function getDepreciation(year, month) {
  assetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ASSET_HEADERS.SHEET_NAME);

  var range = assetSheet.getRange('D2:G'); // Adjust this range if needed
  var values = range.getValues();
  var totalDepreciation = 0;

  for (var i = 0; i < values.length; i++) {
    var purchaseDate = new Date(values[i][0]);
    var currentYear = year;
    var currentMonth = month; // JavaScript months are 0-based

    if (purchaseDate > new Date(currentYear, currentMonth + 1, 0)) {
      continue; // Skip if the asset was purchased after the given month
    }

    var assetValue = values[i][1];
    var depreciationRate = values[i][3];

    var purchaseYear = purchaseDate.getFullYear();
    var purchaseMonth = purchaseDate.getMonth();

    var monthsDiff = (currentYear - purchaseYear) * 12 + (currentMonth - purchaseMonth);


    if (monthsDiff < 0) {
      continue; // Skip if the asset was purchased in the future
    }

    var monthlyDepreciationRate = Math.pow((1 - depreciationRate), 1/12);

    for (var m = 1; m <= monthsDiff; m++) {
      var currentMonthValue = assetValue * (monthlyDepreciationRate);
      if (m == monthsDiff) {
        totalDepreciation += assetValue - currentMonthValue;
      }
      assetValue = currentMonthValue;
    }
  }
  return totalDepreciation;
}

function getTaxes(year, month) {
  var expensesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(EXPENSE_HEADERS.SHEET_NAME);
  var range = expensesSheet.getRange('C2:H');
  var values = range.getValues();
  var totalTax = 0;
  
  for (var i = 0; i < values.length; i++) {
    var date = new Date(values[i][0]);
    var isTax = values[i][1];
    
    if (date.getFullYear() == year && date.getMonth() == month-1 && isTax == 'Tax') {
      totalTax += parseFloat(values[i][5]) || 0;
    }
  }
  return totalTax;
}

function getChangeInWorkingCapitalForMonth(year, month) {
  var currentAsset_1 = getCurrentAssetsForMonth(year, month);
  if (currentAsset_1 == 0) {
    currentAsset_1 = 1;
  }
  
  var currentAsset_2 = getCurrentAssetsForMonth(year, month-1);
  if (currentAsset_2 == 0) {
    currentAsset_2 = 1;
  }

  var workingCapital_1 = getCurrentLiabilitiesForMonth(year, month) / currentAsset_1;
  var workingCapital_2 = getCurrentLiabilitiesForMonth(year, month-1) / currentAsset_2;
  return workingCapital_1 - workingCapital_2;
}

function getCurrentLiabilitiesForMonth(year, month) {
  var expensesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LIABILITY_HEADERS.SHEET_NAME);
  var range = expensesSheet.getRange('K2:N');
  var values = range.getValues();
  
  var totalLiabilities = 0;

  for (var i = 0; i < values.length; i++) {
    var dueDate = new Date(values[i][0]);
    var lastPaymentDate = new Date(values[i][1]);
    var lastPaymentAmount = values[i][2];
    var nextPaymentAmount = values[i][3];
    
    if (lastPaymentDate.getMonth() == month -1 && lastPaymentDate.getFullYear() == year) {
      totalLiabilities += parseFloat(lastPaymentAmount);
    } else if (dueDate.getMonth() == month -1 && dueDate.getFullYear() == year) {
      totalLiabilities += parseFloat(nextPaymentAmount);
    }
  }
  return totalLiabilities;
}

function getCurrentAssetsForMonth(year, month) {
  var expensesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ASSET_HEADERS.SHEET_NAME);
  var range = expensesSheet.getRange('C2:F');
  var values = range.getValues();
  
  var totalAssets = 0;
  filterDate = new Date(year, month, 0);

  for (var i = 0; i < values.length; i++) {
    var category = values[i][0];
    var purchaseDate = new Date(values[i][1]);

    if (purchaseDate <= filterDate && category !== 'Cash and cash equivalents') {
      var currentValue = values[i][3];
      totalAssets += currentValue;
    }
  }
  return totalAssets;
}


// FREE CASH FLOW
function calculateFreeCashFlowForMonth(year, month) {
  return calculateOperatingCashFlow(year, month) - getCapitalExpendituresForMonth(year, month);
}

function getCapitalExpendituresForMonth(year, month) {
  var expensesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(EXPENSE_HEADERS.SHEET_NAME);
  var range = expensesSheet.getRange('C2:H');
  var values = range.getValues();
  var totalCapitalExpenditure = 0;
  
  for (var i = 0; i < values.length; i++) {
    var date = new Date(values[i][0]);
    var isCapitalExpenditure = values[i][4];
    
    if (date.getFullYear() == year && date.getMonth() == month-1 && isCapitalExpenditure == 'Yes') {
      totalCapitalExpenditure += parseFloat(values[i][5]) || 0;
    }
  }
  return totalCapitalExpenditure;
}


// Days Payable Outstanding (DPO)
function calculateDPO(year, month) {
  var averageAccountPay = getAverageAccountsPayable(year, month);
  var costOfGoodsSold = getCostOfGoodsSold(year, month);

  // Handle divide by zero error
  if (costOfGoodsSold === 0) {
    return 0;
  }
  return (averageAccountPay / costOfGoodsSold) * new Date(year, month, 0).getDate();
}

function getAverageAccountsPayable(year, month) {
  var salesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SALES_HEADERS.SHEET_NAME);
  var range = salesSheet.getRange('E2:Q');
  var values = range.getValues();

  var totalAccountsPayableStart = 0;
  var totalAccountsPayableEnd = 0;

  var startPeriod = new Date(year, month-1, 1);
  var endPeriod = new Date(year, month, 0);

  // Start:
  for (var i = 0; i < values.length; i++) {
    var date = new Date(values[i][0]);
    var paymentDate = new Date(values[i][2]);
    if (date < startPeriod) {
      if (isNaN(paymentDate)) {
        totalAccountsPayableStart += parseFloat(values[i][12]) || 0;
      }
      else if (paymentDate > startPeriod) {
        totalAccountsPayableStart += parseFloat(values[i][12]) || 0;
      }
    }
  }

  // End:
  for (var i = 0; i < values.length; i++) {
    var date = new Date(values[i][0]);
    var paymentDate = new Date(values[i][2]);
    if (date <= endPeriod) {
      if (isNaN(paymentDate)) {
        totalAccountsPayableEnd += parseFloat(values[i][12]) || 0;
      }
      else if (paymentDate > endPeriod) {
        totalAccountsPayableEnd += parseFloat(values[i][12]) || 0;
      }
    }
  }

  return (totalAccountsPayableStart + totalAccountsPayableEnd) / 2;
}

function getCostOfGoodsSold(year, month) {
  var salesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SALES_HEADERS.SHEET_NAME);
  var range = salesSheet.getRange('G2:K');
  var values = range.getValues();

  var totalCostOfGoodsSold = 0;

  for (var i = 0; i < values.length; i++) {
    var date = new Date(values[i][0]);
    if (date.getFullYear() == year && date.getMonth() == month-1) {
      totalCostOfGoodsSold += parseFloat(values[i][4]);
    }
  }

  return totalCostOfGoodsSold
}


// Days Sales Outstanding (DSO) = (Average Accounts Payable / Total revenue) * days in period
function calculateDSO(year, month) {
  var averageAccountPayable = getAverageAccountsPayable(year, month);
  var totalRevenue = getTotalRevenueForMonth(year, month);
  
  // Handle divide by zero error
  if (totalRevenue === 0) {
    return 0;
  }

  return (averageAccountPayable / totalRevenue) * new Date(year, month, 0).getDate();
}


// Days Inventory Outstanding (DIO) = (Average inventory/Cost of goods sold) * days in period
function calculateDIO(year, month) {
  var averageInventory = getAverageInventoryCostForMonth(year, month);
  var costOfGoodsSold = getCostOfGoodsSold(year, month);

  if (costOfGoodsSold === 0) {
    return 0;
  }

  return (averageInventory / costOfGoodsSold) * new Date(year, month, 0).getDate();
}

function getAverageInventoryCostForMonth(year, month) {
  var salesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SALES_HEADERS.SHEET_NAME);
  var range = salesSheet.getRange('E2:N');
  var values = range.getValues();

  var totalInventoryCostStart = 0;
  var totalInventoryCostEnd = 0;

  var startPeriod = new Date(year, month-1, 1);
  var endPeriod = new Date(year, month, 0);

  var latestSaleDateBeforeStart = {};
  var latestSaleDateBeforeEnd = {};

  for (var i = 0; i < values.length; i++) {
    var saleDate = new Date(values[i][0]);
    var productId = values[i][5];

    if (saleDate < startPeriod) {
      var inventoryValue = [];

      if (!latestSaleDateBeforeStart[productId] || saleDate >= latestSaleDateBeforeStart[productId][0]) {
        var productStockLevel =  values[i][9];
        var productUnitCost =  values[i][6];
        var value = productUnitCost * productStockLevel;
        
        inventoryValue.push(saleDate);
        inventoryValue.push(productId);
        inventoryValue.push(value);

        latestSaleDateBeforeStart[productId] = inventoryValue;
      }
    }

    if (saleDate <= endPeriod) {
      var inventoryValue = [];

      if (!latestSaleDateBeforeEnd[productId] || saleDate >= latestSaleDateBeforeEnd[productId][0]) {
        var productStockLevel =  values[i][9];
        var productUnitCost =  values[i][6];
        var value = productUnitCost * productStockLevel;

        inventoryValue.push(saleDate);
        inventoryValue.push(productId);
        inventoryValue.push(value);

        latestSaleDateBeforeEnd[productId] = inventoryValue;
      }
    }
  }

  for (var [productId, inventoryValueStart] of Object.entries(latestSaleDateBeforeStart)) {
    totalInventoryCostStart += parseFloat(inventoryValueStart[2]);
  }

  // Calculate total inventory cost at the end period
  for (var [productId, inventoryValueEnd] of Object.entries(latestSaleDateBeforeEnd)) {
    totalInventoryCostEnd += parseFloat(inventoryValueEnd[2]);
  }

  return (totalInventoryCostStart + totalInventoryCostEnd) / 2;
}



function generateOperatingCashflowLineChart() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = spreadsheet.getSheetByName(CASH_FLOW_HEADERS.SHEET_NAME);

  chart = sheet.newChart()
  .asLineChart()
  .addRange(spreadsheet.getRange('A2:B14'))
  .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
  .setTransposeRowsAndColumns(false)
  .setNumHeaders(1)
  .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
  .setOption('bubble.stroke', '#000000')
  .setOption('useFirstColumnAsDomain', true)
  .setOption('isStacked', 'false')
  .setOption('title', 'Operating Cashflow over Time')
  .setOption('annotations.domain.textStyle.color', '#808080')
  .setOption('textStyle.color', '#000000')
  .setOption('legend.textStyle.color', '#1a1a1a')
  .setOption('titleTextStyle.color', '#202124')
  .setOption('titleTextStyle.alignment', 'center')
  .setOption('titleTextStyle.bold', true)
  .setOption('annotations.total.textStyle.color', '#808080')
  .setXAxisTitle('Months')
  .setOption('hAxis.textStyle.color', '#000000')
  .setOption('hAxis.textStyle.bold', true)
  .setOption('hAxis.titleTextStyle.color', '#000000')
  .setOption('hAxis.titleTextStyle.bold', true)
  .setYAxisTitle('Operating Cashflow (RM)')
  .setOption('vAxes.0.textStyle.color', '#000000')
  .setOption('vAxes.0.textStyle.bold', true)
  .setOption('vAxes.0.titleTextStyle.color', '#000000')
  .setOption('vAxes.0.titleTextStyle.bold', true)
  .setOption('series.0.pointSize', 7)
  .setOption('backgroundColor', '#D2E3FC')
  .setPosition(15, 1, 13, 14)
  .build();
  sheet.insertChart(chart);
};

function generateFreeCashflowLineChart() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = spreadsheet.getSheetByName(CASH_FLOW_HEADERS.SHEET_NAME);

  var chart = sheet.newChart()
  .asLineChart()
  .addRange(spreadsheet.getRange('A2:A14'))
  .addRange(spreadsheet.getRange('C2:C14'))
  .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
  .setTransposeRowsAndColumns(false)
  .setNumHeaders(1)
  .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
  .setOption('bubble.stroke', '#000000')
  .setOption('useFirstColumnAsDomain', true)
  .setOption('isStacked', 'false')
  .setOption('title', 'Free Cashflow over Time')
  .setOption('annotations.domain.textStyle.color', '#808080')
  .setOption('textStyle.color', '#000000')
  .setOption('legend.textStyle.color', '#1a1a1a')
  .setOption('titleTextStyle.color', '#202124')
  .setOption('titleTextStyle.alignment', 'center')
  .setOption('titleTextStyle.bold', true)
  .setOption('annotations.total.textStyle.color', '#808080')
  .setXAxisTitle('Months')
  .setOption('hAxis.textStyle.color', '#000000')
  .setOption('hAxis.titleTextStyle.color', '#000000')
  .setOption('hAxis.titleTextStyle.bold', true)
  .setOption('hAxis.textStyle.bold', true)
  .setOption('vAxes.0.minorGridlines.count', 2)
  .setYAxisTitle('Free Cashflow (RM)')
  .setOption('vAxes.0.textStyle.color', '#000000')
  .setOption('vAxes.0.titleTextStyle.color', '#000000')
  .setOption('vAxes.0.titleTextStyle.bold', true)
  .setOption('vAxis.textStyle.bold', true)
  .setOption('series.0.pointSize', 7)
  .setOption('backgroundColor', '#FAD2CF')
  .setPosition(15, 5, 24, 14)
  .build();
  sheet.insertChart(chart);
};

function generateCCCLineChart() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = spreadsheet.getSheetByName(CASH_FLOW_HEADERS.SHEET_NAME);

  var chart = sheet.newChart()
  .asLineChart()
  .addRange(spreadsheet.getRange('A2:A14'))
  .addRange(spreadsheet.getRange('G2:G14'))
  .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
  .setTransposeRowsAndColumns(false)
  .setNumHeaders(1)
  .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
  .setOption('bubble.stroke', '#000000')
  .setOption('useFirstColumnAsDomain', true)
  .setOption('isStacked', 'false')
  .setOption('title', 'Cash Conversion Cycle (CCC) over Time')
  .setOption('annotations.domain.textStyle.color', '#808080')
  .setOption('textStyle.color', '#000000')
  .setOption('legend.textStyle.color', '#1a1a1a')
  .setOption('titleTextStyle.color', '#202124')
  .setOption('titleTextStyle.alignment', 'center')
  .setOption('titleTextStyle.bold', true)
  .setOption('annotations.total.textStyle.color', '#808080')
  .setXAxisTitle('Months')
  .setOption('hAxis.textStyle.color', '#000000')
  .setOption('hAxis.titleTextStyle.color', '#000000')
  .setOption('hAxis.titleTextStyle.bold', true)
  .setOption('hAxis.textStyle.bold', true)
  .setOption('vAxes.0.minorGridlines.count', 2)
  .setYAxisTitle('CCC (Days)')
  .setOption('vAxes.0.textStyle.color', '#000000')
  .setOption('vAxes.0.titleTextStyle.color', '#000000')
  .setOption('vAxes.0.titleTextStyle.bold', true)
  .setOption('vAxis.textStyle.bold', true)
  .setOption('series.0.pointSize', 7)
  .setOption('backgroundColor', '#AFE3D6')
  .setPosition(15, 8, 59, 14)
  .build();
  sheet.insertChart(chart);
};

function generateDPOColumnChart() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = spreadsheet.getSheetByName(CASH_FLOW_HEADERS.SHEET_NAME);

  var chart = sheet.newChart()
  .asColumnChart()
  .addRange(spreadsheet.getRange('A2:A14'))
  .addRange(spreadsheet.getRange('D2:D14'))
  .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
  .setTransposeRowsAndColumns(false)
  .setNumHeaders(1)
  .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
  .setOption('bubble.stroke', '#000000')
  .setOption('useFirstColumnAsDomain', true)
  .setOption('isStacked', 'false')
  .setOption('title', 'Days Payable Outstanding (DPO) over Time')
  .setOption('annotations.domain.textStyle.color', '#808080')
  .setOption('textStyle.color', '#000000')
  .setOption('legend.textStyle.color', '#1a1a1a')
  .setOption('titleTextStyle.color', '#202124')
  .setOption('titleTextStyle.alignment', 'center')
  .setOption('titleTextStyle.bold', true)
  .setOption('annotations.total.textStyle.color', '#808080')
  .setXAxisTitle('Months')
  .setOption('hAxis.textStyle.color', '#000000')
  .setOption('hAxis.textStyle.bold', true)
  .setOption('hAxis.titleTextStyle.color', '#000000')
  .setOption('hAxis.titleTextStyle.bold', true)
  .setOption('vAxes.0.minorGridlines.count', 2)
  .setYAxisTitle('DPO (Days)')
  .setOption('vAxes.0.textStyle.color', '#000000')
  .setOption('vAxes.0.textStyle.bold', true)
  .setOption('vAxes.0.titleTextStyle.color', '#000000')
  .setOption('vAxes.0.titleTextStyle.bold', true)
  .setOption('backgroundColor', '#FEEFC3')
  .setPosition(33, 8, 59, 18)
  .build();
  sheet.insertChart(chart);
};

function generateDSOColumnChart() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = spreadsheet.getSheetByName(CASH_FLOW_HEADERS.SHEET_NAME);

  var chart = sheet.newChart()
  .asColumnChart()
  .addRange(spreadsheet.getRange('A2:A14'))
  .addRange(spreadsheet.getRange('E2:E14'))
  .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
  .setTransposeRowsAndColumns(false)
  .setNumHeaders(1)
  .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
  .setOption('bubble.stroke', '#000000')
  .setOption('useFirstColumnAsDomain', true)
  .setOption('isStacked', 'false')
  .setOption('title', 'Days Sales Outstanding (DSO) over Time')
  .setOption('annotations.domain.textStyle.color', '#808080')
  .setOption('textStyle.color', '#000000')
  .setOption('legend.textStyle.color', '#1a1a1a')
  .setOption('titleTextStyle.color', '#202124')
  .setOption('titleTextStyle.alignment', 'center')
  .setOption('titleTextStyle.bold', true)
  .setOption('annotations.total.textStyle.color', '#808080')
  .setXAxisTitle('Months')
  .setOption('hAxis.textStyle.color', '#000000')
  .setOption('hAxis.textStyle.bold', true)
  .setOption('hAxis.titleTextStyle.color', '#000000')
  .setOption('hAxis.titleTextStyle.bold', true)
  .setOption('vAxes.0.minorGridlines.count', 2)
  .setYAxisTitle('DSO (Days)')
  .setOption('vAxes.0.textStyle.color', '#000000')
  .setOption('vAxes.0.textStyle.bold', true)
  .setOption('vAxes.0.titleTextStyle.color', '#000000')
  .setOption('vAxes.0.titleTextStyle.bold', true)
  .setOption('backgroundColor', '#CEEAD6')
  .setPosition(33, 1, 13, 18)
  .build();
  sheet.insertChart(chart);
};

function generateDIOColumnChart() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = spreadsheet.getSheetByName(CASH_FLOW_HEADERS.SHEET_NAME);

  var chart = sheet.newChart()
  .asColumnChart()
  .addRange(spreadsheet.getRange('A2:A14'))
  .addRange(spreadsheet.getRange('F2:F14'))
  .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
  .setTransposeRowsAndColumns(false)
  .setNumHeaders(1)
  .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
  .setOption('bubble.stroke', '#000000')
  .setOption('useFirstColumnAsDomain', true)
  .setOption('isStacked', 'false')
  .setOption('title', 'Days Inventory Outstanding (DIO) over Time')
  .setOption('annotations.domain.textStyle.color', '#808080')
  .setOption('textStyle.color', '#000000')
  .setOption('legend.textStyle.color', '#1a1a1a')
  .setOption('titleTextStyle.color', '#202124')
  .setOption('titleTextStyle.alignment', 'center')
  .setOption('titleTextStyle.bold', true)
  .setOption('annotations.total.textStyle.color', '#808080')
  .setXAxisTitle('Months')
  .setOption('hAxis.textStyle.color', '#000000')
  .setOption('hAxis.textStyle.bold', true)
  .setOption('hAxis.titleTextStyle.color', '#000000')
  .setOption('hAxis.titleTextStyle.bold', true)
  .setOption('vAxes.0.minorGridlines.count', 2)
  .setYAxisTitle('DIO (Days)')
  .setOption('vAxes.0.textStyle.color', '#000000')
  .setOption('vAxes.0.textStyle.bold', true)
  .setOption('vAxes.0.titleTextStyle.color', '#000000')
  .setOption('vAxes.0.titleTextStyle.bold', true)
  .setOption('backgroundColor', '#E2DDFE')
  .setPosition(33, 5, 24, 18)
  .build();
  sheet.insertChart(chart);
};
