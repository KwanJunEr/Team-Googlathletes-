function generateCustomerDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CUSTOMER_HEADERS.SHEET_NAME);
  if (sheet) {
    removeAllCharts(sheet);
  } else {
    createCustomerTable();
  }
  ss.setActiveSheet(sheet);

  generateCustomerSegmentsPieChart();
  generateSaleCountAndDaysScatterChart();
  generateSalesCountAndTotalRevenueScatterChart();
  generateSalesCountBySegmentBarChart();
  generateTotalSalesBySegmentBarChart();
  generateTotalSalesAndDaysScatterChart();
}

function generateCustomerSegmentsPieChart() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(CUSTOMER_HEADERS.SHEET_NAME);

  var chart = sheet.newChart()
  .asPieChart()
  .addRange(spreadsheet.getRange('L1:L'))
  .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
  .setTransposeRowsAndColumns(false)
  .setNumHeaders(1)
  .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
  .setOption('applyAggregateData', 0)
  .setOption('bubble.stroke', '#000000')
  .setOption('pieSliceText', 'value')
  .setOption('title', 'Percentage of Customers in Each Segment')
  .setOption('annotations.domain.textStyle.color', '#202124')
  .setOption('textStyle.color', '#000000')
  .setOption('legend.textStyle.fontSize', 14)
  .setOption('legend.textStyle.color', '#202124')
  .setOption('legend.textStyle.bold', true)
  .setOption('pieSliceTextStyle.color', '#000000')
  .setOption('pieSliceTextStyle.bold', true)
  .setOption('titleTextStyle.color', '#202124')
  .setOption('titleTextStyle.alignment', 'center')
  .setOption('titleTextStyle.bold', true)
  .setPosition(2, 13, 10, 5)
  .build();
  sheet.insertChart(chart);
};

function generateTotalSalesBySegmentBarChart() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(CUSTOMER_HEADERS.SHEET_NAME);

  var chart = sheet.newChart()
    .asBarChart()
    .addRange(spreadsheet.getRange('L1:L'))
    .addRange(spreadsheet.getRange('G1:G'))
    .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
    .setTransposeRowsAndColumns(false)
    .setNumHeaders(1)
    .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
    .setOption('bubble.stroke', '#000000')
    .setOption('title', 'Total Sales by Segment')
    .setOption('annotations.domain.textStyle.color', '#808080')
    .setOption('textStyle.color', '#000000')
    .setOption('legend.textStyle.color', '#1a1a1a')
    .setOption('titleTextStyle.color', '#202124')
    .setOption('titleTextStyle.alignment', 'center')
    .setOption('titleTextStyle.bold', true)
    .setOption('annotations.total.textStyle.color', '#808080')
    .setOption('hAxis.minorGridlines.count', 2)
    .setXAxisTitle('Total Sales (RM)')
    .setOption('hAxis.textStyle.color', '#000000')
    .setOption('hAxis.titleTextStyle.color', '#000000')
    .setOption('hAxis.titleTextStyle.bold', true)
    .setYAxisTitle('Segments')
    .setOption('vAxes.0.textStyle.color', '#000000')
    .setOption('vAxes.0.titleTextStyle.color', '#000000')
    .setOption('vAxes.0.titleTextStyle.bold', true)
    .setOption('series.0.hasAnnotations', true)
    .setOption('series.0.dataLabel', 'value')
    .setOption('series.0.textStyle.bold', true)
    .setOption('series.0.color', '#4285F4')
    .setPosition(20, 13, 10, 9)
    .build();
  sheet.insertChart(chart);
};

function generateSalesCountBySegmentBarChart() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(CUSTOMER_HEADERS.SHEET_NAME);

  var chart = sheet.newChart()
    .asBarChart()
    .addRange(spreadsheet.getRange('L1:L'))
    .addRange(spreadsheet.getRange('F1:F'))
    .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
    .setTransposeRowsAndColumns(false)
    .setNumHeaders(1)
    .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
    .setOption('bubble.stroke', '#000000')
    .setOption('title', 'Sale Count by Segment')
    .setOption('annotations.domain.textStyle.color', '#808080')
    .setOption('textStyle.color', '#000000')
    .setOption('legend.textStyle.color', '#1a1a1a')
    .setOption('titleTextStyle.color', '#202124')
    .setOption('titleTextStyle.alignment', 'center')
    .setOption('titleTextStyle.bold', true)
    .setOption('annotations.total.textStyle.color', '#808080')
    .setOption('hAxis.minorGridlines.count', 2)
    .setXAxisTitle('Sale Count')
    .setOption('hAxis.textStyle.color', '#000000')
    .setOption('hAxis.titleTextStyle.color', '#000000')
    .setOption('hAxis.titleTextStyle.bold', true)
    .setYAxisTitle('Segments')
    .setOption('vAxes.0.textStyle.color', '#000000')
    .setOption('vAxes.0.titleTextStyle.color', '#000000')
    .setOption('vAxes.0.titleTextStyle.bold', true)
    .setOption('series.0.hasAnnotations', true)
    .setOption('series.0.dataLabel', 'value')
    .setOption('series.0.textStyle.bold', true)
    .setOption('series.0.color', '#EA4335') 
    .setPosition(2, 19, 19, 5)
    .build();
  sheet.insertChart(chart);
};

function generateSalesCountAndTotalRevenueScatterChart() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(CUSTOMER_HEADERS.SHEET_NAME);

  var chart = sheet.newChart()
  .asScatterChart()
  .addRange(spreadsheet.getRange('F1:G'))
  .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
  .setTransposeRowsAndColumns(false)
  .setNumHeaders(1)
  .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
  .setOption('bubble.stroke', '#000000')
  .setOption('title', 'Sale Count vs Sales')
  .setOption('annotations.domain.textStyle.color', '#808080')
  .setOption('textStyle.color', '#000000')
  .setOption('legend.textStyle.color', '#1a1a1a')
  .setOption('titleTextStyle.color', '#202124')
  .setOption('titleTextStyle.alignment', 'center')
  .setOption('titleTextStyle.bold', true)
  .setOption('annotations.total.textStyle.color', '#808080')
  .setOption('hAxis.minorGridlines.count', 2)
  .setXAxisTitle('Sale Count')
  .setOption('hAxis.textStyle.color', '#000000')
  .setOption('hAxis.titleTextStyle.color', '#000000')
  .setOption('hAxis.titleTextStyle.bold', true)
  .setOption('vAxes.0.minorGridlines.count', 2)
  .setYAxisTitle('Sales (RM)')
  .setOption('vAxes.0.textStyle.color', '#000000')
  .setOption('vAxes.0.titleTextStyle.color', '#000000')
  .setOption('vAxes.0.titleTextStyle.bold', true)
  .setOption('series.0.color', '#34A853')
  .setPosition(20, 19, 19, 9)
  .build();
  sheet.insertChart(chart);
};

function generateTotalSalesAndDaysScatterChart() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(CUSTOMER_HEADERS.SHEET_NAME);

  var chart = sheet.newChart()
  .asScatterChart()
  .addRange(spreadsheet.getRange('G1:H'))
  .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
  .setTransposeRowsAndColumns(false)
  .setNumHeaders(1)
  .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
  .setOption('bubble.stroke', '#000000')
  .setOption('title', 'Sales vs Days Since Last Purchase')
  .setOption('annotations.domain.textStyle.color', '#808080')
  .setOption('textStyle.color', '#000000')
  .setOption('legend.textStyle.color', '#1a1a1a')
  .setOption('titleTextStyle.color', '#202124')
  .setOption('titleTextStyle.alignment', 'center')
  .setOption('titleTextStyle.bold', true)
  .setOption('annotations.total.textStyle.color', '#808080')
  .setOption('hAxis.minorGridlines.count', 2)
  .setXAxisTitle('Sales (RM)')
  .setOption('hAxis.textStyle.color', '#000000')
  .setOption('hAxis.titleTextStyle.color', '#000000')
  .setOption('hAxis.titleTextStyle.bold', true)
  .setOption('vAxes.0.minorGridlines.count', 2)
  .setYAxisTitle('Days Since Last Purchase (Days)')
  .setOption('vAxes.0.textStyle.color', '#000000')
  .setOption('vAxes.0.titleTextStyle.color', '#000000')
  .setOption('vAxes.0.titleTextStyle.bold', true)
  .setOption('series.0.color', '#FBBC04')
  .setPosition(38, 13, 10, 13)
  .build();
  sheet.insertChart(chart);
};

function generateSaleCountAndDaysScatterChart() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(CUSTOMER_HEADERS.SHEET_NAME);

  var chart = sheet.newChart()
  .asScatterChart()
  .addRange(spreadsheet.getRange('F1:F'))
  .addRange(spreadsheet.getRange('H1:H'))
  .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
  .setTransposeRowsAndColumns(false)
  .setNumHeaders(1)
  .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
  .setOption('bubble.stroke', '#000000')
  .setOption('title', 'Sale Count vs Days Since Last Purchase')
  .setOption('annotations.domain.textStyle.color', '#808080')
  .setOption('textStyle.color', '#000000')
  .setOption('legend.textStyle.color', '#1a1a1a')
  .setOption('titleTextStyle.color', '#202124')
  .setOption('titleTextStyle.alignment', 'center')
  .setOption('titleTextStyle.bold', true)
  .setOption('annotations.total.textStyle.color', '#808080')
  .setOption('hAxis.minorGridlines.count', 2)
  .setXAxisTitle('Sale Count')
  .setOption('hAxis.textStyle.color', '#000000')
  .setOption('hAxis.titleTextStyle.color', '#000000')
  .setOption('hAxis.titleTextStyle.bold', true)
  .setOption('vAxes.0.minorGridlines.count', 2)
  .setYAxisTitle('Days Since Last Purchase (Days)')
  .setOption('vAxes.0.textStyle.color', '#000000')
  .setOption('vAxes.0.titleTextStyle.color', '#000000')
  .setOption('vAxes.0.titleTextStyle.bold', true)
  .setOption('series.0.color', '#762ac7')
  .setPosition(38, 19, 19, 13)
  .build();
  sheet.insertChart(chart);
};
