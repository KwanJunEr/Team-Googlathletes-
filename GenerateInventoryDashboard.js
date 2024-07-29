function generateInventoryDashboard() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();
  
  // Remove old charts
  var charts = sheet.getCharts();
  for (var i in charts) {
    var chart = charts[i];
    sheet.removeChart(chart);
  }
  
  // Generate new charts
  generateQuantityInStockByProductBarChart();
  generateQuantityInStockVsTotalSoldBubbleChart();
  generateQuantityInStockAndTotalSoldStackedBarChart();
  generateSellingPriceVsUnitCostScatterChart();
  generateTotalSoldtoProductLineChart();
  generateSellingPriceDistributionHistogram();
}

function generateQuantityInStockByProductBarChart() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();
  
  var chart = sheet.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(sheet.getRange('B2:B'))
    .addRange(sheet.getRange('D2:D'))
    .setOption('title', 'Quantity in Stock by Product')
    .setOption('hAxis.title', 'Quantity in Stock')
    .setOption('vAxis.title', 'Product Name')
    .setOption('legend.position', 'none')
    .setOption('hAxis.textStyle.color', '#000000')
    .setOption('hAxis.titleTextStyle.bold', true)
    .setOption('hAxis.titleTextStyle.fontSize', 14)
    .setOption('vAxis.textStyle.color', '#000000')
    .setOption('vAxis.titleTextStyle.bold', true)
    .setOption('vAxis.titleTextStyle.fontSize', 14)
    .setOption('titleTextStyle.color', '#202124')
    .setOption('titleTextStyle.alignment', 'center')
    .setOption('titleTextStyle.bold', true)
    .setOption('textStyle.color', '#000000')
    .setOption('series.color', '#4285F4')
    .setOption('legend.textStyle.fontSize', 12)
    .setOption('legend.textStyle.color', '#202124')
    .setOption('legend.textStyle.bold', true)
    .setOption('series', { 0: {color: '#FFA50C'} })
    .setPosition(2, 13, 10, 5)
    .build();
    
  sheet.insertChart(chart);
}

function generateQuantityInStockVsTotalSoldBubbleChart() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();
  
  var chart = sheet.newChart()
    .setChartType(Charts.ChartType.BUBBLE)
    .addRange(sheet.getRange('B2:B'))
    .addRange(sheet.getRange('G2:G'))
    .addRange(sheet.getRange('D2:D')) 
    .setOption('title', 'Quantity in Stock vs Total Sold')
    .setOption('hAxis.title', 'Quantity in Stock')
    .setOption('hAxis.titleTextStyle.bold', true)
    .setOption('hAxis.titleTextStyle.fontSize', 14)
    .setOption('vAxis.title', 'Total Sold')
    .setOption('vAxis.titleTextStyle.bold', true)
    .setOption('vAxis.titleTextStyle.fontSize', 14)
    .setOption('textStyle.color', '#000000')
    .setOption('titleTextStyle.color', '#202124')
    .setOption('titleTextStyle.alignment', 'center')
    .setOption('titleTextStyle.bold', true)
    .setOption('hAxis.textStyle.fontSize', 12)
    .setOption('vAxis.textStyle.fontSize', 12)
    .setOption('hAxis.textStyle.color', '#000000')
    .setOption('vAxis.textStyle.color', '#000000')
    .setOption('legend.position', 'none')
    .setOption('series', { 0: {color: '#6700F1'} })
    .setPosition(20, 13, 10, 7)
    .build();
    
  sheet.insertChart(chart);
}


function generateQuantityInStockAndTotalSoldStackedBarChart() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();
  
  var chart = sheet.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(sheet.getRange('B2:B'))
    .addRange(sheet.getRange('D2:D'))
    .addRange(sheet.getRange('G2:G'))
    .setOption('title', 'Quantity in Stock and Total Sold by Product')
    .setOption('hAxis.title', 'Quantity')
    .setOption('hAxis.titleTextStyle.bold', true)
    .setOption('hAxis.titleTextStyle.fontSize', 14)
    .setOption('vAxis.title', 'Product Name')
    .setOption('vAxis.titleTextStyle.bold', true)
    .setOption('vAxis.titleTextStyle.fontSize', 14)
    .setOption('legend.position', 'bottom')
    .setOption('legend.textStyle.fontSize', 12)
    .setOption('legend.textStyle.color', '#202124')
    .setOption('legend.textStyle.bold', true)
    .setOption('textStyle.color', '#000000')
    .setOption('titleTextStyle.color', '#202124')
    .setOption('titleTextStyle.bold', true)
    .setOption('titleTextStyle.alignment', 'center')
    .setOption('hAxis.textStyle.fontSize', 12)
    .setOption('vAxis.textStyle.fontSize', 12)
    .setOption('hAxis.textStyle.color', '#000000')
    .setOption('vAxis.textStyle.color', '#000000')
    .setOption('isStacked', true)
    .setOption('series', {
      0: {color: '#EA4335', labelInLegend: 'Quantity in Stock'},
      1: {color: '#5BD534', labelInLegend: 'Total Sold'}
    })
    .setOption('annotations', {
      alwaysOutside: true,
      textStyle: {
        fontSize: 12,
        bold: true,
        color: '#000000'
      }
    })
    .setPosition(38, 13, 10, 9)
    .build();
    
  sheet.insertChart(chart);
}

function generateSellingPriceVsUnitCostScatterChart() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();
  
  var chart = sheet.newChart()
    .setChartType(Charts.ChartType.SCATTER)
    .addRange(sheet.getRange('F2:F'))
    .addRange(sheet.getRange('E2:E'))
    .setOption('title', 'Selling Price vs Unit Cost')
    .setOption('hAxis.title', 'Unit Cost (RM)')
    .setOption('hAxis.titleTextStyle.bold', true)
    .setOption('hAxis.titleTextStyle.fontSize', 14)
    .setOption('vAxis.title', 'Selling Price (RM)')
    .setOption('vAxis.titleTextStyle.bold', true)
    .setOption('vAxis.titleTextStyle.fontSize', 14)
    .setOption('textStyle.color', '#000000')
    .setOption('titleTextStyle.color', '#202124')
    .setOption('titleTextStyle.bold', true)
    .setOption('titleTextStyle.alignment', 'center')
    .setOption('hAxis.textStyle.fontSize', 12)
    .setOption('vAxis.textStyle.fontSize', 12)
    .setOption('hAxis.textStyle.color', '#000000')
    .setOption('vAxis.textStyle.color', '#000000')
    .setOption('legend.position', 'none')
    .setOption('series', { 0: { color: '#0C3CFF' } })
    .setPosition(38, 19, 19, 9)
    .build();
    
  sheet.insertChart(chart);
}

function generateTotalSoldtoProductLineChart() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();
  
  var chart = sheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(sheet.getRange('B2:B'))
    .addRange(sheet.getRange('G2:G'))
    .setOption('title', 'Total Sold According to Product')
    .setOption('hAxis.title', 'Product Name')
    .setOption('hAxis.titleTextStyle.bold', true)
    .setOption('hAxis.titleTextStyle.fontSize', 14)
    .setOption('vAxis.title', 'Total Sold')
    .setOption('vAxis.titleTextStyle.bold', true)
    .setOption('vAxis.titleTextStyle.fontSize', 14)
    .setOption('legend.position', 'none')
    .setOption('textStyle.color', '#000000')
    .setOption('titleTextStyle.color', '#202124')
    .setOption('titleTextStyle.bold', true)
    .setOption('titleTextStyle.alignment', 'center')
    .setOption('hAxis.textStyle.fontSize', 12)
    .setOption('vAxis.textStyle.fontSize', 12)
    .setOption('hAxis.textStyle.color', '#000000')
    .setOption('vAxis.textStyle.color', '#000000')
    .setOption('series', { 0: { color: '#FD0E9E' } })
    .setPosition(20, 19, 19, 7)
    .build();
    
  sheet.insertChart(chart);
}

function generateSellingPriceDistributionHistogram() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();
  
  var chart = sheet.newChart()
    .setChartType(Charts.ChartType.HISTOGRAM)
    .addRange(sheet.getRange('E2:E'))
    .setOption('title', 'Distribution of Selling Prices')
    .setOption('hAxis.title', 'Selling Price (RM)')
    .setOption('hAxis.titleTextStyle.bold', true)
    .setOption('hAxis.titleTextStyle.fontSize', 14)
    .setOption('vAxis.title', 'Number of Products')
    .setOption('vAxis.titleTextStyle.bold', true)
    .setOption('vAxis.titleTextStyle.fontSize', 14)
    .setOption('textStyle.color', '#000000')
    .setOption('titleTextStyle.color', '#202124')
    .setOption('titleTextStyle.bold', true)
    .setOption('titleTextStyle.alignment', 'center')
    .setOption('hAxis.textStyle.fontSize', 12)
    .setOption('vAxis.textStyle.fontSize', 12)
    .setOption('hAxis.textStyle.color', '#000000')
    .setOption('vAxis.textStyle.color', '#000000')
    .setOption('series', { 0: { color: '#36F7FF' } })
    .setPosition(2, 19, 19, 5)
    .build();
    
  sheet.insertChart(chart);
}

