const SALES_DASHBOARD = 'Sales Dashboard';

function generateSalesDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SALES_HEADERS.SHEET_NAME);
  var sheetCount = ss.getSheets().length;
  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();

  let salesDashboard = ss.getSheetByName(SALES_DASHBOARD);
  if (salesDashboard) {
    salesDashboard.clear();
    removeAllCharts(salesDashboard);
  } else {
    salesDashboard = ss.insertSheet(SALES_DASHBOARD);
  }
  ss.setActiveSheet(salesDashboard);

  generateSalesbyBranchSheet(data, ss);
  generateSalesbyChannelSheet(data, ss);
  generateSalesbyStatusSheet(data,ss);
  generateSalesbyProductSheet(data,ss);
  generateMonthlySalesSheet(data,ss);
}

function removeAllCharts(sheet) {
  const charts = sheet.getCharts();
  for (const chart of charts) {
    sheet.removeChart(chart);
  }
}

function generateMonthlySalesSheet(data, ss) {
  const monthlySalesData = [['Month', 'Total Sales (RM)']];
  const monthSales = {};
  
  const monthNames = {
    1: 'January', 2: 'February', 3: 'March', 4: 'April', 5: 'May', 6: 'June',
    7: 'July', 8: 'August', 9: 'September', 10: 'October', 11: 'November', 12: 'December'
  };

  for (let i = 1; i < data.length; i++) { // Skip header row
    const dateStr = data[i][4]; // Date column
    const total = parseFloat(data[i][16]) || 0; // Total Sales (RM) column
    const status = data[i][7]; // Status column

    if (status === 'Completed') {
      let month;
      try {
        // Attempt to parse date in 'YYYY-MM-DD' format
        const date = new Date(dateStr);
        month = date.getMonth() + 1; // JavaScript months are 0-based
      } catch (e) {
        // Attempt to parse date in 'DD-MM-YYYY' format
        const parts = dateStr.split('-');
        if (parts.length === 3) {
          // Handle 'DD-MM-YYYY' format
          const day = parseInt(parts[2], 10);
          const monthStr = parseInt(parts[1], 10);
          const year = parseInt(parts[0], 10);
          const date = new Date(year, monthStr - 1, day);
          month = date.getMonth() + 1;
        } else {
          continue; // Skip invalid date formats
        }
      }

      if (month) {
        if (!monthSales[month]) {
          monthSales[month] = 0;
        }
        monthSales[month] += total;
      }
    }
  }

  // Convert monthSales to the format required
  for (const [key, value] of Object.entries(monthSales)) {
    monthlySalesData.push([monthNames[parseInt(key)], value]);
  }

  const newSheet = ss.getSheetByName('MonthlySalesData') || ss.insertSheet('MonthlySalesData');
  newSheet.clear();
  newSheet.getRange(1, 1, monthlySalesData.length, monthlySalesData[0].length).setValues(monthlySalesData);

  newSheet.hideSheet();
  generateMonthlySalesGraph(ss);
}

function generateSalesbyProductSheet(data,ss){
  const productSalesData = [['Product', 'Total Sales (RM)']];
  const products = [];
  const productSums = {};

  for (let i = 1; i < data.length; i++) { // Start from 1 to skip header row
    const total = parseFloat(data[i][16]) || 0;
    const product = data[i][9];
    const status = data[i][7];

    if (product && status === 'Completed') {
      if (!products.includes(product)){
        products.push(product);
        productSums[product] = 0;
      }
      productSums[product] += total;
    }
  }

  for (const product of products) {
    productSalesData.push([product, productSums[product]]);
  }

  const newSheet = ss.getSheetByName('ProductSalesData') || ss.insertSheet('ProductSalesData');
  newSheet.clear();
  newSheet.getRange(1, 1, productSalesData.length, productSalesData[0].length).setValues(productSalesData);

  newSheet.hideSheet();
  generateProductSalesGraph(ss);
}

function generateSalesbyStatusSheet(data,ss){
  const statusSalesData = [['Status', 'Total Sales (RM)']];
  const statuses = [];
  const statusSums = {};

  for (let i = 1; i < data.length; i++) { // Start from 1 to skip header row
    const total = parseFloat(data[i][16]) || 0;
    const status = data[i][7];

    if (status) {
      if (!statuses.includes(status)){
        statuses.push(status);
        statusSums[status] = 0;
      }
      statusSums[status] += total;
    }
  }

  for (const status of statuses) {
    statusSalesData.push([status, statusSums[status]]);
  }

  const newSheet = ss.getSheetByName('StatusSalesData') || ss.insertSheet('StatusSalesData');
  newSheet.clear();
  newSheet.getRange(1, 1, statusSalesData.length, statusSalesData[0].length).setValues(statusSalesData);

  newSheet.hideSheet();
  generateStatusGraph(ss);

}

function generateSalesbyChannelSheet(data, ss) {
  const channelSalesData = [['Sales Channel', 'Total Sales (RM)']];
  const channels = [];
  const channelSums = {};

  for (let i = 1; i < data.length; i++) { // Start from 1 to skip header row
    const salesChannel = data[i][8];
    const total = parseFloat(data[i][16]) || 0;
    const status = data[i][7];

    if (salesChannel && status === 'Completed') {
      if (!channels.includes(salesChannel)) {
        channels.push(salesChannel);
        channelSums[salesChannel] = 0;
      }
      channelSums[salesChannel] += total;
    }
  }

  for (const salesChannel of channels) {
    channelSalesData.push([salesChannel, channelSums[salesChannel]]);
  }

  const newSheet = ss.getSheetByName('ChannelSalesData') || ss.insertSheet('ChannelSalesData');
  newSheet.clear();
  newSheet.getRange(1, 1, channelSalesData.length, channelSalesData[0].length).setValues(channelSalesData);

  newSheet.hideSheet();
  generateSalesChannelGraph(ss);
}

function generateSalesbyBranchSheet(data, ss) {
  const branchSalesData = [['Branch', 'Total Sales (RM)']];
  const branches = [];
  const branchSums = {};

  for (let i = 1; i < data.length; i++) { // Start from 1 to skip header row
    const branch = data[i][2];
    const total = parseFloat(data[i][16]) || 0;
    const status = data[i][7];

    if (branch && status === 'Completed') {
      if (!branches.includes(branch)) {
        branches.push(branch);
        branchSums[branch] = 0;
      }
      branchSums[branch] += total;
    }
  }

  for (const branch of branches) {
    branchSalesData.push([branch, branchSums[branch]]);
  }

  const newSheet = ss.getSheetByName('BranchSalesData') || ss.insertSheet('BranchSalesData');
  newSheet.clear();
  newSheet.getRange(1, 1, branchSalesData.length, branchSalesData[0].length).setValues(branchSalesData);

  newSheet.hideSheet();

  generateSalesbyBranchGraph(ss);
}

function generateMonthlySalesGraph(ss) {
  const dataSheet = ss.getSheetByName('MonthlySalesData');
  const dataRange = dataSheet.getDataRange();

  const graphSheet = ss.getSheetByName(SALES_DASHBOARD);
  const monthlySalesChart = graphSheet.newChart()
    .asLineChart()
    .addRange(dataRange)
    .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
    .setTransposeRowsAndColumns(false)
    .setNumHeaders(1)
    .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
    .setOption('bubble.stroke', '#000000')
    .setOption('useFirstColumnAsDomain', true)
    .setOption('isStacked', 'false')
    .setOption('title', 'Total Sales over Time')
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
    .setYAxisTitle('Total Sales (RM)')
    .setOption('vAxes.0.textStyle.color', '#000000')
    .setOption('vAxes.0.textStyle.bold', true)
    .setOption('vAxes.0.titleTextStyle.color', '#000000')
    .setOption('vAxes.0.titleTextStyle.bold', true)
    .setOption('series.0.pointSize', 7)
    .setOption('backgroundColor', '#D2E3FC')
    .setOption('width', 1504)
    .setPosition(45, 1, 0, 0)
    .build();

  graphSheet.insertChart(monthlySalesChart);
}

function generateProductSalesGraph(ss){
  const dataSheet = ss.getSheetByName('ProductSalesData');
  const dataRange = dataSheet.getDataRange();

  const graphSheet = ss.getSheetByName(SALES_DASHBOARD);
  const productSalesChart = graphSheet.newChart()
    .asColumnChart()
    .addRange(dataRange)
    .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
    .setTransposeRowsAndColumns(false)
    .setNumHeaders(1)
    .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
    .setOption('bubble.stroke', '#000000')
    .setOption('useFirstColumnAsDomain', true)
    .setOption('isStacked', 'false')
    .setOption('title', 'Total Sales by Product')
    .setOption('annotations.domain.textStyle.color', '#808080')
    .setOption('textStyle.color', '#000000')
    .setOption('legend.textStyle.color', '#1a1a1a')
    .setOption('titleTextStyle.color', '#202124')
    .setOption('titleTextStyle.alignment', 'center')
    .setOption('titleTextStyle.bold', true)
    .setOption('annotations.total.textStyle.color', '#808080')
    .setXAxisTitle('Product')
    .setOption('hAxis.textStyle.color', '#000000')
    .setOption('hAxis.textStyle.bold', true)
    .setOption('hAxis.titleTextStyle.color', '#000000')
    .setOption('hAxis.titleTextStyle.bold', true)
    .setOption('vAxes.0.minorGridlines.count', 2)
    .setYAxisTitle('Total Sales (RM)')
    .setOption('vAxes.0.textStyle.color', '#000000')
    .setOption('vAxes.0.textStyle.bold', true)
    .setOption('vAxes.0.titleTextStyle.color', '#000000')
    .setOption('vAxes.0.titleTextStyle.bold', true)
    .setOption('backgroundColor', '#FEEFC3')
    .setPosition(25, 10, 0, 0)
    .build();

  graphSheet.insertChart(productSalesChart);
}

function generateStatusGraph(ss){
  const dataSheet = ss.getSheetByName('StatusSalesData');
  const dataRange = dataSheet.getDataRange();

  const graphSheet = ss.getSheetByName(SALES_DASHBOARD);
  const statusSalesChart = graphSheet.newChart()
    .asBarChart()
    .addRange(dataRange)
    .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
    .setTransposeRowsAndColumns(false)
    .setNumHeaders(1)
    .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
    .setOption('bubble.stroke', '#000000')
    .setOption('useFirstColumnAsDomain', true)
    .setOption('isStacked', 'false')
    .setOption('title', 'Total Sales by Status')
    .setOption('annotations.domain.textStyle.color', '#808080')
    .setOption('textStyle.color', '#000000')
    .setOption('legend.textStyle.color', '#1a1a1a')
    .setOption('titleTextStyle.color', '#202124')
    .setOption('titleTextStyle.alignment', 'center')
    .setOption('titleTextStyle.bold', true)
    .setOption('annotations.total.textStyle.color', '#808080')
    .setXAxisTitle('Total Sales (RM)')
    .setOption('hAxis.textStyle.color', '#000000')
    .setOption('hAxis.textStyle.bold', true)
    .setOption('hAxis.titleTextStyle.color', '#000000')
    .setOption('hAxis.titleTextStyle.bold', true)
    .setOption('vAxes.0.minorGridlines.count', 2)
    .setYAxisTitle('Status')
    .setOption('vAxes.0.textStyle.color', '#000000')
    .setOption('vAxes.0.textStyle.bold', true)
    .setOption('vAxes.0.titleTextStyle.color', '#000000')
    .setOption('vAxes.0.titleTextStyle.bold', true)
    .setOption('backgroundColor', '#CEEAD6')
    .setPosition(25, 1, 0, 0)
    .build();

  graphSheet.insertChart(statusSalesChart);
}

function generateSalesChannelGraph(ss) {
  const dataSheet = ss.getSheetByName('ChannelSalesData');
  const dataRange = dataSheet.getDataRange();

  const graphSheet = ss.getSheetByName(SALES_DASHBOARD);
  const salesChannelChart = graphSheet.newChart()
    .asPieChart()
    .addRange(dataRange)
    .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
    .setTransposeRowsAndColumns(false)
    .setNumHeaders(1)
    .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
    .setOption('applyAggregateData', 0)
    .setOption('bubble.stroke', '#000000')
    .setOption('pieSliceText', 'value')
    .setOption('title', 'Total Sales by Sales Channel')
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
    .setPosition(5, 10, 0, 0) // Adjusted position to place beside the other chart
    .build();

  graphSheet.insertChart(salesChannelChart);
}

function generateSalesbyBranchGraph(ss) {
  const dataSheet = ss.getSheetByName('BranchSalesData');
  const dataRange = dataSheet.getDataRange();

  const graphSheet = ss.getSheetByName(SALES_DASHBOARD);
  const branchSalesChart = graphSheet.newChart()
    .asBarChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(dataRange)
    .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
    .setTransposeRowsAndColumns(false)
    .setNumHeaders(1)
    .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
    .setOption('bubble.stroke', '#000000')
    .setOption('useFirstColumnAsDomain', true)
    .setOption('isStacked', 'false')
    .setOption('title', 'Total Sales by Branch')
    .setOption('annotations.domain.textStyle.color', '#808080')
    .setOption('textStyle.color', '#000000')
    .setOption('legend.textStyle.color', '#1a1a1a')
    .setOption('titleTextStyle.color', '#202124')
    .setOption('titleTextStyle.alignment', 'center')
    .setOption('titleTextStyle.bold', true)
    .setOption('annotations.total.textStyle.color', '#808080')
    .setXAxisTitle('Total Sales (RM)')
    .setOption('hAxis.textStyle.color', '#000000')
    .setOption('hAxis.textStyle.bold', true)
    .setOption('hAxis.titleTextStyle.color', '#000000')
    .setOption('hAxis.titleTextStyle.bold', true)
    .setOption('vAxes.0.minorGridlines.count', 2)
    .setYAxisTitle('Branch')
    .setOption('vAxes.0.textStyle.color', '#000000')
    .setOption('vAxes.0.textStyle.bold', true)
    .setOption('vAxes.0.titleTextStyle.color', '#000000')
    .setOption('vAxes.0.titleTextStyle.bold', true)
    .setOption('backgroundColor', '#E2DDFE')
    .setPosition(5, 1, 0, 0)
    .build();

  graphSheet.insertChart(branchSalesChart);
}

