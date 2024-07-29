function logCustomerData(form) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Customer');

  var lastRow = sheet.getLastRow();

  var contactNumberCell = 'C'+(lastRow+1);
  var recentOrderCell = 'E'+(lastRow+1);
  var saleCountCell = 'F'+(lastRow+1);
  var totalSalesCell = 'G'+(lastRow+1);

  var rPercentile = 'PERCENTRANK($E$2:E,'+recentOrderCell+', 2)';
  var rScore = '=IFERROR(IF('+rPercentile+'>= 0.8, 5, IF('+rPercentile+'>= 0.6, 4, IF('+rPercentile+'>= 0.4, 3, IF('+rPercentile+'>= 0.2, 2, 1)))), "")';

  var fPercentile = 'PERCENTRANK($F$2:F,'+saleCountCell+', 2)';
  var fScore = '=IFERROR(IF('+fPercentile+' >= 0.8, 5, IF('+fPercentile+' >= 0.6, 4, IF('+fPercentile+' >= 0.4, 3, IF('+fPercentile+' >= 0.2, 2, 1)))), "")';

  var mPercentile = 'PERCENTRANK($G$2:G, '+totalSalesCell+', 2)';
  var mScore = '=IFERROR(IF('+mPercentile+' >= 0.8, 5,IF('+mPercentile+' >= 0.6, 4,IF('+mPercentile+' >= 0.4, 3,IF('+mPercentile+' >= 0.2, 2, 1)))), "")';

  var rCell = 'I'+(lastRow+1);
  var fCell = 'J'+(lastRow+1);
  var gCell = 'K'+(lastRow+1);

  var segment = '=IF(AND('+rCell+'=5,'+fCell+'=5,'+gCell+'=5),"High Value",IF(AND('+rCell+'=1),"Lost",IF(AND('+rCell+'>2,'+fCell+'>3),"Core",IF(AND('+rCell+'>2,'+fCell+'=3),"Promising",IF(AND('+rCell+'=5,'+fCell+'<3),"New",IF(AND('+rCell+'=2,'+fCell+'>3),"Need Attention", "General"))))))';

  var newRow = [
    '="CUS-" & TEXT(ROW() - 1, "000")',
    form.name,
    form.contactNumber,
    form.email,
    '=MAXIFS(Sales!E:E, Sales!D:D,'+contactNumberCell+')',
    '=COUNTIFS(Sales!D:D,'+contactNumberCell+')',
    '=SUMIFS(Sales!Q:Q, Sales!D:D,'+contactNumberCell+')',
    '=TODAY()-'+recentOrderCell,
    rScore,
    fScore,
    mScore,
    segment
  ];
  sheet.appendRow(newRow);

  return "Record added successfully!";
}
