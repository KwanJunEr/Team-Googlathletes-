function logExpensesData(form) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Expenses');

  var lastRow = sheet.getLastRow();
  var idValue = 'EX-' + Utilities.formatString('%03d', lastRow);

  var newRow = [
    idValue,
    form.name,
    form.date,
    form.category,
    form.description,
    form.operatingExpense,
    form.capitalExpenditure,
    form.total
  ];
  
  sheet.appendRow(newRow);

  return "Record added successfully!";
}

