function logIncomeData(form) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Income');

  var lastRow = sheet.getLastRow();
  var idValue = 'IN-' + Utilities.formatString('%03d', lastRow);

  var newRow = [
    idValue,
    form.name,
    form.date,
    form.category,
    form.description,
    form.total
  ];
  sheet.appendRow(newRow);

  return "Record added successfully!";
}

