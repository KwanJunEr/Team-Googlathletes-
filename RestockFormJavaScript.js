function logRestockData(form) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Restock');

  var lastRow = sheet.getLastRow();
  var idValue = 'STOCK-' + Utilities.formatString('%03d', lastRow);

  var newRow = [
    idValue,
    form.priority,
    form.branch,
    form.supplier,
    form.product,
    form.quantity,
    form.status,
    form.date,
    form.arrivalDate,
    form.paymentDate,
    form.total,
    form.paymentMethod,
    form.notes
  ];
  sheet.appendRow(newRow);

  return "Record added successfully!";
}

