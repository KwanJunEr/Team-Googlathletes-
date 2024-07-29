function logInventoryData(form) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Inventory');

  var lastRow = sheet.getLastRow();
  var idValue = 'PRO-' + Utilities.formatString('%03d', lastRow);

  var newRow = [
    idValue,
    form.name,
    form.category,
    form.quantity,
    form.sellingPrice,
    form.unitCost,
    0,
    '',
    form.notes
  ];
  sheet.appendRow(newRow);

  return "Record added successfully!";
}
