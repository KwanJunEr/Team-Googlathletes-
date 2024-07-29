function logAssetsData(form) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Assets');

  var lastRow = sheet.getLastRow();
  var idValue = 'ASS-' + Utilities.formatString('%03d', lastRow);

  var newRow = [
    idValue,
    form.name,
    form.category,
    form.purchaseDate,
    form.cost,
    form.currentValue,
    form.depreciationRate + '%',
    form.salvageValue,
    form.usefulYears
  ];
  sheet.appendRow(newRow);

  return "Record added successfully!";
}
