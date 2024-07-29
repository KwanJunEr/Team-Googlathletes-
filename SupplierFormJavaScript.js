function logSupplierData(form) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Supplier');

  var lastRow = sheet.getLastRow();
  var idValue = 'SUP-' + Utilities.formatString('%03d', lastRow);

  var newRow = [
    idValue,
    form.name,
    form.category,
    form.contactNumber,
    form.email,
    form.website,
    'N/A',
    form.notes
  ];
  sheet.appendRow(newRow);

  return "Record added successfully!";
}
