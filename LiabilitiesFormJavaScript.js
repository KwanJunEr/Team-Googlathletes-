function logLiabilitiesData(form) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Liabilities');

  var lastRow = sheet.getLastRow();
  var idValue = 'LIB-' + Utilities.formatString('%03d', lastRow);

  var taxable = 'Yes';
  if (form.taxable == false) {
    taxable = 'No';
  }

  var newRow = [
    idValue,
    form.name,
    form.account,
    form.creditor,
    form.category,
    form.initialAmount,
    form.currentBalance,
    form.interestRate,
    form.frequency,
    form.minPayment,
    form.dueDate,
    form.lastPayment,
    form.lastPaymentAmount,
    form.nextPaymentAmount,
    form.startDate,
    form.maturityDate,
    form.currency,
    taxable,
    form.notes,
    form.status
  ];
  sheet.appendRow(newRow);

  return "Record added successfully!";
}
