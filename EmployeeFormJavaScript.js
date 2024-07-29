function logEmployeeData(form) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Employee');

  var lastRow = sheet.getLastRow();
  var idValue = 'EMP-' + Utilities.formatString('%03d', lastRow);

  var newRow = [
    idValue,
    form.name,
    form.branch,
    form.status,
    form.jobTitle,
    form.department,
    form.salary,
    form.address,
    form.contactNumber,
    form.email,
    form.birthDate,
    form.joiningDate,
    form.exitDate,
    form.bankAccount,
    form.epf,
    form.socso
  ];
  sheet.appendRow(newRow);

  return "Record added successfully!";
}


