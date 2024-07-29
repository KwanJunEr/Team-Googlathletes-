function logSaleData(form) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sales');

  // Generate ID
  var lastRow = sheet.getLastRow();
  var idValue = 'SALE-' + Utilities.formatString('%03d', lastRow);

  form.products.forEach(function (product) {
    var price = getProductPrice(product.product);
    var cost = getProductCost(product.product);
    var stockLevel = getProductStock(product.product) - product.quantity;

    setProductStock(product.product, stockLevel);

    if (price !== null) {
      // Calculate sub-total
      var subTotal = parseFloat(price) * parseFloat(product.quantity);
      var taxRate = 1 + (parseFloat(getSSTRate() + getGSTRate())/100);
      var total = parseFloat(subTotal * taxRate);

      // Log the form data along with the calculated sub-total
      var logData = [
        idValue,
        form.priority,
        form.branch,
        form.customer,
        form.date,
        form.time,
        form.paymentDate,
        form.status,
        form.salesChannel,
        product.product,
        cost,
        price,
        product.quantity,
        stockLevel,
        form.paymentMethod,
        subTotal, // Use calculated sub-total here
        total, // Use calculated total here
        form.notes
      ];
      sheet.appendRow(logData);
      }
  });
  
  return "Record added successfully!";
}
function logSaleData(form) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sales');

  // Generate ID
  var lastRow = sheet.getLastRow();
  var idValue = 'SALE-' + Utilities.formatString('%03d', lastRow);

  form.products.forEach(function (product) {
    var price = getProductPrice(product.product);
    var cost = getProductCost(product.product);
    var stockLevel = getProductStock(product.product) - product.quantity;

    setProductStock(product.product, stockLevel);

    if (price !== null) {
      // Calculate sub-total
      var subTotal = parseFloat(price) * parseFloat(product.quantity);
      var taxRate = 1 + (parseFloat(getSSTRate() + getGSTRate())/100);
      var total = parseFloat(subTotal * taxRate);

      // Log the form data along with the calculated sub-total
      var logData = [
        idValue,
        form.priority,
        form.branch,
        form.customer,
        form.date,
        form.time,
        form.paymentDate,
        form.status,
        form.salesChannel,
        product.product,
        cost,
        price,
        product.quantity,
        stockLevel,
        form.paymentMethod,
        subTotal, // Use calculated sub-total here
        total, // Use calculated total here
        form.notes
      ];
      sheet.appendRow(logData);
      }
  });
  
  return "Record added successfully!";
}
