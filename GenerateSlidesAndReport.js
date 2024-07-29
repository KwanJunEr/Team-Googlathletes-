function generateSalesReport() {
  generateSalesSlides();
}

function generateSalesSlides(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SALES_DASHBOARD);
  const charts = sheet.getCharts();

  const slides = SlidesApp.create(sheet.getName());
  const els = slides.getSlides()[0].getPageElements();
  els.forEach((el)=>{
  el.remove();
  })
  slides.getSlides()[0].insertTextBox('Sales Report');
  
  charts.forEach((chart)=>{
  const newSlide = slides.appendSlide();
  newSlide.insertSheetsChart(chart,20,10,600,400);
  })

  const url = slides.getUrl();
  generateAndSendSalesReport(url);
}

function generateAndSendSalesReport(url) {
  const email = Session.getActiveUser().getEmail();
  const subject = 'Sales Report Slides Attached';
  const body = `The Sales Report Slides have been generated. You can view them using the following link: ${url}`;
  
  MailApp.sendEmail(email, subject, body);
}

function extractDocumentId(url) {
  const parts = url.split('=');
  if (parts.length > 1) {
    return parts[1];
  }
  return null;
}



