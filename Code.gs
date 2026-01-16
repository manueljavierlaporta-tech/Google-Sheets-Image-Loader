function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Extra Functions")
    .addItem("Upload Images", "showScreen")
    .addToUi();
}

function showScreen() {
  var template = HtmlService.createTemplateFromFile("createInsertionForm");
  var html = template.evaluate()
    .setTitle("ShutterSheets")
    .setWidth(1000)
    .setHeight(1000);

  SpreadsheetApp.getUi().showModalDialog(html, 'ShutterSheets');
}

function newRow(allValuesArray){
  SpreadsheetApp.getActiveSheet().appendRow(allValuesArray);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
