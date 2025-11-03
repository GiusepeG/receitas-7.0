function onOpen() {
  DocumentApp.getUi()
      .createMenu('Content Assistant')
      .addItem('Show', 'showSidebar')
      .addToUi();
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('Content Assistant for Medical Professionals');
  DocumentApp.getUi().showSidebar(html);
}

function getTemplates() {
  return HtmlService.createHtmlOutputFromFile('sheetLibrary.json.html').getContent();
}

function insertText(text) {
  var cursor = DocumentApp.getActiveDocument().getCursor();
  if (cursor) {
    cursor.insertText(text);
  } else {
    throw new Error('Please place your cursor in the document to insert text.');
  }
}