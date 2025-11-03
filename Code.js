/**
 * @OnlyCurrentDoc
 */

/**
 * Adds a custom menu to the active document, containing a single menu item
 * for showing the sidebar.
 *
 * @param {object} e The event parameter for a simple onOpen trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode.
 */
function onOpen(e) {
  DocumentApp.getUi()
      .createMenu('Content Assistant')
      .addItem('Show', 'showSidebar')
      .addToUi();
}


/**
 * Shows a sidebar in the document.
 */
function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('Content Assistant');
  DocumentApp.getUi().showSidebar(html);
}

/**
 * Gets the template data from the sheetLibrary.json.html file.
 * @returns {object} The template data as a JSON object.
 */
function getTemplates() {
  const htmlContent = HtmlService.createHtmlOutputFromFile('sheetLibrary.json.html').getContent();
  const jsonString = htmlContent.substring(htmlContent.indexOf('{'), htmlContent.lastIndexOf('}') + 1);
  try {
    const data = JSON.parse(jsonString);
    return data.templates || []; // Return the array directly, or an empty array if it doesn't exist
  } catch (e) {
    console.error("Failed to parse JSON from sheetLibrary.json.html: " + e.message);
    return []; // Return an empty array on failure
  }
}

/**
 * Inserts text at the current cursor position in the document.
 * @param {string} text The text to insert.
 */
function insertText(text) {
  const cursor = DocumentApp.getActiveDocument().getCursor();
  if (cursor) {
    cursor.insertText(text);
  } else {
    DocumentApp.getUi().alert('Could not find a cursor position to insert text.');
  }
}

/**
 * Appends text to the end of the document.
 * @param {string} text The text to append.
 */
function appendText(text) {
  const body = DocumentApp.getActiveDocument().getBody();
  body.appendParagraph(text);
}
