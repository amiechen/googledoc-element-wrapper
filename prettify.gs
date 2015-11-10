var selection = DocumentApp.getActiveDocument().getSelection();

function onOpen(e) {
 DocumentApp.getUi().createAddonMenu()
      .addItem('Add Text Box', 'addTextBox')
      .addItem('Add Code Block', 'showSidebar')
      .addToUi();
}

function onInstall(e) {
  onOpen(e);
}

function addTextBox () {
  var body = DocumentApp.getActiveDocument().getBody();

  if (selection) {
    var tableStyle = {},
        cellStyle = {};

    tableStyle[DocumentApp.Attribute.BORDER_COLOR] = '#d9d9d9';
    tableStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Consolas';
    tableStyle[DocumentApp.Attribute.FONT_SIZE] = 9;
    cellStyle[DocumentApp.Attribute.BACKGROUND_COLOR] = '#f5f5f5';

    // TODO: instead of append table to body, delete all selected text then draw the table
    var table = body.appendTable(),
        cell = table.appendTableRow().appendTableCell([[getSelectedText()]]);

    table.setAttributes(tableStyle);
    cell.setAttributes(cellStyle);
  }
}

function showSidebar () {
  var ui = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('Code Prettifier')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  DocumentApp.getUi().showSidebar(ui);
}

function getSelectedText() {
  var elements = selection.getRangeElements(),
      text = [];

  for (var i = 0; i < elements.length; i++) {
    var element = elements[i].getElement().editAsText().getText();
    text.push(element);
  }
  return text.join('\r');
}

