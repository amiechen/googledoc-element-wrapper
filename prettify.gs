var selection = DocumentApp.getActiveDocument().getSelection();

function onOpen(e) {
 DocumentApp.getUi().createAddonMenu()
      .addItem('Add Text Box', 'addTextBox')
      .addToUi();
}

function onInstall(e) {
  onOpen(e);
}

function addTextBox () {
  var body = DocumentApp.getActiveDocument().getBody();

  if (selection) {
    var tableStyle = {};
    var cellStyle = {};
    var selectedText = getSelectedText();
    var firstIndex;
    var table;
    var cell;

    firstIndex = replaceSelection();
    table = body.insertTable(firstIndex);
    cell = table.appendTableRow().appendTableCell([[selectedText]]);

    tableStyle[DocumentApp.Attribute.BORDER_COLOR] = '#d9d9d9';
    tableStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Consolas';
    tableStyle[DocumentApp.Attribute.FONT_SIZE] = 9;
    cellStyle[DocumentApp.Attribute.BACKGROUND_COLOR] = '#f5f5f5';

    table.setAttributes(tableStyle);
    cell.setAttributes(cellStyle);
  }
}

function getSelectedText() {
  var elements = selection.getRangeElements();
  var text = [];

  for (var i = 0; i < elements.length; i++) {
    var element = elements[i].getElement().editAsText().getText();
    text.push(element);
  }
  return text.join('\r');
}

function replaceSelection() {
  var firstIndex;
  if (selection) {
    var elements = selection.getRangeElements();

    for (var i = 0; i < elements.length; i++) {
      if (elements[i].isPartial()) {
        var element = elements[i].getElement().asText();
        var startIndex = elements[i].getStartOffset();
        var endIndex = elements[i].getEndOffsetInclusive();
        var text = element.getText().substring(startIndex, endIndex + 1);

        if (i === 0) {
          firstIndex = DocumentApp.getActiveDocument().getBody().getChildIndex(elements[i].getElement().getParent());
        }

        element.deleteText(startIndex, endIndex);
      } else {
        var element = elements[i].getElement();
        if (i === 0) {
          firstIndex = DocumentApp.getActiveDocument().getBody().getChildIndex(element);
        }
        element.removeFromParent();
      }
    }
  }
  return firstIndex;
}
