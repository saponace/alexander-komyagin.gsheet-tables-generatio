function populateMap() {
  var _outMap = {}
  _outMap["Group A"] = {}
  _outMap["Group B"] = {}
  _outMap["Group A"]["Part I"] = [["Lorem ipsum: lorem ipsum","Some text [optional text]"],["Some text [A]","N/A"]]
  _outMap["Group A"]["Part II"] = [["Lorem ipsum: lorem ipsum","Some text"],["Some text","N/A"]]
  _outMap["Group B"]["Part I"] = [["Lorem ipsum: lorem ipsum","Some [A/B] text [more optional text]"],["Some text","[optional text] AAA"]]

  Logger.log(_outMap);
  return _outMap;
}


/**
 * Init script data
 */
function init() {
  this.insertParagraphName =  "2.1 Items";
  this.tableStyles = {
    header1: {
      backgroudColor: "#999999",
      foregroundColor: "#FFFFFF",
    },
    header2: {
      backgroudColor: "#E3E3E3",
      foregroundColor: "#000000",
    }
  };
  this.beforeColonRegex = '^([^:]+):';
  this.insideSquareBracketsRegex = '\\[([^\\]]+)\\]';
  this.squareBracketsHighlightColor = "#FFFF00";
  this.driveDirectoryToCreateDoc = '198XnJHfK_7l8_SM4pWx2VUBEmCXzudSr';
}

/**
 * Main function. Generates a copy of the template and injects the tables into it
 */
function generateDoc() {
  init();
  var documentCopyId = copyTemplate();
  var documentBody = DocumentApp.openById(documentCopyId).getBody();
  this.documentBody = documentBody;
  var insertParagraph = getInsertParagraph();
  appendTables(insertParagraph, populateMap());
  applyBeforeColonStyle(documentBody);
  applySquareBracketsStyle(documentBody);
}

/**
 * Create a copy of the template in the specified directory
 * @return The copy's id
 */
function copyTemplate() {
  var currentDocument = DocumentApp.getActiveDocument();
  var templateId = currentDocument.getId();
  var templateFile = DriveApp.getFileById(templateId);
  var destinationDirectory = DriveApp.getFolderById(this.driveDirectoryToCreateDoc);
  return templateFile.makeCopy(currentDocument.getName(), destinationDirectory).getId();
}

/**
 * Get the paragraph tables should be inserted into
 * @returns The paragraph object if it exists, null otherwise
 */
function getInsertParagraph() {
  var paragraphs = this.documentBody.getParagraphs();
  for (var i = 0; i < paragraphs.length; i++) {
    var p = paragraphs[i];
    if(p.getHeading() === DocumentApp.ParagraphHeading.HEADING2 && p.getText() === this.insertParagraphName){
      return p;
    }
  }
  return null;
}

/**
 * Apply the given style to text matching the given regex
 * @param element The element in which the search should be
 * @param pattern The regex
 * @param style The style to apply
 */
function applyStyle(element, pattern, style) {
  var found = element.findText(pattern);
  while (found) {
    found.getElement().setAttributes(found.getStartOffset(), found.getEndOffsetInclusive(), style);
    found = element.findText(pattern, found);
  }
}

/**
 * Set text before the FIRST colon in bold (including the colon itself) in the given GDocs element
 * @param element The element in which the analysis should be performed
 */
function applyBeforeColonStyle(element) {
  var style = {};
  style[DocumentApp.Attribute.BOLD] = true;
  applyStyle(element, this.beforeColonRegex, style);
}

/**
 * Apply yellow highlight of all text in square brackets (including the square brackets) in the given GDocs element
 * @param element The element in which the analysis should be performed
 */
function applySquareBracketsStyle(element) {
  var style = {};
  style[DocumentApp.Attribute.BACKGROUND_COLOR] = this.squareBracketsHighlightColor;
  applyStyle(element, this.insideSquareBracketsRegex, style);
}


/**
 * Append tables to the paragraph
 * @param paragraph The paragraph in which tables should be inserted
 * @param inputMap The data structure containing text to insert in the tables
 */
function appendTables(paragraph, inputMap) {
  var tableCounter = 0;
  var tablesToInsert = [];
  for (var key in inputMap) {
    if (inputMap.hasOwnProperty(key)) {
      tableCounter++;
      var table = createTable(key, inputMap[key]);
      var sectionName = getParagraphNumber(paragraph) + "." + tableCounter + " " + key;
      tablesToInsert.push({paragraph: paragraph, sectionName: sectionName, table: table});
    }
  }
  // Iterate trough the array in reverse to insert tables becaus ethey are inserted
  // right after the paragraph heading (so before any other table)
  for (var i = tablesToInsert.length - 1; i >= 0; i--) {
    var t = tablesToInsert[i];
    appendTable(t.paragraph, t.sectionName, t.table);
  }
}

/**
 * Get the numbering of the paragraph
 * @param paragraph The paragraph
 * @returns The section number
 */
function getParagraphNumber(paragraph) {
  return paragraph.getText().split(' ')[0];
}

/**
 * Create a table suitable for inserting into GDocs
 * @param tableTitle The title of the table
 * @param tableContent The content of the table
 * @returns {{headersIndexes: {"1": integer, "2": Array}, table: Array}} headersIndexes: Indexes of table headers
 */
function createTable(tableTitle, tableContent) {
  var retVal = [];
  var headersIndexes = {
    1: null,
    2: []
  };
  retVal.push(["Key", "Value"]);
  headersIndexes[1] = 0;
  var row = 0;
  for (var key in tableContent) {
    row++;
    if (tableContent.hasOwnProperty(key)) {
      retVal.push([key, ""]);
      headersIndexes[2].push(row);
      retVal = retVal.concat(tableContent[key]);
      row += tableContent[key].length;
    }
  }
  return {
    headersIndexes: headersIndexes,
    table: retVal
  };
}


/**
 * Append a table to the document
 * @param paragraph The paragraph in which the table should be inserted
 * @param tableName The name of the table to insert
 * @param tableAndHeadersIndexes object returned by createTable()
 */
function appendTable(paragraph, tableName, tableAndHeadersIndexes) {
  var table = tableAndHeadersIndexes.table;
  var headersIndexes = tableAndHeadersIndexes.headersIndexes;
  var paragraphIndex = paragraph.getParent().getChildIndex(paragraph);
  var paragraphInsertPoint = paragraphIndex + 1;
  var insertedParagraph = this.documentBody.insertParagraph(paragraphInsertPoint, tableName);
  this.documentBody.insertParagraph(paragraphInsertPoint+1, "");
  setHeader3Style(insertedParagraph);
  var insertedParagraphIndex = insertedParagraph.getParent().getChildIndex(paragraph);
  var tableInsertPoint = insertedParagraphIndex + 3;
  var table = insertedParagraph.getParent().insertTable(tableInsertPoint, table);
  this.documentBody.insertParagraph(tableInsertPoint+1, "");

  applyStyleToRow(table.getRow(0), setTableHeader1Style);
  for (var i = 1; i < table.getNumRows(); i++) {
    var styleToApply = null;
    if(headersIndexes[2].indexOf(i) !== -1)
      styleToApply = setTableHeader2Style;
    else
      styleToApply = setTableBodyStyle;
    applyStyleToRow(table.getRow(i), styleToApply);
  }
}

/**
 * Set the style for header3 on a paragraph
 * @param paragraph The paragraph
 */
function setHeader3Style(paragraph) {
  paragraph.setHeading(DocumentApp.ParagraphHeading.HEADING3);
}

/**
 * Apply style to a table row
 * @param row The row which the style should be applied to
 * @param styleApplyFunction The function that applies style to a cell
 */
function applyStyleToRow(row, styleApplyFunction) {
  for (var i = 0; i < row.getNumCells(); i++) {
    styleApplyFunction(row.getCell(i));
  }
}

/**
 * Apply the body style to a cell
 * @param cell The cell
 */
function setTableBodyStyle(cell) {
  var style = {};
  style[DocumentApp.Attribute.FONT_SIZE] = 9;
  style[DocumentApp.Attribute.PADDING_BOTTOM] = 2;
  style[DocumentApp.Attribute.PADDING_TOP] =  2;
  style[DocumentApp.Attribute.PADDING_LEFT] = 2.5;
  cell.setAttributes(style);
}

/**
 * Apply the header 1 style to a cell
 * @param cell The cell
 */
function setTableHeader1Style(cell) {
  cell.setBackgroundColor(this.tableStyles.header1.backgroudColor);
  var style = {};
  style[DocumentApp.Attribute.FOREGROUND_COLOR] = this.tableStyles.header1.foregroundColor;
  style[DocumentApp.Attribute.BOLD] = true;
  style[DocumentApp.Attribute.FONT_SIZE] = 11;
  style[DocumentApp.Attribute.PADDING_BOTTOM] = 2;
  style[DocumentApp.Attribute.PADDING_TOP] =  2;
  style[DocumentApp.Attribute.PADDING_LEFT] = 2.5;
  cell.setAttributes(style);
}

/**
 * Apply the header 2 style to a cell
 * @param cell The cell
 */
function setTableHeader2Style(cell) {
  cell.setBackgroundColor(this.tableStyles.header2.backgroudColor);
  var style = {};
  style[DocumentApp.Attribute.FOREGROUND_COLOR] = this.tableStyles.header2.foregroundColor;
  style[DocumentApp.Attribute.BOLD] = true;
  style[DocumentApp.Attribute.FONT_SIZE] = 11;
  style[DocumentApp.Attribute.PADDING_BOTTOM] = 2;
  style[DocumentApp.Attribute.PADDING_TOP] =  2;
  style[DocumentApp.Attribute.PADDING_LEFT] = 2.5;
  cell.setAttributes(style);
}
