// Provide with specific values
var TEMPLATE = "";
var SOURCE = "";
var TARGET_FOLDER = "";

var TIMESTAMP_FORMAT = "yyyyMMddHHmmss";

RegExp.quote = function(str) {
  return (str+'').replace(/[.?*+^$[\]\\(){}|-]/g, "\\$&");
};

function getTimestamp() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), TIMESTAMP_FORMAT);
}

function parseInternalFormatting(paragraph) {
  var paragraphText = paragraph.editAsText();
  var paragraphTextString = paragraphText.getText();
  var regexp = new RegExp("<(.)>(.+?)</\\1>");
  
  var match = regexp.exec(paragraphTextString);  
  while (match != null) {
    Logger.log("match=" + match);
    var r = paragraphText.findText(RegExp.quote(match[2]));
    if (r != null) {
      if (match[1] == "b") {
        paragraphText.setBold(r.getStartOffset(), r.getEndOffsetInclusive(), true);
      } else if (match[1] == "i") {
        paragraphText.setItalic(r.getStartOffset(), r.getEndOffsetInclusive(), true);
      }
    }
    paragraphTextString = paragraphTextString.replace("<" + match[1] + ">", "").replace("</" + match[1] + ">", "");
    match = regexp.exec(paragraphTextString);
  }
  // Remove tags
  paragraphText.replaceText("</?[bi]>", "");
  // Add new lines
  paragraphText.replaceText("<br/>", "\n");
}

function insertBody(source, destination) {
  for(var i = 0; i < source.getNumChildren(); i++) {
    var element = source.getChild(i).copy();
    var type = element.getType();
    if (type == DocumentApp.ElementType.PARAGRAPH) {
      Logger.log(element.getText());
      parseInternalFormatting(element);
      destination.appendParagraph(element);
    } else if (type == DocumentApp.ElementType.TABLE) {
      destination.appendTable(element);
    } else if (type == DocumentApp.ElementType.LIST_ITEM) {
      destination.appendListItem(element);
    } else {
      throw new Error("According to the doc this type couldn't appear in the body: "+type);
    }
  }
}

function populateTemplate(templateBody, columns, values) {
  var body = templateBody.copy();
  
  for (var q = 0; q < columns.length; q++) {
    if (values[q] != null) {
      body.replaceText(":" + columns[q] + ":", values[q]);
    } else {
      body.replaceText(":" + columns[q] + ":", "");
    }
  }
  
  return body;
}

function getNumPages(doc) {
  var blob = doc.getAs("application/pdf");
  var data = blob.getDataAsString();
  var pages = parseInt(data.match(/ \/N (\d+) /)[1], 10);
  // Logger.log("data = " + data);
  Logger.log("pages = " + pages);
  return pages; 
}

function getRowAsArray(sheet, row) {
  var dataRange = sheet.getRange(row, 1, 1, sheet.getMaxColumns());
  var data = dataRange.getValues();
  var columns = [];

  for (i in data) {
    var row = data[i];

    for(var l = 0; l < sheet.getMaxColumns(); l++) {
        columns.push(row[l]);
    }
  }

  return columns;
}

function generate() {

  var sourceSpreadSheet = SpreadsheetApp.openById(SOURCE);
  var sourceSheet = SpreadsheetApp.getActiveSheet();
  Logger.log("sourceSheet=" + sourceSheet.getName());
  
  var columns = getRowAsArray(sourceSheet, 1);
  Logger.log("columns=" + columns);
  if (columns.length == 0) {
    // Empty
    Logger.log("First row is empty");
    return;
  }
  
  // Create new document
  var templateDoc = DocumentApp.openById(TEMPLATE);
  var newFile;
  if (TARGET_FOLDER != null) {
    targetFolder = DriveApp.getFolderById(TARGET_FOLDER);
    newFile = DriveApp.getFileById(TEMPLATE).makeCopy(sourceSheet.getName()+"_"+getTimestamp(), DriveApp.getFolderById(TARGET_FOLDER));
  } else {
    newFile = DriveApp.getFileById(TEMPLATE).makeCopy(sourceSheet.getName()+"_"+getTimestamp());
  }
  var docId = newFile.getId();
  var doc = DocumentApp.openById(docId);
  
  var templateBody = templateDoc.getBody().copy();
  var destBody = doc.getBody();
  var tmpBody;
  destBody.clear();
  var r = 1;
  var values = getRowAsArray(sourceSheet, r + 1);
  while (values.length != 0 && values[0]) {
    Logger.log("Processing data row " + r);
    Logger.log("values=" + values);
    Logger.log("Populating template for data row " + r);
    tmpBody = populateTemplate(templateBody, columns, values);
    Logger.log("Copying populated baody for data row " + r);
    insertBody(tmpBody, destBody);
    destBody.appendPageBreak();
    if (r == 1) {
      // This removes first empty paragraph
      destBody.removeChild(destBody.getChild(0));
    }
    // TODO below causes performance issues
    // Check if number of pages is even
    /*doc.saveAndClose();
    var pages = getNumPages(doc);
    doc = DocumentApp.openById(docId);
    destBody = doc.getBody();
    if (pages % 2 != 0) {
      destBody.appendPageBreak();
    }*/
    // Continue loop
    r++;
    values = getRowAsArray(sourceSheet, r + 1);
    Logger.log("r="+(r % 100));
    if ((r % 50) == 0) {
      doc.saveAndClose();
      doc = DocumentApp.openById(docId);
      destBody = doc.getBody();
    }
  }
}
