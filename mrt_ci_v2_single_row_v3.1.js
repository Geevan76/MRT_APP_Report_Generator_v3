// === CONFIGURATION VARIABLES (EASY ACCESS) ===
var dataStartRow = 11;  // This is where data starts in the sheet
var batchRowSize = 200;  // Maximum number of rows to process per batch
var maxExecutionTime = 10 * 60 * 1000;  // 10 minutes in milliseconds
var triggerInterval = 0.2;  // 12 seconds
var maxImageWidth = 100;  // Maximum image width in pixels
var maxImageHeight = 100;  // Maximum image height in pixels

// === REPORT TRIGGER FUNCTIONS ===

function generateVisualReport() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Visual_Cleaned_Report");
  if (sheet) {
    processInspectionData(sheet, 'V');
  } else {
    SpreadsheetApp.getUi().alert("Error: 'Visual_Cleaned_Report' sheet not found.");
  }
}

function generateFunctionalReport() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Functional_Cleaned_Report");
  if (sheet) {
    processInspectionData(sheet, 'F');
  } else {
    SpreadsheetApp.getUi().alert("Error: 'Functional_Cleaned_Report' sheet not found.");
  }
}

// === MAIN SCRIPT FUNCTIONS ===

function processInspectionData(sheet, reportPrefix) {
  Logger.log("Processing sheet: " + sheet.getName());

  // Record start date and time in cell F3
  var startTime = new Date();
  sheet.getRange("F3").setValue(startTime);
  sheet.getRange("F3").setNumberFormat("dd-MM-yy HH:mm:ss");

  // Set the column mapping dynamically based on the report type
  var columnMapping = {};
  if (reportPrefix === 'V') {
    // Visual Inspection Mapping
    columnMapping = {
      '{{Inspection ID}}': 2,
      '{{UserName}}': 5,
      '{{trainNo}}': 7,
      '{{Location}}': 8,
      '{{Car Body}}': 11,
      '{{Section Name}}': 13,
      '{{Subsystem Name}}': 15,
      '{{Serial Number}}': 16,
      '{{Subcomponent}}': 18,
      '{{SubSubcomponent}}': 19,
      '{{Condition}}': 20,
      '{{Defect Type}}': 21,
      '{{Remarks}}': 22,
      '{{Image URL}}': 28
    };
  } else {
    // Functional Inspection Mapping
    columnMapping = {
      '{{Inspection ID}}': 2,
      '{{UserName}}': 5,
      '{{trainNo}}': 7,
      '{{Location}}': 8,
      '{{Car Body}}': 11,
      '{{Section Name}}': 13,
      '{{Subsystem Name}}': 15,
      '{{Serial Number}}': 16,
      '{{Subcomponent}}': 18,
      '{{Condition}}': 19,
      '{{Defect Type}}': 20,
      '{{Remarks}}': 21,
      '{{Image URL}}': 27
    };
  }

  // Fetch the last data row in the sheet
  var lastDataRow = sheet.getLastRow();
  var totalDataRows = lastDataRow - dataStartRow + 1;
  var rowsToProcess = Math.min(batchRowSize, totalDataRows);

  // Read the data starting from dataStartRow and fetch rowsToProcess rows
  var dataRange = sheet.getRange(dataStartRow, 1, rowsToProcess, sheet.getLastColumn());
  var data = dataRange.getValues();

  if (!data || data.length === 0) {
    SpreadsheetApp.getUi().alert("No data found.");
    return;
  }

  // Fetch Start Item No from F6
  var startItemNo = sheet.getRange("F6").getValue();
  if (!startItemNo || isNaN(startItemNo)) {
    SpreadsheetApp.getUi().alert("Error: Missing or invalid Start Item Number in cell F6.");
    return;
  }

  // Process the batch of rows
  generateReportsFromData(data, reportPrefix, startItemNo, columnMapping);

  // Record end date and time in cell G3
  var endTime = new Date();
  sheet.getRange("G3").setValue(endTime);
  sheet.getRange("G3").setNumberFormat("dd-MM-yy HH:mm:ss");

  // Calculate duration and write to cell H3
  var durationMillis = endTime.getTime() - startTime.getTime();
  var diffHours = Math.floor(durationMillis / (1000 * 60 * 60));
  var diffMinutes = Math.floor((durationMillis % (1000 * 60 * 60)) / (1000 * 60));
  var diffSeconds = Math.floor((durationMillis % (1000 * 60)) / 1000);
  var durationFormatted = diffHours + "h " + diffMinutes + "m " + diffSeconds + "s";
  sheet.getRange("H3").setValue(durationFormatted);

  SpreadsheetApp.getUi().alert("Report generation complete.");
}

// === GENERATE REPORTS ===

function generateReportsFromData(dataBatch, reportPrefix, startItemNo, columnMapping) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var groupedData = [];
  var seenIds = {};

  dataBatch.forEach(function(row, rowIndex) {
    var inspectionId = row[columnMapping['{{Inspection ID}}'] - 1];
    var imageUrl = row[columnMapping['{{Image URL}}'] - 1];
    var timestamp = new Date(row[columnMapping['{{Inspection Timestamp}}'] - 1]);

    if (!seenIds[inspectionId]) {
      groupedData.push({
        inspectionId: inspectionId,
        row: row,
        rowIndex: rowIndex,
        timestamp: timestamp
      });
      seenIds[inspectionId] = row;
    }

    if (imageUrl && imageUrl.trim() !== "") {
      groupedData[groupedData.length - 1].row = row;
      groupedData[groupedData.length - 1].timestamp = timestamp;
    }
  });

  groupedData.sort(function(a, b) {
    if (a.rowIndex !== b.rowIndex) {
      return a.rowIndex - b.rowIndex;
    }
    return a.timestamp - b.timestamp;
  });

  var templateId = sheet.getRange("B4").getValue();
  var doc = createDocumentFromTemplate(templateId, "Inspection_Report");

  groupedData.forEach(function(group, index) {
    var row = group.row;
    var tableRow = doc.getBody().appendTable().appendTableRow();
    tableRow.appendTableCell((startItemNo + index).toString());
    tableRow.appendTableCell(row[columnMapping['{{Location}}'] - 1].toString());
    tableRow.appendTableCell(row[columnMapping['{{Car Body}}'] - 1].toString());
    tableRow.appendTableCell(row[columnMapping['{{UserName}}'] - 1].toString());
    tableRow.appendTableCell(row[columnMapping['{{Section Name}}'] - 1].toString());
    tableRow.appendTableCell(row[columnMapping['{{Subsystem Name}}'] - 1].toString());
    tableRow.appendTableCell(row[columnMapping['{{Serial Number}}'] - 1].toString());
    tableRow.appendTableCell(row[columnMapping['{{Subcomponent}}'] - 1].toString());
    tableRow.appendTableCell(row[columnMapping['{{Condition}}'] - 1].toString());
    tableRow.appendTableCell(row[columnMapping['{{Defect Type}}'] - 1].toString());
    tableRow.appendTableCell(row[columnMapping['{{Remarks}}'] - 1].toString());

    if (reportPrefix === 'V') {
      tableRow.appendTableCell(row[columnMapping['{{SubSubcomponent}}'] - 1].toString());
    }

    var imageUrl = row[columnMapping['{{Image URL}}'] - 1].toString();
    var imageCell = tableRow.appendTableCell();

    if (imageUrl && imageUrl.trim() !== "") {
      try {
        var response = UrlFetchApp.fetch(imageUrl);
        var imageBlob = response.getBlob();
        if (imageBlob.getContentType().indexOf("image") !== -1) {
          imageCell.appendImage(imageBlob).setWidth(maxImageWidth).setHeight(maxImageHeight);
        } else {
          imageCell.appendParagraph("Invalid image content");
        }
      } catch (e) {
        imageCell.appendParagraph("Error fetching image: " + e.message);
      }
    } else {
      imageCell.appendParagraph("No image available");
    }
  });

  doc.saveAndClose();
  saveDocumentToFolder(doc, sheet.getName(), groupedData[0].row[columnMapping['{{trainNo}}'] - 1]);
}

// === DOCUMENT CREATION AND MANAGEMENT ===

function createDocumentFromTemplate(templateId, fileName) {
  var templateDoc = DriveApp.getFileById(templateId);
  var copyDoc = templateDoc.makeCopy(fileName);
  return DocumentApp.openById(copyDoc.getId());
}

function saveDocumentToFolder(doc, sheetName, trainNo) {
  var spreadsheetFile = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
  var parentFolder = spreadsheetFile.getParents().next();
  var reportFolderName = (sheetName === "Visual_Cleaned_Report") ? "Visual_Inspection_Reports" : "Functional_Inspection_Reports";
  var reportFolder = getOrCreateFolder(parentFolder, reportFolderName);
  var trainFolder = getOrCreateFolder(reportFolder, trainNo);
  var file = DriveApp.getFileById(doc.getId());
  file.moveTo(trainFolder);
}

function getOrCreateFolder(parentFolder, folderName) {
  var folders = parentFolder.getFoldersByName(folderName);
  return folders.hasNext() ? folders.next() : parentFolder.createFolder(folderName);
}
