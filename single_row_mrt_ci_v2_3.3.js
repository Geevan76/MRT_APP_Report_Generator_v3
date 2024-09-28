// === CONFIGURATION VARIABLES (EASY ACCESS) ===
var dataStartRow = 11;  // This is where data starts in the sheet
var batchRowSize = 200;  // Maximum number of rows to process per batch
var maxExecutionTime = 10 * 60 * 1000;  // 10 minutes in milliseconds
var triggerInterval = 0.2;  // 12 seconds
var maxImageWidth = 100;  // Maximum image width in pixels
var maxImageHeight = 100;  // Maximum image height in pixels

// === REPORT TRIGGER FUNCTIONS ===

function generateVisualReport() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Report (Visual)");
  if (sheet) {
    processInspectionData(sheet, 'V');
  } else {
    SpreadsheetApp.getUi().alert("Error: 'Visual_Cleaned_Report' sheet not found.");
  }
}

function generateFunctionalReport() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Report (Functional)");
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
      '{{SubSubcomponent}}': 22, // Used for Visual reports
      '{{Condition}}': 19,
      '{{Defect Type}}': 20,
      '{{Remarks}}': 21,
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

  // Get Start Item No from F6
  var startItemNo = sheet.getRange("F6").getValue();
  if (!startItemNo || isNaN(startItemNo)) {
    SpreadsheetApp.getUi().alert("Error: Missing or invalid Start Item Number in cell F6.");
    return;
  }

  // Group data by Inspection ID and drop non-image rows if images exist
  var groupedResult = groupDataByInspectionID(data, columnMapping);
  var inspections = groupedResult.inspections;
  var inspectionOrder = groupedResult.inspectionOrder;

  // Generate the report and get endItemNo
  var endItemNo = generateConsolidatedReport(inspections, inspectionOrder, reportPrefix, startItemNo, columnMapping);

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

// === GROUP AND PROCESS DATA ===

function groupDataByInspectionID(data, columnMapping) {
  var inspections = {};
  var inspectionOrder = [];
  
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var primaryVInspectionID = row[columnMapping['{{Inspection ID}}'] - 1];
    var imageRelatedVInspectionID = row[columnMapping['{{Image VInspectionID}}'] - 1];

    var VInspectionID = primaryVInspectionID ? primaryVInspectionID.toString().trim() : '';
    var imageVInspectionID = imageRelatedVInspectionID ? imageRelatedVInspectionID.toString().trim() : '';
    
    // Use imageVInspectionID if present; otherwise, use primaryVInspectionID
    var keyVInspectionID = imageVInspectionID !== '' ? imageVInspectionID : VInspectionID;

    // If keyVInspectionID is not yet in inspections, add it and record its order
    if (!inspections[keyVInspectionID]) {
      inspections[keyVInspectionID] = { inspectionRow: null, imageRows: [] };
      inspectionOrder.push(keyVInspectionID);
    }

    var imageUrl = row[columnMapping['{{Image URL}}'] - 1];
    imageUrl = imageUrl ? imageUrl.toString().trim() : '';

    if (imageUrl !== '') {
      // This is an image row
      inspections[keyVInspectionID].imageRows.push(row);
    } else {
      // This is an inspection-only row, only keep if no images exist for this inspection ID
      if (inspections[keyVInspectionID].imageRows.length === 0) {
        inspections[keyVInspectionID].inspectionRow = row;
      }
    }
  }

  return { inspections: inspections, inspectionOrder: inspectionOrder };
}

// === GENERATE CONSOLIDATED REPORT ===

function generateConsolidatedReport(inspections, inspectionOrder, reportPrefix, startItemNo, columnMapping) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Get Train Number
  var trainNo = sheet.getRange(dataStartRow, columnMapping['{{trainNo}}']).getValue();
  trainNo = trainNo ? trainNo.toString().trim() : "Unknown_Train_No";

  var rowsToAdd = [];
  var currentItemNo = startItemNo;

  // Prepare data rows
  for (var i = 0; i < inspectionOrder.length; i++) {
    var inspectionID = inspectionOrder[i];
    var inspectionData = inspections[inspectionID];
    var inspectionRow = inspectionData.inspectionRow;
    var imageRows = inspectionData.imageRows;

    if (imageRows.length === 0 && inspectionRow) {
      // No images, include inspectionRow as is
      rowsToAdd.push({ row: inspectionRow, itemNo: currentItemNo });
      currentItemNo++;
    } else if (imageRows.length > 0) {
      // Has images, only include imageRows, discard non-image row
      for (var j = 0; j < imageRows.length; j++) {
        rowsToAdd.push({ row: imageRows[j], itemNo: currentItemNo });
        currentItemNo++;
      }
    }
  }

  // Check if there's data to process
  if (rowsToAdd.length === 0) {
    SpreadsheetApp.getUi().alert("No data rows to process for report generation.");
    return startItemNo - 1;
  }

  var endItemNo = currentItemNo - 1;

  // Construct the filename with both startItemNo and endItemNo
  var fileName = reportPrefix + "_Inspection_Report_for_" + trainNo + "_" + startItemNo + "_" + endItemNo;

  // Get Template ID from B4
  var templateId = sheet.getRange("B4").getValue();
  if (!templateId) {
    SpreadsheetApp.getUi().alert("Error: Missing Template ID in cell B4.");
    return startItemNo - 1;
  }

  var doc = createDocumentFromTemplate(templateId, fileName);

  // Replace {{trainNo}} in header
  replaceTrainNoPlaceholder(doc, trainNo, reportPrefix);

  // Append data rows to document
  for (var k = 0; k < rowsToAdd.length; k++) {
    var item = rowsToAdd[k];
    appendTableRowToDocument(doc.getBody(), item.row, item.itemNo, columnMapping, reportPrefix);
  }

  // Remove the first row (placeholder row) from the table after data insertion
  var table = doc.getBody().getTables()[0];
  if (table && table.getNumRows() > 1) {
    table.removeRow(0);  // Remove the first row, which contains placeholders
    Logger.log("Placeholder row removed.");
  } else {
    Logger.log("No placeholder row to remove.");
  }

  // Save and move the document
  doc.saveAndClose();
  saveDocumentToFolder(doc, sheet.getName(), trainNo);

  // Store the file name in cell H6
  sheet.getRange("H6").setValue(fileName);

  // Get the URL of the Google Doc and store it in cell I6
  var fileUrl = doc.getUrl();
  sheet.getRange("I6").setValue(fileUrl);

  // Log the processing details
  Logger.log("Batch processed. Start Item No: " + startItemNo + ", End Item No: " + endItemNo);

  return endItemNo;
}

// === HELPER FUNCTIONS ===

function appendTableRowToDocument(body, row, itemNo, columnMapping, reportPrefix) {
  Logger.log("Appending Row Data: " + JSON.stringify(row));  // Log entire row data for debugging

  var table = body.getTables()[0]; // Assumes first table
  if (!table) {
    Logger.log("No table found in the document.");
    return;
  }

  var newRow = table.appendTableRow();
  var headerRow = table.getRow(0);
  var numColumns = headerRow.getNumCells();

  // Ensure the new row has the correct number of cells
  while (newRow.getNumCells() < numColumns) {
    newRow.appendTableCell();
  }

  // Populate cells with data
  newRow.getCell(0).setText(itemNo.toString());  // Item No
  newRow.getCell(1).setText((row[columnMapping['{{Location}}'] - 1] || "").toString());        // Location
  newRow.getCell(2).setText((row[columnMapping['{{Car Body}}'] - 1] || "").toString());        // Car Body
  newRow.getCell(3).setText((row[columnMapping['{{UserName}}'] - 1] || "").toString());        // UserName
  newRow.getCell(4).setText((row[columnMapping['{{Section Name}}'] - 1] || "").toString());    // Section Name
  newRow.getCell(5).setText((row[columnMapping['{{Subsystem Name}}'] - 1] || "").toString());  // Subsystem Name
  newRow.getCell(6).setText((row[columnMapping['{{Serial Number}}'] - 1] || "").toString());   // Serial Number
  newRow.getCell(7).setText((row[columnMapping['{{Subcomponent}}'] - 1] || "").toString());    // Subcomponent
  newRow.getCell(8).setText((row[columnMapping['{{Condition}}'] - 1] || "").toString());       // Condition
  newRow.getCell(9).setText((row[columnMapping['{{Defect Type}}'] - 1] || "").toString());     // Defect Type
  newRow.getCell(10).setText((row[columnMapping['{{Remarks}}'] - 1] || "").toString());        // Remarks

  if (reportPrefix === 'V') {
    newRow.getCell(11).setText((row[columnMapping['{{SubSubcomponent}}'] - 1] || "").toString());  // SubSubcomponent (Visual Reports)
  }

  // Handle image insertion
  var imageUrl = row[columnMapping['{{Image URL}}'] - 1];
  var imageCell = newRow.getCell(numColumns - 1); // Last cell for Image
  imageCell.clear(); // Clear existing content

  if (imageUrl && imageUrl.trim() !== "") {
    try {
      var response = UrlFetchApp.fetch(imageUrl);
      var imageBlob = response.getBlob();

      if (imageBlob.getContentType().indexOf("image") !== -1) {
        imageCell.appendImage(imageBlob).setWidth(maxImageWidth).setHeight(maxImageHeight);
      } else {
        imageCell.setText("Invalid image content");
      }
    } catch (e) {
      imageCell.setText("Error fetching image");
      Logger.log("Error fetching image URL: " + imageUrl + " - " + e.message);
    }
  } else {
    imageCell.setText("No image available");
  }
}

function replaceTrainNoPlaceholder(doc, trainNo, reportPrefix) {
  var inspectionType = (reportPrefix === 'F') ? 'Functional Inspection' : 'Visual Inspection';
  var fullTrainNo = trainNo + ' (' + inspectionType + ')';

  var header = doc.getHeader();
  if (header) {
    header.replaceText('{{trainNo}}', fullTrainNo);
  }
}

function createDocumentFromTemplate(templateId, fileName) {
  var templateDoc = DriveApp.getFileById(templateId);
  var copyDoc = templateDoc.makeCopy(fileName);
  return DocumentApp.openById(copyDoc.getId());
}

function saveDocumentToFolder(doc, sheetName, trainNo) {
  var spreadsheetFile = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
  var parentFolder = spreadsheetFile.getParents().next();

  // Correct folder selection based on actual sheet names
  var reportFolderName = (sheetName === "Report (Visual)") ? "Visual_Inspection_Reports" : "Functional_Inspection_Reports";
  
  var reportFolder = getOrCreateFolder(parentFolder, reportFolderName);
  var trainFolder = getOrCreateFolder(reportFolder, trainNo);
  var file = DriveApp.getFileById(doc.getId());
  file.moveTo(trainFolder);
}


function getOrCreateFolder(parentFolder, folderName) {
  var folders = parentFolder.getFoldersByName(folderName);
  return folders.hasNext() ? folders.next() : parentFolder.createFolder(folderName);
}
