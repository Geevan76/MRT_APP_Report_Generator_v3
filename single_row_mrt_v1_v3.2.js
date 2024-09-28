// === CONFIGURATION VARIABLES ===
var dataStartRow = 11;       // Starting row for data in the sheet
var batchRowSize = 200;      // Maximum number of rows to process per batch
var maxImageWidth = 100;     // Maximum image width in pixels
var maxImageHeight = 100;    // Maximum image height in pixels

// Column Mapping (adjust based on your sheet's structure)
var columnMapping = {
  '{{Inspection ID}}': 2,          // Column B (index 2)
  '{{Image VInspectionID}}': 24,   // Column X (index 24)
  '{{UserName}}': 5,               // Column E (index 5)
  '{{trainNo}}': 7,                // Column G (index 7)
  '{{Location}}': 8,               // Column H (index 8)
  '{{Car Body}}': 11,              // Column K (index 11)
  '{{Section Name}}': 13,          // Column M (index 13)
  '{{Subsystem Name}}': 15,        // Column O (index 15)
  '{{Serial Number}}': 16,         // Column P (index 16)
  '{{Subcomponent}}': 18,          // Column R (index 18)
  '{{Condition}}': 19,             // Column S (index 19)
  '{{Defect Type}}': 20,           // Column T (index 20)
  '{{Remarks}}': 21,               // Column U (index 21)
  '{{Image URL}}': 27,             // Column AA (index 27)
  '{{ImgDescription}}': 26,        // Column Z (index 26)
  '{{item No}}': 0                 // Custom item number (handled dynamically)
};

// === REPORT TRIGGER FUNCTIONS ===

/**
 * Function to trigger Visual report generation for "Report (Visual)".
 */
function generateVisualReport() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Report (Visual)");
  if (sheet) {
    processInspectionData(sheet, 'V');
  } else {
    SpreadsheetApp.getUi().alert("Error: 'Report (Visual)' sheet not found.");
  }
}

/**
 * Function to trigger Functional report generation for "Report (Functional)".
 */
function generateFunctionalReport() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Report (Functional)");
  if (sheet) {
    processInspectionData(sheet, 'F');
  } else {
    SpreadsheetApp.getUi().alert("Error: 'Report (Functional)' sheet not found.");
  }
}

// === MAIN SCRIPT FUNCTIONS ===

/**
 * Processes inspection data and generates the report.
 * @param {Sheet} sheet - The sheet containing the data.
 * @param {String} reportPrefix - "V" for Visual, "F" for Functional.
 */
function processInspectionData(sheet, reportPrefix) {
  Logger.log("Processing sheet: " + sheet.getName());

  // Record start time in cell F3
  var startTime = new Date();
  sheet.getRange("F3").setValue(startTime);
  sheet.getRange("F3").setNumberFormat("dd-MM-yy HH:mm:ss");

  // Determine the range of data to process
  var lastDataRow = sheet.getLastRow();
  var totalDataRows = lastDataRow - dataStartRow + 1;
  var rowsToProcess = Math.min(batchRowSize, totalDataRows);

  if (rowsToProcess <= 0) {
    SpreadsheetApp.getUi().alert("No data rows to process.");
    return;
  }

  // Fetch data
  var dataRange = sheet.getRange(dataStartRow, 1, rowsToProcess, sheet.getLastColumn());
  var data = dataRange.getValues();
  Logger.log("Data fetched: " + JSON.stringify(data));

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

  // Group data by Inspection ID, maintaining order
  var groupedResult = groupDataByInspectionID(data);
  var inspections = groupedResult.inspections;
  var inspectionOrder = groupedResult.inspectionOrder;

  // Generate the report and get endItemNo
  var endItemNo = generateConsolidatedReport(inspections, inspectionOrder, reportPrefix, startItemNo);

  // Record end time in cell G3
  var endTime = new Date();
  sheet.getRange("G3").setValue(endTime);
  sheet.getRange("G3").setNumberFormat("dd-MM-yy HH:mm:ss");

  // Calculate duration and write to H3
  var durationMillis = endTime.getTime() - startTime.getTime();
  var diffHours = Math.floor(durationMillis / (1000 * 60 * 60));
  var diffMinutes = Math.floor((durationMillis % (1000 * 60 * 60)) / (1000 * 60));
  var diffSeconds = Math.floor((durationMillis % (1000 * 60)) / 1000);
  var durationFormatted = diffHours + "h " + diffMinutes + "m " + diffSeconds + "s";
  sheet.getRange("H3").setValue(durationFormatted);

  // Update Start Item No for next batch in F6
  sheet.getRange("F6").setValue(endItemNo + 1);

  // Update endItemNo in J6
  sheet.getRange("J6").setValue(endItemNo);

  SpreadsheetApp.getUi().alert("Report generation complete.\nEnd Item No: " + endItemNo);
}

/**
 * Groups data by Inspection ID, maintaining the order of first appearance.
 * @param {Array} data - The data to group.
 * @returns {Object} - Contains grouped data and inspection order.
 */
function groupDataByInspectionID(data) {
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
      // This is an inspection-only row
      inspections[keyVInspectionID].inspectionRow = row;
    }
  }

  return { inspections: inspections, inspectionOrder: inspectionOrder };
}

/**
 * Generates the consolidated report in Google Docs.
 * @param {Object} inspections - Data grouped by Inspection ID.
 * @param {Array} inspectionOrder - Ordered array of Inspection IDs.
 * @param {String} reportPrefix - "V" or "F".
 * @param {Number} startItemNo - Starting item number.
 * @returns {Number} - End item number after processing.
 */
function generateConsolidatedReport(inspections, inspectionOrder, reportPrefix, startItemNo) {
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
      // Has images, merge inspection data with each imageRow
      for (var j = 0; j < imageRows.length; j++) {
        var imageRow = imageRows[j];
        if (inspectionRow) {
          mergeInspectionDataIntoImageRow(inspectionRow, imageRow);
        }
        rowsToAdd.push({ row: imageRow, itemNo: currentItemNo });
        currentItemNo++;
      }
    } else {
      // Edge case: no inspectionRow but has imageRows
      for (var j = 0; j < imageRows.length; j++) {
        var imageRow = imageRows[j];
        rowsToAdd.push({ row: imageRow, itemNo: currentItemNo });
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
    appendTableRowToDocument(doc.getBody(), item.row, item.itemNo);
  }

  // Remove the first row (the placeholder row) after appending the data
  var table = doc.getBody().getTables()[0];
  if (table && table.getNumRows() > 1) {
    table.removeRow(0);  // Remove the first row of the table (placeholders)
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

/**
 * Merges inspection data into the image row without overwriting the image description or URL.
 * @param {Array} inspectionRow - The row with inspection data.
 * @param {Array} imageRow - The row with image data.
 */
function mergeInspectionDataIntoImageRow(inspectionRow, imageRow) {
  Logger.log("Merging inspectionRow: " + JSON.stringify(inspectionRow));
  Logger.log("Merging imageRow before merge: " + JSON.stringify(imageRow));
  
  var numColumns = inspectionRow.length;

  for (var col = 0; col < numColumns; col++) {
    // Keep the image description and image URL intact
    if (col !== columnMapping['{{ImgDescription}}'] - 1 && col !== columnMapping['{{Image URL}}'] - 1) {
      imageRow[col] = inspectionRow[col];
    }
  }
  
  Logger.log("Merging imageRow after merge: " + JSON.stringify(imageRow));
}

/**
 * Appends a single row of data to the document's table.
 * @param {Object} body - The body of the Google Doc.
 * @param {Array} row - The data row to append.
 * @param {Number} itemNo - The item number.
 */
function appendTableRowToDocument(body, row, itemNo) {
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

  // Add the image description before the image
  var imgDescription = (row[columnMapping['{{ImgDescription}}'] - 1] || "").toString();
  Logger.log("Appending Image Description: " + imgDescription);  // Debugging log
  newRow.getCell(11).setText(imgDescription);  // Image Description

  // Handle image insertion
  var imageCell = newRow.getCell(numColumns - 1); // Last cell for Image
  imageCell.clear(); // Clear existing content

  var imageUrl = row[columnMapping['{{Image URL}}'] - 1];
  imageUrl = imageUrl ? imageUrl.toString().trim() : "";

  if (imageUrl !== "") {
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

/**
 * Replaces the {{trainNo}} placeholder in the document's header.
 * @param {Object} doc - The Google Doc object.
 * @param {String} trainNo - The train number.
 * @param {String} reportPrefix - "V" or "F".
 */
function replaceTrainNoPlaceholder(doc, trainNo, reportPrefix) {
  var inspectionType = (reportPrefix === 'F') ? 'Functional Inspection' : 'Visual Inspection';
  var fullTrainNo = trainNo + ' (' + inspectionType + ')';

  var header = doc.getHeader();
  if (header) {
    header.replaceText('{{trainNo}}', fullTrainNo);
  }
}

/**
 * Creates a new Google Doc from a template.
 * @param {String} templateId - The ID of the Google Doc template.
 * @param {String} fileName - The name for the new document.
 * @returns {Object} - The new Google Doc.
 */
function createDocumentFromTemplate(templateId, fileName) {
  var templateDoc = DriveApp.getFileById(templateId);
  var copyDoc = templateDoc.makeCopy(fileName);
  return DocumentApp.openById(copyDoc.getId());
}

/**
 * Saves the document to the appropriate folder in Drive.
 * @param {Object} doc - The Google Doc object.
 * @param {String} sheetName - Name of the sheet ("Report (Visual)" or "Report (Functional)").
 * @param {String} trainNo - The train number for folder naming.
 */
function saveDocumentToFolder(doc, sheetName, trainNo) {
  var spreadsheetFile = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
  var parentFolder = spreadsheetFile.getParents().next();
  var reportFolderName = (sheetName === "Report (Visual)") ? "Visual_Inspection_Reports" : "Functional_Inspection_Reports";
  var reportFolder = getOrCreateFolder(parentFolder, reportFolderName);
  var trainFolder = getOrCreateFolder(reportFolder, trainNo);

  var file = DriveApp.getFileById(doc.getId());
  file.moveTo(trainFolder);
}

/**
 * Retrieves an existing folder by name or creates it if it doesn't exist.
 * @param {Object} parentFolder - The parent folder.
 * @param {String} folderName - The name of the folder to retrieve or create.
 * @returns {Object} - The retrieved or newly created folder.
 */
function getOrCreateFolder(parentFolder, folderName) {
  var folders = parentFolder.getFoldersByName(folderName);
  return folders.hasNext() ? folders.next() : parentFolder.createFolder(folderName);
}
