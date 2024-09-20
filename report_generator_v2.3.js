// === CONFIGURATION VARIABLES (EASY ACCESS) ===
var dataStartRow = 11;  // This is where data starts in the sheet
var batchRowSize = 200;  // Maximum number of rows to process per batch
var maxExecutionTime = 10 * 60 * 1000;  // 10 minutes in milliseconds
var triggerInterval = 0.2;  // 12 seconds
var maxImageWidth = 100;  // Maximum image width in pixels
var maxImageHeight = 100;  // Maximum image height in pixels

// === REPORT TRIGGER FUNCTIONS ===

/**
 * Function to trigger Visual report generation for "Visual_Cleaned_Report".
 * This is linked to the button on the Visual sheet.
 */
function generateVisualReport() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Visual_Cleaned_Report");
  if (sheet) {
    processInspectionData(sheet, 'V');
  } else {
    SpreadsheetApp.getUi().alert("Error: 'Visual_Cleaned_Report' sheet not found.");
  }
}

/**
 * Function to trigger Functional report generation for "Functional_Cleaned_Report".
 * This is linked to the button on the Functional sheet.
 */
function generateFunctionalReport() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Functional_Cleaned_Report");
  if (sheet) {
    processInspectionData(sheet, 'F');
  } else {
    SpreadsheetApp.getUi().alert("Error: 'Functional_Cleaned_Report' sheet not found.");
  }
}

// === MAIN SCRIPT FUNCTIONS ===

/**
 * Main function to process data and generate reports.
 * Handles batch processing, time limits, and logging.
 * @param {Sheet} sheet The sheet object to process data from.
 * @param {String} reportPrefix "V" for Visual, "F" for Functional.
 */
function processInspectionData(sheet, reportPrefix) {
  Logger.log("Processing sheet: " + sheet.getName());

  // Record start date and time in cell F3
  var startTime = new Date();
  sheet.getRange("F3").setValue(startTime);
  sheet.getRange("F3").setNumberFormat("dd-MM-yy HH:mm:ss");

  // Set the column mapping dynamically based on the report type
  var columnMapping = {};
  if (reportPrefix === 'V') {
    // Visual Inspection Mapping (SubSubcomponent added, Image URL shifted)
    columnMapping = {
      '{{Inspection ID}}': 2,    // Column B (index 2)
      '{{UserName}}': 5,         // Column E (index 5)
      '{{trainNo}}': 7,          // Column G (index 7)
      '{{Location}}': 8,         // Column H (index 8)
      '{{Car Body}}': 11,        // Column K (index 11)
      '{{Section Name}}': 13,    // Column M (index 13)
      '{{Subsystem Name}}': 15,  // Column O (index 15)
      '{{Serial Number}}': 16,   // Column P (index 16)
      '{{Subcomponent}}': 18,    // Column R (index 18)
      '{{SubSubcomponent}}': 19, // Column S (VSubcomponent_2_Name)
      '{{Condition}}': 20,       // Column T (index 20)
      '{{Defect Type}}': 21,     // Column U (index 21)
      '{{Remarks}}': 22,         // Column V (index 22)
      '{{Image URL}}': 28,       // Column AB (index 28) for Visual
      '{{item No}}': 0           // Custom item number (we'll handle this dynamically)
    };
  } else {
    // Functional Inspection Mapping (original)
    columnMapping = {
      '{{Inspection ID}}': 2,    // Column B (index 2)
      '{{UserName}}': 5,         // Column E (index 5)
      '{{trainNo}}': 7,          // Column G (index 7)
      '{{Location}}': 8,         // Column H (index 8)
      '{{Car Body}}': 11,        // Column K (index 11)
      '{{Section Name}}': 13,    // Column M (index 13)
      '{{Subsystem Name}}': 15,  // Column O (index 15)
      '{{Serial Number}}': 16,   // Column P (index 16)
      '{{Subcomponent}}': 18,    // Column R (index 18)
      '{{Condition}}': 19,       // Column S (index 19)
      '{{Defect Type}}': 20,     // Column T (index 20)
      '{{Remarks}}': 21,         // Column U (index 21)
      '{{Image URL}}': 27,       // Column AA (index 27) for Functional
      '{{item No}}': 0           // Custom item number (we'll handle this dynamically)
    };
  }

  // Fetch the last data row in the sheet
  var lastDataRow = sheet.getLastRow();

  // Calculate the total number of data rows from dataStartRow
  var totalDataRows = lastDataRow - dataStartRow + 1;

  // Calculate rowsToProcess: minimum of batchRowSize and totalDataRows
  var rowsToProcess = Math.min(batchRowSize, totalDataRows);

  // Read the data starting from dataStartRow and fetch rowsToProcess rows
  var dataRange = sheet.getRange(dataStartRow, 1, rowsToProcess, sheet.getLastColumn());
  var data = dataRange.getValues();

  Logger.log("Data fetched: " + JSON.stringify(data));

  if (!data || data.length === 0) {
    SpreadsheetApp.getUi().alert("No data found.");
    return;
  }

  // Fetch Start Item No from F6 and validate it
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

/**
 * Generates the report for the current batch and logs feedback.
 * @param {Array} dataBatch The batch of data being processed (all rows in the range).
 * @param {String} reportPrefix "V" for Visual, "F" for Functional.
 * @param {Number} startItemNo The starting item number for this batch.
 * @param {Object} columnMapping The mapping of columns for placeholders.
 */
function generateReportsFromData(dataBatch, reportPrefix, startItemNo, columnMapping) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Fetch the Train Number from the data
  var trainNo = dataBatch[0][columnMapping['{{trainNo}}'] - 1].toString();

  if (!trainNo || trainNo.trim() === "") {
    SpreadsheetApp.getUi().alert("Error: Missing or invalid Train Number in the data.");
    return;
  }

  // Calculate the end item number based on the data batch size
  var endItemNo = startItemNo + dataBatch.length - 1;

  // Construct the file name using the report prefix, train number, and item numbers
  var fileName = reportPrefix + "-Inspection_Report_for_" + trainNo + "_" + startItemNo + "-" + endItemNo;

  // Fetch Google Doc template ID from cell B4
  var templateId = sheet.getRange("B4").getValue();
  var doc = createDocumentFromTemplate(templateId, fileName);

  // Replace {{trainNo}} in the document header with trainNo and inspection type (Functional or Visual)
  replaceTrainNoPlaceholder(doc, trainNo, reportPrefix);

  // Populate the Google Doc with the current batch of data
  appendTableToDocument(doc.getBody(), dataBatch, startItemNo, reportPrefix, columnMapping);

  doc.saveAndClose();
  saveDocumentToFolder(doc, sheet.getName(), trainNo);

  // Set endItemNo to cell J6
  sheet.getRange("J6").setValue(endItemNo);

  // Add the file name to cell H6
  sheet.getRange("H6").setValue(fileName);

  // Add the document link / URL to cell I6
  var fileUrl = doc.getUrl();
  sheet.getRange("I6").setValue(fileUrl);

  var folderName = reportPrefix === "V" ? "Visual_Inspection_Reports" : "Functional_Inspection_Reports";
  var folderPath = folderName + "/" + trainNo;

  SpreadsheetApp.getUi().alert("Batch processed: " + startItemNo + "-" + endItemNo +
                               "\nFile Name: " + fileName +
                               "\nSaved in: " + folderPath +
                               "\nEnd Item No saved to cell J6.");
}

/**
 * Appends a new table to the Google Doc and fills it with data from the batch.
 * @param {Object} body The body of the Google Doc.
 * @param {Array} dataBatch The batch of data being processed (all rows in the range).
 * @param {Number} startItemNo The starting item number for this batch.
 * @param {String} reportPrefix "V" for Visual, "F" for Functional.
 * @param {Object} columnMapping The mapping of columns for placeholders.
 */
function appendTableToDocument(body, dataBatch, startItemNo, reportPrefix, columnMapping) {
  var tables = body.getTables();  // Get all tables in the document
  var table = tables[0];  // Assume the first table is the one we're appending to

  // Loop through each row in the data batch and append the data to the table
  dataBatch.forEach(function(row, index) {
    var tableRow = table.appendTableRow();
    tableRow.appendTableCell((startItemNo + index).toString());  // No (Item No)
    tableRow.appendTableCell(row[columnMapping['{{Location}}'] - 1].toString());         // Loc
    tableRow.appendTableCell(row[columnMapping['{{Car Body}}'] - 1].toString());         // Car
    tableRow.appendTableCell(row[columnMapping['{{UserName}}'] - 1].toString());         // PIC (User Name)
    tableRow.appendTableCell(row[columnMapping['{{Section Name}}'] - 1].toString());     // Section
    tableRow.appendTableCell(row[columnMapping['{{Subsystem Name}}'] - 1].toString());   // Sub System
    tableRow.appendTableCell(row[columnMapping['{{Serial Number}}'] - 1].toString());    // Serial No
    tableRow.appendTableCell(row[columnMapping['{{Subcomponent}}'] - 1].toString());     // Sub Component
    tableRow.appendTableCell(row[columnMapping['{{Condition}}'] - 1].toString());        // Condition
    tableRow.appendTableCell(row[columnMapping['{{Defect Type}}'] - 1].toString());      // Defect
    tableRow.appendTableCell(row[columnMapping['{{Remarks}}'] - 1].toString());          // Remarks

    // Only insert the new SubSubcomponent field for Visual Reports (reportPrefix 'V')
    if (reportPrefix === 'V') {
      tableRow.appendTableCell(row[columnMapping['{{SubSubcomponent}}'] - 1].toString());  // SubSubcomponent (for Visual reports)
    }

    // Handle image insertion instead of URL
    var imageUrl = row[columnMapping['{{Image URL}}'] - 1].toString();
    var imageCell = tableRow.appendTableCell();

    if (imageUrl && imageUrl.trim() !== "") {
      try {
        var response = UrlFetchApp.fetch(imageUrl);  // Fetch the image
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

  // Remove the first row (the placeholder row) after appending the data
  if (table.getNumRows() > 1) {
    table.removeRow(0);  // Remove the first row of the table (placeholders)
  }
}

/**
 * Replaces the {{trainNo}} placeholder in the document's header with the actual train number
 * and appends "Functional Inspection" or "Visual Inspection" based on the report type.
 * @param {Object} doc The Google Doc object where the replacement will take place.
 * @param {String} trainNo The train number to be inserted.
 * @param {String} reportPrefix "V" for Visual or "F" for Functional.
 */
function replaceTrainNoPlaceholder(doc, trainNo, reportPrefix) {
  var inspectionType = (reportPrefix === 'F') ? 'Functional Inspection' : 'Visual Inspection';
  var fullTrainNo = trainNo + ' (' + inspectionType + ')';  // Combine trainNo with inspection type

  var header = doc.getHeader();  // Get the document's header
  if (header) {
    header.replaceText('{{trainNo}}', fullTrainNo);  // Replace the placeholder with the full train number and inspection type
  }
}

/**
 * Creates a new Google Doc from a template with the specified file name.
 * @param {String} templateId The ID of the Google Doc template.
 * @param {String} fileName The name of the file to be created.
 * @returns {Object} The created Google Doc.
 */
function createDocumentFromTemplate(templateId, fileName) {
  var templateDoc = DriveApp.getFileById(templateId);
  var copyDoc = templateDoc.makeCopy(fileName);
  return DocumentApp.openById(copyDoc.getId());
}

/**
 * Saves the Google Doc in the appropriate folder.
 * @param {Object} doc The Google Doc object to be saved.
 * @param {String} sheetName The name of the sheet (used to determine Visual or Functional).
 * @param {String} trainNo The Train Number for folder naming.
 */
function saveDocumentToFolder(doc, sheetName, trainNo) {
  var spreadsheetFile = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
  var parentFolder = spreadsheetFile.getParents().next();
  var reportFolderName = (sheetName === "Visual_Cleaned_Report") ? "Visual_Inspection_Reports" : "Functional_Inspection_Reports";
  var reportFolder = getOrCreateFolder(parentFolder, reportFolderName);
  var trainFolder = getOrCreateFolder(reportFolder, trainNo);

  var file = DriveApp.getFileById(doc.getId());
  file.moveTo(trainFolder);
}

/**
 * Helper function to get or create a folder in Google Drive.
 * @param {Object} parentFolder The parent folder where the new folder will be created (if needed).
 * @param {String} folderName The name of the folder to be retrieved or created.
 * @returns {Object} The retrieved or created folder.
 */
function getOrCreateFolder(parentFolder, folderName) {
  var folders = parentFolder.getFoldersByName(folderName);
  return folders.hasNext() ? folders.next() : parentFolder.createFolder(folderName);
}
