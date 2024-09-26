// === CONFIGURATION VARIABLES ===
var dataStartRow = 11;       // Starting row for data in the sheet
var batchRowSize = 200;      // Maximum number of rows to process per batch
var maxImageWidth = 100;     // Maximum image width in pixels
var maxImageHeight = 100;    // Maximum image height in pixels

// Column Mapping (adjust based on your sheet's structure)
var columnMapping = {
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
  '{{Image URL}}': 27,       // Column AA (index 27)
  '{{ImgDescription}}': 26,  // Column Z (index 26)
  '{{item No}}': 0           // Custom item number (handled dynamically)
};

// === REPORT TRIGGER FUNCTIONS ===

/**
 * Function to trigger Visual report generation for "Visual_Cleaned_Report".
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

  // Group data by Inspection ID
  var groupedData = groupByInspectionID(data);

  // Log grouped data to verify correctness
  Logger.log("Grouped Data: " + JSON.stringify(groupedData));

  // Generate the report and get endItemNo
  var endItemNo = generateConsolidatedReport(groupedData, reportPrefix, startItemNo);

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
 * Groups data by Inspection ID.
 * @param {Array} data - The data to group.
 * @returns {Object} - Grouped data.
 */
function groupByInspectionID(data) {
  var groupedData = {};
  data.forEach(function(row) {
    var inspectionID = row[columnMapping['{{Inspection ID}}'] - 1];
    if (!groupedData[inspectionID]) {
      groupedData[inspectionID] = [];
    }
    groupedData[inspectionID].push(row);
  });
  return groupedData;
}

/**
 * Generates the consolidated report in Google Docs.
 * @param {Object} groupedData - Data grouped by Inspection ID.
 * @param {String} reportPrefix - "V" or "F".
 * @param {Number} startItemNo - Starting item number.
 * @returns {Number} - End item number after processing.
 */
function generateConsolidatedReport(groupedData, reportPrefix, startItemNo) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Get Train Number
  var firstDataRow = sheet.getRange(dataStartRow, columnMapping['{{trainNo}}']).getValue();
  var trainNo = firstDataRow ? firstDataRow.toString().trim() : "Unknown_Train_No";

  var rowsToAdd = [];

  // Prepare data rows
  for (var inspectionID in groupedData) {
    var rows = groupedData[inspectionID];
    var inspectionRow = null;
    var imageRows = [];

    rows.forEach(function(row) {
      var imageUrl = row[columnMapping['{{Image URL}}'] - 1];
      imageUrl = imageUrl ? imageUrl.toString().trim() : "";
      if (imageUrl !== "") {
        imageRows.push(row);
      } else {
        inspectionRow = row;
      }
    });

    if (imageRows.length === 0 && inspectionRow) {
      rowsToAdd.push({ row: inspectionRow, startItemNo: startItemNo });
      startItemNo++;
    } else if (imageRows.length > 0 && inspectionRow) {
      imageRows.forEach(function(imageRow) {
        mergeInspectionDataIntoImageRow(inspectionRow, imageRow);
        rowsToAdd.push({ row: imageRow, startItemNo: startItemNo });
        startItemNo++;
      });
    }
  }

  // Check if there's data to process
  if (rowsToAdd.length === 0) {
    SpreadsheetApp.getUi().alert("No data rows to process for report generation.");
    return startItemNo - 1;
  }

  // Calculate endItemNo
  var endItemNo = startItemNo - 1;

  // Construct the filename with both startItemNo and endItemNo
  var fileName = reportPrefix + "-Inspection_Report_for_" + trainNo + "_" + rowsToAdd[0].startItemNo + "_" + endItemNo;

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
  rowsToAdd.forEach(function(item) {
    appendTableRowToDocument(doc.getBody(), item.row, item.startItemNo);
  });

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
  Logger.log("Batch processed. Start Item No: " + rowsToAdd[0].startItemNo + ", End Item No: " + endItemNo);

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
  
  for (var col = 0; col < columnMapping['{{Image URL}}'] - 1; col++) {
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
 * @param {Number
