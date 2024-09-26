function updateVisualInspectionData() {
  var sourceSpreadsheetId = '1u_ZCbLY0F2Do85q7A3uVbVq1PZe5EJPmJQf3bQ_qTFs'; // Source Spreadsheet ID
  var sourceRange = 'VSummary!A:AA'; // Source Sheet and Range for Visual Data
  var destinationSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("New Visual Version"); // Destination Sheet for Visual Data
  
  // Clear the existing data in the destination sheet
  destinationSheet.clear(); 
  
  // Get the source data including headers
  var sourceSpreadsheet = SpreadsheetApp.openById(sourceSpreadsheetId);
  var sourceData = sourceSpreadsheet.getRange(sourceRange).getValues();
  
  // Paste the imported data into the destination sheet
  destinationSheet.getRange(1, 1, sourceData.length, sourceData[0].length).setValues(sourceData);
  
  // Apply background color to the header row (first row)
  destinationSheet.getRange(1, 1, 1, sourceData[0].length).setBackground("#9df8ff"); // Set the background color for headers

  SpreadsheetApp.getUi().alert("Visual Inspection Data Updated Successfully!");
}



function updateFunctionalInspectionData() {
  var sourceSpreadsheetId = '1u_ZCbLY0F2Do85q7A3uVbVq1PZe5EJPmJQf3bQ_qTFs'; // Source Spreadsheet ID
  var sourceRange = 'FSummary!A:AA'; // Source Sheet and Range for Functional Data
  var destinationSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("New Functional Version"); // Destination Sheet for Functional Data
  
  // Clear the existing data in the destination sheet
  destinationSheet.clear(); 
  
  // Get the source data including headers
  var sourceSpreadsheet = SpreadsheetApp.openById(sourceSpreadsheetId);
  var sourceData = sourceSpreadsheet.getRange(sourceRange).getValues();
  
  // Paste the imported data into the destination sheet
  destinationSheet.getRange(1, 1, sourceData.length, sourceData[0].length).setValues(sourceData);
  
  // Apply background color to the header row (first row)
  destinationSheet.getRange(1, 1, 1, sourceData[0].length).setBackground("#ffb9d8"); // Set the background color for headers

  SpreadsheetApp.getUi().alert("Functional Inspection Data Updated Successfully!");
}


function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Create a new custom menu called "Data Import"
  ui.createMenu('Data Import')
    .addItem('Update Visual Inspection Data', 'updateVisualInspectionData') // Adds a button for Visual Data Update
    .addItem('Update Functional Inspection Data', 'updateFunctionalInspectionData') // Adds a button for Functional Data Update
    .addToUi();
}
