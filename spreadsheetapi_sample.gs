
function runExample() {
  // Open Spreadsheet.
  openSpreadsheet();

  // Export data(csv) to sheets, format sheet, generate pivot table.
  let spreadsheetId = exportDataToSheets();

  // Format sheet.
  formatSheet(spreadsheetId);

  // Generate pivot table from csv data.
  generatePivotTable(spreadsheetId);
}


// Open Spreadsheet.
function openSpreadsheet () {
  // Create new Spreadsheet.
  let newSpreadsheet = SpreadsheetApp.create('New Spreadsheet 1');

  // Add new sheet.
  let newSheet = newSpreadsheet.insertSheet();
  // Set it's name to "New Sheet 1".
  newSheet.setName('New Sheet 1');

  // Open active Spreadsheet.
  let openSpreadsheet = SpreadsheetApp.openById(newSpreadsheet.getId());

  // Get existing sheet.
  let existingSheet = openSpreadsheet.getSheetByName('New Sheet 1');
  // Set it's name to "New Sheet 2".
  existingSheet.setName('New Sheet 2');
}


// Export data(csv) to sheets, format sheet, generate pivot table.
function exportDataToSheets () {
  // CSV contents to write.
  const csvContents = `id,class,email,name,math,science,english
1,1A,john@example.com,John J. Coons,90,88,96
2,1A,crystal@example.com,Crystal C. Burnett,32,44,89
3,1B,anthony@example.com,Anthony T. Dudley,72,68,24
4,1B,francisca@example.com,Francisca H. Rapp,89,94,92`;

  // Create new Spreadsheet.
  let newSpreadsheet = SpreadsheetApp.create('CSV import 1');

  // Add new empty sheet.
  let newSheet = newSpreadsheet.insertSheet();
  // Set it's name to "CSV 1".
  newSheet.setName('CSV 1');

  let csvLines = csvContents.split(/[\r\n]/);
  for (var i=0; i<csvLines.length; i++) {
    let csvValues = csvLines[i].split(/,/);

    // Output to sheet line by line.
    let outputRange = newSheet.getRange(i+1, 1, 1, csvValues.length);
    outputRange.setValues([csvValues]);
  }

  return newSpreadsheet.getId();
}


// Format sheet.
function formatSheet(spreadsheetId) {
  // Open existing spreadsheet.
  let spreadsheet = SpreadsheetApp.openById(spreadsheetId);

  // Get data sheet.
  let sheet = spreadsheet.getSheetByName('CSV 1');

  // Clear current format rule.
  sheet.clearFormats();

  // Set column width.
  sheet.setColumnWidth(1,40);
  sheet.setColumnWidth(2,60);
  sheet.setColumnWidth(3,180);
  sheet.setColumnWidth(4,180);

  // Set row height.
  sheet.setRowHeight(1, 30);

  // Set the header line font to bold.
  let headerRange = sheet.getRange(1, 1, 1, 7);
  headerRange.setFontWeight('bold');

  // Freeze header row.
  sheet.setFrozenRows(1);

  // Set border.
  let dataRange = sheet.getRange('A1:G5');
  dataRange.setBorder(true, true, true, true, true, true, 'black', SpreadsheetApp.BorderStyle.SOLID);

  // Create new format rule for score range.
  let scoreRange = sheet.getRange('E2:G5');
  let formatRule1 = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(80)
      .setBackground('#bbccff')
      .setRanges([scoreRange])
      .build();
  let formatRule2 = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(40)
      .setBackground('#ffccbb')
      .setRanges([scoreRange])
      .build();

  // Get current format rules.
  let formatRules = sheet.getConditionalFormatRules();

  // Add new rules.
  formatRules.push(formatRule1);
  formatRules.push(formatRule2);

  // Apply new format rule.
  sheet.setConditionalFormatRules(formatRules);
}


// Generate pivot table from csv data.
function generatePivotTable (spreadsheetId) {
  // Open existing spreadsheet.
  let spreadsheet = SpreadsheetApp.openById(spreadsheetId);

  // Get data sheet.
  let sheet = spreadsheet.getSheetByName('CSV 1');

  // Get source/destination range.
  let sourceRange = sheet.getDataRange();
  let destinationRange = sheet.getRange("c8");

  // Create pivot table.
  let pivotTable = destinationRange.createPivotTable(sourceRange);

  // Set class and name rows.
  pivotTable.addRowGroup(2);
  pivotTable.addRowGroup(3);

  // Set math, science, english as pivot values.
  pivotTable.addPivotValue(5, SpreadsheetApp.PivotTableSummarizeFunction.AVERAGE);
  pivotTable.addPivotValue(6, SpreadsheetApp.PivotTableSummarizeFunction.AVERAGE);
  pivotTable.addPivotValue(7, SpreadsheetApp.PivotTableSummarizeFunction.AVERAGE);
}
