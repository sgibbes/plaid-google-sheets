/** @format */

// if I want to manually edit the categories, this will summarize the totals and not overwrite manual category edits

// filters data by clicked on check boxes next to category totals
// function onEdit(e) {
//   var editedCell = e.value;
//   const editedSheet = e.range.getSheet();
//   const range = e.range;

//   const editedRow = range.getRow();
//   const editedCol = range.getColumn();
//   const editedCellLabelCol = editedCol - 1;
//   const editedCellLabel = editedSheet.getRange(editedRow, editedCellLabelCol).getValue();
//   // the edited cell is the check box AND the value to the left is 'Run Categories'
//   if (
//     range.getColumn() === 7 &&
//     range.getRow() === 2 &&
//     editedCell === "TRUE" &&
//     editedCellLabel === "Run Categories"
//   ) {
//     SpreadsheetApp.getActiveSpreadsheet().toast("categories set");

//     lock.setProperty("running", "true");

//     categorizeTransactions();

//     lock.deleteProperty("running");
//     return;
//   }

//   const CHECKBOX_COL = 8; // Column I
//   const CATEGORY_COL = 9; // Column J
//   const SUMMARY_START_ROW = 1; // Adjust as needed
//   const SUMMARY_END_ROW = 20; // Adjust as needed
//   const DATA_START_ROW = 2;
//   const CATEGORY_DATA_COL = 4; // Column G

//   var sheet = e.source;
//   var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

//   // Only run if a checkbox in the summary section was edited
//   if (
//     range.getColumn() === CHECKBOX_COL &&
//     range.getRow() >= SUMMARY_START_ROW &&
//     range.getRow() <= SUMMARY_END_ROW &&
//     e.value === "TRUE"
//   ) {
//     // Gather selected categories
//     const categories = [];
//     for (let row = SUMMARY_START_ROW; row <= SUMMARY_END_ROW; row++) {
//       const checked = sheet.getRange(row, CHECKBOX_COL).getValue();
//       const category = sheet.getRange(row, CATEGORY_COL).getValue();
//       if (checked && category) {
//         categories.push(category);
//       }
//     }
//     Logger.log(categories);
//     // Remove existing filter if there is one
//     let filter = sheet.getFilter();
//     if (filter) filter.remove();

//     // Only apply filter if there are checked categories
//     if (categories.length > 0) {
//       const lastRow = sheet.getLastRow();
//       const categoryColumn = CATEGORY_DATA_COL; // Column where categories are located
//       const dataRange = sheet.getRange(DATA_START_ROW, categoryColumn, lastRow - DATA_START_ROW + 1);
//       const values = dataRange.getValues(); // Get all the values in the category column

//       // Loop through the values and hide rows that don't match the categories
//       const rowsToHide = [];
//       for (let i = 0; i < values.length; i++) {
//         const cellValue = values[i][0]; // Get the category value from the cell
//         if (!categories.includes(cellValue)) {
//           // If it doesn't match any category
//           rowsToHide.push(i + DATA_START_ROW); // Add the row number to hide
//         }
//       }

//       // Show all rows first before hiding the non-matching rows
//       sheet.showRows(DATA_START_ROW, lastRow - DATA_START_ROW + 1);

//       // Hide the rows that don't match the categories
//       rowsToHide.forEach((row) => sheet.hideRows(row));

//       Logger.log("Rows filtered based on exact categories.");
//       return;
//     }
//   }
// }

// function test_onEdit() {
//   const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("runScript");
//   const range = sheet.getRange("B1");

//   const fakeEvent = {
//     range: range,
//     value: 'TRUE',
//     source: SpreadsheetApp.getActiveSpreadsheet(),
//   };

//   onEdit(fakeEvent);
// }
