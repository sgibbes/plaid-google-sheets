/** @format */

function onEdit(e) {
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(100)) {
    // Another instance is already running
    return;
  }

  try {
    var editedCell = e.value;
    const editedSheet = e.range.getSheet();
    const range = e.range;

    const editedRow = range.getRow();
    const editedCol = range.getColumn();
    const editedCellLabelCol = editedCol - 1;
    const editedCellLabel = editedSheet.getRange(editedRow, editedCellLabelCol).getValue();

    const CHECKBOX_COL = 8; // Column I
    const CATEGORY_COL = 9; // Column J
    const SUMMARY_START_ROW = 1; // Adjust as needed
    const SUMMARY_END_ROW = 20; // Adjust as needed
    const DATA_START_ROW = 2;
    const CATEGORY_DATA_COL = 4; // Column G
    // the edited cell is the check box AND the value to the left is 'Run Categories'
    // CATEGORIZE
    if (
      range.getColumn() === 7 &&
      range.getRow() === 2 &&
      editedCell === "TRUE" &&
      editedCellLabel === "Run Categories"
    ) {
      SpreadsheetApp.getActiveSpreadsheet().toast("Categorizing Transactions");

      categorizeTransactions();
      editedSheet.getRange(2, 7).setValue("FALSE");
    }

    // SUMMARIZE
    if (
      range.getColumn() === 7 &&
      range.getRow() === 3 &&
      editedCell === "TRUE" &&
      editedCellLabel === "Create Summary Table"
    ) {
      SpreadsheetApp.getActiveSpreadsheet().toast("creating summary table");
      summarizeByCategory();
      editedSheet.getRange(3, 7).setValue("FALSE");
    }

    // FILTER
    if (
      range.getColumn() === CHECKBOX_COL &&
      range.getRow() >= SUMMARY_START_ROW &&
      range.getRow() <= SUMMARY_END_ROW &&
      e.value === "TRUE"
    ) {
      // Gather selected categories
      const categories = [];
      for (let row = SUMMARY_START_ROW; row <= SUMMARY_END_ROW; row++) {
        const checked = sheet.getRange(row, CHECKBOX_COL).getValue();
        const category = sheet.getRange(row, CATEGORY_COL).getValue();
        if (checked && category) {
          categories.push(category);
        }
      }
      Logger.log(categories);
      // Remove existing filter if there is one
      let filter = sheet.getFilter();
      if (filter) filter.remove();

      // Only apply filter if there are checked categories
      if (categories.length > 0) {
        const lastRow = sheet.getLastRow();
        const categoryColumn = CATEGORY_DATA_COL; // Column where categories are located
        const dataRange = sheet.getRange(DATA_START_ROW, categoryColumn, lastRow - DATA_START_ROW + 1);
        const values = dataRange.getValues(); // Get all the values in the category column

        // Loop through the values and hide rows that don't match the categories
        const rowsToHide = [];
        for (let i = 0; i < values.length; i++) {
          const cellValue = values[i][0]; // Get the category value from the cell
          if (!categories.includes(cellValue)) {
            // If it doesn't match any category
            rowsToHide.push(i + DATA_START_ROW); // Add the row number to hide
          }
        }

        // Show all rows first before hiding the non-matching rows
        sheet.showRows(DATA_START_ROW, lastRow - DATA_START_ROW + 1);

        // Hide the rows that don't match the categories
        rowsToHide.forEach((row) => sheet.hideRows(row));

        Logger.log("Rows filtered based on exact categories.");
        return;
      }
    }

    // CLEAR FILTER
    if (range.getColumn() === 15 && editedCell === "TRUE") {
      editedSheet.showRows(1, 100);
      editedSheet.getRange(1, 15).setValue("FALSE");
    }
  } finally {
    lock.releaseLock();
  }
}
