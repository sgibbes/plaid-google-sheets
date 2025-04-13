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

    const checkboxCol = 8; // Column I
    const categoryCol = 9; // Column J
    const summaryStartRow = 1; // Adjust as needed
    const summaryEndRow = 20; // Adjust as needed
    const dataStartRow = 2;
    const categoryDataCol = 4; // Column G
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
      range.getColumn() === checkboxCol &&
      range.getRow() >= summaryStartRow &&
      range.getRow() <= summaryEndRow &&
      e.value === "TRUE"
    ) {
      // Gather selected categories
      const categories = [];
      for (let row = summaryStartRow; row <= summaryEndRow; row++) {
        const checked = sheet.getRange(row, checkboxCol).getValue();
        const category = sheet.getRange(row, categoryCol).getValue();
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
        const categoryColumn = categoryDataCol; // Column where categories are located
        const dataRange = sheet.getRange(dataStartRow, categoryColumn, lastRow - dataStartRow + 1);
        const values = dataRange.getValues(); // Get all the values in the category column

        // Loop through the values and hide rows that don't match the categories
        const rowsToHide = [];
        for (let i = 0; i < values.length; i++) {
          const cellValue = values[i][0]; // Get the category value from the cell
          if (!categories.includes(cellValue)) {
            // If it doesn't match any category
            rowsToHide.push(i + dataStartRow); // Add the row number to hide
          }
        }

        // Show all rows first before hiding the non-matching rows
        sheet.showRows(dataStartRow, lastRow - dataStartRow + 1);

        // Hide the rows that don't match the categories
        rowsToHide.forEach((row) => sheet.hideRows(row));

        return;
      }
    }

    // CLEAR FILTER
    if (range.getColumn() === 15 && editedCell === "TRUE") {
      // unhides the rows
      editedSheet.showRows(1, 100);
      // set clear filter checkbox back to false
      editedSheet.getRange(1, 15).setValue("FALSE");
      // set all category filters back to false
      const numRows = sheet.getLastRow() + 1;
      const checkboxRange = sheet.getRange(2, checkboxCol, numRows);
      const checkboxValues = checkboxRange.getValues();
      const updatedValues = checkboxValues.map((row) => [row[0] === true ? false : row[0]]);
      checkboxRange.setValues(updatedValues);
    }

    // DOWNLOAD
    const runningFromScriptPage = editedSheet.getName() === "runScript" && editedCell === "TRUE";
    const runningFromCurrentPage =
      range.getColumn() === 7 &&
      range.getRow() === 4 &&
      editedCell === "TRUE" &&
      editedCellLabel === "Download Data Again";
    if (runningFromScriptPage || runningFromCurrentPage) {
      let month = null;
      let year = null;
      if (runningFromScriptPage) {
        editedSheet.getRange(1, 2).setValue("FALSE");
        const monthYr = editedSheet.getRange(2, 2).getValue();
        month = monthYr.split("-")[0];
        year = monthYr.split("-")[1];
      }

      if (runningFromCurrentPage) {
        const sheetName = editedSheet.getName();
        month = sheetName.split("-")[0];
        year = sheetName.split("-")[1];
      }

      if (month && year) {
        getRealTransactions(month, year);
      } else {
        getRealTransactions();
      }

      return;
    }
  } finally {
    lock.releaseLock();
  }
}
