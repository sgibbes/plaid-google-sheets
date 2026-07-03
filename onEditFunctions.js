/** @format */

function resetCommandCheckbox_(sheet, row, label, fallbackCol) {
  const rowValues = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  const labelCol = rowValues.indexOf(label) + 1;
  const checkboxCol = labelCol ? labelCol + 1 : fallbackCol;

  sheet.getRange(row, checkboxCol).setValue("FALSE");
}

function onEdit(e) {
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(100)) {
    Logger.log('Another instance is already running')
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

    const headerRow = editedSheet.getRange(1, 1, 1, editedSheet.getLastColumn()).getValues()[0];
    const checkboxCol = headerRow.indexOf("Filter") + 1 || 8;
    const categoryCol = headerRow.indexOf("Category Totals") + 1 || 9;
    const summaryStartRow = 1; // Adjust as needed
    const summaryEndRow = 20; // Adjust as needed
    const dataStartRow = 2;
    const categoryDataCol = headerRow.indexOf("Category") + 1 || 4;
    // the edited cell is the check box AND the value to the left is 'Run Categories'
    // CATEGORIZE
    if (
      range.getRow() === 2 &&
      editedCell === "TRUE" &&
      editedCellLabel === "Run Categories"
    ) {
      SpreadsheetApp.getActiveSpreadsheet().toast("Categorizing Transactions");

      categorizeTransactions();
      resetCommandCheckbox_(editedSheet, 2, "Run Categories", editedCol);
    }

    // SUMMARIZE
    if (
      range.getRow() === 3 &&
      editedCell === "TRUE" &&
      editedCellLabel === "Create Summary Table"
    ) {
      SpreadsheetApp.getActiveSpreadsheet().toast("creating summary table");
      summarizeByCategory();
      resetCommandCheckbox_(editedSheet, 3, "Create Summary Table", editedCol);
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
        const checked = editedSheet.getRange(row, checkboxCol).getValue();
        const category = editedSheet.getRange(row, categoryCol).getValue();
        if (checked && category) {
          categories.push(category);
        }
      }
      Logger.log(categories);
      // Remove existing filter if there is one
      let filter = editedSheet.getFilter();
      if (filter) filter.remove();

      // Only apply filter if there are checked categories
      if (categories.length > 0) {
        const lastRow = editedSheet.getLastRow();
        const categoryColumn = categoryDataCol; // Column where categories are located
        const dataRange = editedSheet.getRange(dataStartRow, categoryColumn, lastRow - dataStartRow + 1);
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
        editedSheet.showRows(dataStartRow, lastRow - dataStartRow + 1);

        // Hide the rows that don't match the categories
        rowsToHide.forEach((row) => editedSheet.hideRows(row));

        return;
      }
    }

    // CLEAR FILTER
    if (editedRow === 1 && editedCell === "TRUE" && editedCellLabel === "Clear Filter") {
      // unhides the rows
      editedSheet.showRows(1, 1000);
      // set clear filter checkbox back to false
      resetCommandCheckbox_(editedSheet, 1, "Clear Filter", editedCol);
      // set all category filters back to false
      const numRows = editedSheet.getLastRow() + 1;
      const checkboxRange = editedSheet.getRange(2, checkboxCol, numRows);
      const checkboxValues = checkboxRange.getValues();
      const updatedValues = checkboxValues.map((row) => [row[0] === true ? false : row[0]]);
      checkboxRange.setValues(updatedValues);
    }

    // DOWNLOAD
    const runningFromScriptPage = editedSheet.getName() === "runScript" && editedCell === "TRUE";
    const runningFromCurrentPage = range.getRow() === 4 && editedCell === "TRUE" && editedCellLabel === "Re-Download Data";
    if (runningFromScriptPage || runningFromCurrentPage) {
      let month = null;
      let year = null;
      if (runningFromScriptPage) {
        editedSheet.getRange(1, 2).setValue("FALSE");
        const monthYr = editedSheet.getRange(2, 2).getValue();
        Logger.log({monthYr})
        month = monthYr.split("-")[0];
        year = monthYr.split("-")[1];
        getRealTransactions(month, year);
      }

      if (runningFromCurrentPage) {
        resetCommandCheckbox_(editedSheet, 4, "Re-Download Data", editedCol);
        const sheetName = editedSheet.getName();
        month = sheetName.split("-")[0];
        year = sheetName.split("-")[1];
        getRealTransactions(month, year, true);
      }

      return;
    }

    // CHART
    if (editedRow === 5 && editedCell === "TRUE" && editedCellLabel === "Create Charts") {
      // remove existing charts
      const charts = editedSheet.getCharts();
      charts.forEach((chart) => editedSheet.removeChart(chart));

      // clear contents
      const remainingSummaryCol = headerRow.indexOf("Remaining") + 1 || 15;
      const rangeToClear = editedSheet.getRange(1, remainingSummaryCol + 2, 2000, 22);
      rangeToClear.clear();

      const categories = ["groceries", "discretionary", "tolls", "gas"];
      let chartNum = 0;
      // find row in summary table that contains 'groceries'
      categories.forEach((x) => {
        chartNum += 1;
        const columnValues = editedSheet.getRange(1, categoryCol, editedSheet.getMaxRows()).getValues();
        const index = columnValues.findIndex((row) => row[0].toString().toLowerCase() === x.toLowerCase()); // find row# in summary table
        if (index != -1) {
          const rowNum = index + 1;
          createChart(rowNum, chartNum);
        }
      });

      resetCommandCheckbox_(editedSheet, 5, "Create Charts", editedCol);
    }
  } finally {
    lock.releaseLock();
  }
}
