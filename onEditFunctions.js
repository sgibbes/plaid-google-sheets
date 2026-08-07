/** @format */

function resetCommandCheckbox_(sheet, row, label, fallbackCol) {
  const rowValues = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  const labelCol = rowValues.indexOf(label) + 1;
  const checkboxCol = labelCol ? labelCol + 1 : fallbackCol;

  sheet.getRange(row, checkboxCol).setValue("FALSE");
}

function onEdit(e) {
  if (shouldUseInstallableEditTrigger_(e)) {
    return;
  }

  handleEdit_(e);
}

function installedOnEdit(e) {
  handleEdit_(e);
}

function installOnEditTrigger() {
  const spreadsheet = SpreadsheetApp.getActive();
  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach((trigger) => {
    if (trigger.getHandlerFunction() === "installedOnEdit") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger("installedOnEdit").forSpreadsheet(spreadsheet).onEdit().create();
  SpreadsheetApp.getActiveSpreadsheet().toast("Installable edit trigger is ready.");
}

function shouldUseInstallableEditTrigger_(e) {
  if (!e || !e.range || e.value !== "TRUE") {
    return false;
  }

  const sheet = e.range.getSheet();
  const row = e.range.getRow();
  const col = e.range.getColumn();
  const label = col > 1 ? sheet.getRange(row, col - 1).getValue() : "";

  return sheet.getName() === "runScript" || label === "Re-Download Data";
}

function handleEdit_(e) {
  if (!e || !e.range || e.value !== "TRUE") {
    return;
  }

  const lock = LockService.getScriptLock();

  if (!lock.tryLock(100)) {
    Logger.log('Another instance is already running')
    // Another instance is already running
    return;
  }

  try {
    const editedCell = e.value;
    const editedSheet = e.range.getSheet();
    const range = e.range;

    const editedRow = range.getRow();
    const editedCol = range.getColumn();
    const editedCellLabel =
      editedCol > 1 ? editedSheet.getRange(editedRow, editedCol - 1).getValue() : "";

    const headerRow = editedSheet.getRange(1, 1, 1, editedSheet.getLastColumn()).getValues()[0];
    const checkboxCol = headerRow.indexOf("Filter") + 1 || 11;
    const categoryCol = headerRow.indexOf("Category Totals") + 1 || 12;
    const summaryStartRow = 1;
    const dataStartRow = 2;
    const categoryDataCol = headerRow.indexOf("Category") + 1 || TRANSACTION_CATEGORY_COL;
    // the edited cell is the check box AND the value to the left is 'Run Categories'
    // CATEGORIZE
    if (
      range.getRow() === 2 &&
      editedCell === "TRUE" &&
      editedCellLabel === "Run Categories"
    ) {
      SpreadsheetApp.getActiveSpreadsheet().toast("Categorizing Transactions");

      categorizeTransactions();
      range.setValue(false);
    }

    // SUMMARIZE
    if (
      range.getRow() === 3 &&
      editedCell === "TRUE" &&
      editedCellLabel === "Create Summary Table"
    ) {
      SpreadsheetApp.getActiveSpreadsheet().toast("creating summary table");
      summarizeByCategory();
      range.setValue(false);
    }

    // FILTER
    const isSummaryFilter = range.getColumn() === checkboxCol;
    const summaryEndRow = isSummaryFilter
      ? editedSheet
          .getRange(editedSheet.getMaxRows(), categoryCol)
          .getNextDataCell(SpreadsheetApp.Direction.UP)
          .getRow()
      : 0;
    if (
      isSummaryFilter &&
      range.getRow() >= summaryStartRow &&
      range.getRow() <= summaryEndRow &&
      e.value === "TRUE"
    ) {
      // Gather selected categories
      const summaryValues = editedSheet
        .getRange(summaryStartRow, checkboxCol, summaryEndRow - summaryStartRow + 1, categoryCol - checkboxCol + 1)
        .getValues();
      const categoryOffset = categoryCol - checkboxCol;
      const categories = new Set(
        summaryValues
        .filter((row) => row[0] === true && row[categoryOffset])
          .map((row) => row[categoryOffset]),
      );
      // Remove existing filter if there is one
      let filter = editedSheet.getFilter();
      if (filter) filter.remove();

      // Only apply filter if there are checked categories
      if (categories.size > 0) {
        const lastRow = editedSheet
          .getRange(editedSheet.getMaxRows(), categoryDataCol)
          .getNextDataCell(SpreadsheetApp.Direction.UP)
          .getRow();
        const categoryColumn = categoryDataCol; // Column where categories are located
        const dataRange = editedSheet.getRange(dataStartRow, categoryColumn, lastRow - dataStartRow + 1);
        const values = dataRange.getValues(); // Get all the values in the category column

        // Group adjacent rows so one hideRows call can hide an entire block.
        const rowGroupsToHide = [];
        let groupStart = null;
        for (let i = 0; i < values.length; i++) {
          const rowNumber = i + dataStartRow;
          if (!categories.has(values[i][0])) {
            groupStart = groupStart === null ? rowNumber : groupStart;
          } else if (groupStart !== null) {
            rowGroupsToHide.push([groupStart, rowNumber - groupStart]);
            groupStart = null;
          }
        }
        if (groupStart !== null) {
          rowGroupsToHide.push([groupStart, lastRow - groupStart + 1]);
        }

        // Show all rows first before hiding the non-matching rows
        editedSheet.showRows(dataStartRow, lastRow - dataStartRow + 1);

        // Hide the rows that don't match the categories
        rowGroupsToHide.forEach(([startRow, count]) => editedSheet.hideRows(startRow, count));

        return;
      }
    }

    // CLEAR FILTER
    if (editedRow === 1 && editedCell === "TRUE" && editedCellLabel === "Clear Filter") {
      // unhides the rows
      editedSheet.showRows(1, editedSheet.getMaxRows());
      // set clear filter checkbox back to false
      range.setValue(false);
      // set all category filters back to false
      const lastSummaryRow = editedSheet
        .getRange(editedSheet.getMaxRows(), categoryCol)
        .getNextDataCell(SpreadsheetApp.Direction.UP)
        .getRow();
      if (lastSummaryRow >= 2) {
        editedSheet.getRange(2, checkboxCol, lastSummaryRow - 1, 1).uncheck();
      }
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
        const transactionSheet = getRealTransactions(month, year);
        if (transactionSheet) {
          const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
          spreadsheet.setActiveSheet(transactionSheet);
          spreadsheet.toast("Categorizing Transactions");
          categorizeTransactions();
          spreadsheet.toast("Creating Summary Table");
          summarizeByCategory();
          spreadsheet.toast("Transactions, categories, and summary are ready.");
        }
      }

      if (runningFromCurrentPage) {
        range.setValue(false);
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
      const actualSummaryCol = headerRow.indexOf("Actual") + 1 || 14;
      const chartDataStartCol = remainingSummaryCol + 2;
      const rangeToClear = editedSheet.getRange(
        1,
        chartDataStartCol,
        editedSheet.getMaxRows(),
        editedSheet.getMaxColumns() - chartDataStartCol + 1,
      );
      rangeToClear.clear();

      const categories = ["groceries", "discretionary", "tolls", "gas"];
      const summaryValues = editedSheet
        .getRange(
          1,
          categoryCol,
          editedSheet.getMaxRows(),
          remainingSummaryCol - categoryCol + 1,
        )
        .getValues();
      const actualOffset = actualSummaryCol - categoryCol;
      const remainingOffset = remainingSummaryCol - categoryCol;
      categories.forEach((category, index) => {
        const chartNum = index + 1;
        const rowIndex = summaryValues.findIndex(
          (row) => String(row[0]).toLowerCase() === category,
        );
        if (rowIndex !== -1) {
          createChart(rowIndex + 1, chartNum, {
            sheet: editedSheet,
            categoryCol,
            actualCol: actualSummaryCol,
            remainingCol: remainingSummaryCol,
            category: summaryValues[rowIndex][0],
            spent: summaryValues[rowIndex][actualOffset],
            remaining: summaryValues[rowIndex][remainingOffset],
          });
        }
      });

      range.setValue(false);
    }
  } finally {
    lock.releaseLock();
  }
}
