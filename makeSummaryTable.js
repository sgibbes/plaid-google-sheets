/** @format */

function summarizeByCategory() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  ensureTransactionLayout_(sheet);

  const data = sheet.getDataRange().getValues(); // includes headers
  const headers = data[0];
  const categoryColIndex = headers.indexOf("Category");
  const outputColStart = categoryColIndex + 6; // leaves room for Notes after Category

  // clear contents
  const rangeToClear = sheet.getRange(1, outputColStart, 2000, 8);
  rangeToClear.clear();

  const amountColIndex = headers.indexOf("Transaction Amount");

  if (categoryColIndex === -1 || amountColIndex === -1) {
    throw new Error('Missing required headers: "Category" or "Transaction Amount"');
  }

  // Fetch Samaris's budget values from the monthly budget sheet.
  // Column A = category, column D = amount, column H = person.
  const budgetSheetName = "Household Expenses May 2026";
  const budgetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(budgetSheetName);

  if (!budgetSheet) {
    throw new Error('Missing sheet: "' + budgetSheetName + '". Check the tab name.');
  }

  const budgetData = budgetSheet.getDataRange().getDisplayValues();
  const expenseSources = [];

  budgetData.slice(1).forEach((row, rowIndex) => {
    const expense = String(row[0] || "").trim();
    const person = String(row[7] || "").trim().toLowerCase();

    if (expense && person === "samaris") {
      const budgetRow = rowIndex + 2;
      expenseSources.push({
        expense: "'" + budgetSheetName + "'!" + budgetSheet.getRange(budgetRow, 1).getA1Notation(),
        amount: "'" + budgetSheetName + "'!" + budgetSheet.getRange(budgetRow, 4).getA1Notation(),
      });
    }
  });

  const totalRowStart = 1; // leave one blank row after data

  // Write headers
  sheet.getRange(totalRowStart, outputColStart).setValue("Filter");

  sheet.getRange(totalRowStart, outputColStart + 1).setValue("Category Totals");
  sheet.getRange(totalRowStart, outputColStart + 2).setValue("Budgeted");
  sheet.getRange(totalRowStart, outputColStart + 3).setValue("Actual");
  sheet.getRange(totalRowStart, outputColStart + 4).setValue("Remaining");

  let currentRow = totalRowStart + 1;
  for (const expenseSource of expenseSources) {
    // Insert checkbox in column H (col 8)
    const checkboxCell = sheet.getRange(currentRow, outputColStart);
    checkboxCell.insertCheckboxes();
    checkboxCell.setValue(false);

    const categoryCell = sheet.getRange(currentRow, outputColStart + 1);
    const budgetCell = sheet.getRange(currentRow, outputColStart + 2);
    const actualCell = sheet.getRange(currentRow, outputColStart + 3);
    const remainingCell = sheet.getRange(currentRow, outputColStart + 4);

    categoryCell.setFormula("=" + expenseSource.expense);

    const categoryRef = categoryCell.getA1Notation();
    const transactionCategoryCol = sheet
      .getRange(1, categoryColIndex + 1)
      .getA1Notation()
      .replace(/\d+/g, "");
    const transactionAmountCol = sheet
      .getRange(1, amountColIndex + 1)
      .getA1Notation()
      .replace(/\d+/g, "");

    budgetCell.setFormula("=" + expenseSource.amount);
    actualCell.setFormula(
      "=IF(LOWER(TRIM(" +
        categoryRef +
        '))="income",' +
        "SUMIF($" +
        transactionCategoryCol +
        ":$" +
        transactionCategoryCol +
        "," +
        categoryRef +
        ",$" +
        transactionAmountCol +
        ":$" +
        transactionAmountCol +
        ")," +
        "-SUMIF($" +
        transactionCategoryCol +
        ":$" +
        transactionCategoryCol +
        "," +
        categoryRef +
        ",$" +
        transactionAmountCol +
        ":$" +
        transactionAmountCol +
        "))",
    );
    remainingCell.setFormula("=" + budgetCell.getA1Notation() + "-" + actualCell.getA1Notation());

    currentRow++;
  }

  // set formats to currency
  const numRows = expenseSources.length;
  if (numRows > 0) {
    const currencyRange = sheet.getRange(totalRowStart + 1, outputColStart + 2, numRows, 3);
    currencyRange.setNumberFormat("$#,##0.00");
    sheet
      .getRange(totalRowStart + 1, outputColStart + 4, numRows, 1)
      .setNumberFormat("$#,##0.00;[Red]-$#,##0.00");
  }
}
