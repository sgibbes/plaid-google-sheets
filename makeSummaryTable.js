/** @format */

function summarizeByCategory() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  ensureTransactionLayout_(sheet);

  const data = sheet.getDataRange().getValues(); // includes headers
  const headers = data[0];
  const categoryColIndex = headers.indexOf("Category");
  const subcategoryColIndex = headers.indexOf("SubCategory");
  const outputColStart = categoryColIndex + 7; // accounts for SubCategory and Notes after Category

  // clear contents
  const rangeToClear = sheet.getRange(1, outputColStart, 2000, 8);
  rangeToClear.clear();

  const amountColIndex = headers.indexOf("Transaction Amount");

  if (categoryColIndex === -1 || subcategoryColIndex === -1 || amountColIndex === -1) {
    throw new Error('Missing required headers: "Category", "SubCategory", or "Transaction Amount"');
  }

  const subcategoriesByCategory = new Map();
  data.slice(1).forEach((row) => {
    const category = String(row[categoryColIndex] || "").trim();
    const subcategory = String(row[subcategoryColIndex] || "").trim();
    if (!category || !subcategory) {
      return;
    }

    const categoryKey = category.toLowerCase();
    const subcategories = subcategoriesByCategory.get(categoryKey) || [];
    if (!subcategories.some((value) => value.toLowerCase() === subcategory.toLowerCase())) {
      subcategories.push(subcategory);
      subcategoriesByCategory.set(categoryKey, subcategories);
    }
  });

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
        category: expense,
        expense: "'" + budgetSheetName + "'!" + budgetSheet.getRange(budgetRow, 1).getA1Notation(),
        amount: "'" + budgetSheetName + "'!" + budgetSheet.getRange(budgetRow, 4).getA1Notation(),
      });
    }
  });

  ["SumZero", "UNCATEGORIZED"].forEach((category) => {
    const existingIndex = expenseSources.findIndex(
      (source) => source.category.toLowerCase() === category.toLowerCase(),
    );
    if (existingIndex !== -1) {
      expenseSources.splice(existingIndex, 1);
    }

    expenseSources.push({ category, expense: null, amount: null });
  });

  const totalRowStart = 1; // leave one blank row after data

  // Write headers
  sheet.getRange(totalRowStart, outputColStart).setValue("Filter");

  sheet.getRange(totalRowStart, outputColStart + 1).setValue("Category Totals");
  sheet.getRange(totalRowStart, outputColStart + 2).setValue("Budgeted");
  sheet.getRange(totalRowStart, outputColStart + 3).setValue("Actual");
  sheet.getRange(totalRowStart, outputColStart + 4).setValue("Subcategory Totals");
  sheet.getRange(totalRowStart, outputColStart + 5).setValue("Remaining");

  // Overall totals and balance calculations live two columns to the right of
  // Remaining (R:S in the standard transaction layout).
  const balanceLabelCol = outputColStart + 7;
  const balanceValueCol = balanceLabelCol + 1;
  sheet.getRange(2, balanceLabelCol, 3, 2).setValues([
    ["total spend within budgeted categories", "=SUM(N:N)"],
    ["budget", "=SUM(M:M)"],
    ["remaining", "=SUM(P:P)"],
  ]);
  sheet.getRange(6, balanceLabelCol).setValue("Samaris Ending Balance");
  sheet.getRange(7, balanceLabelCol, 2, 2).setValues([
    ["magic number", "='Budget May 2026'!B13"],
    ["Magic number (-) ending balance", "=S7-S6"],
  ]);
  sheet.getRange(2, balanceValueCol, 3, 1).setNumberFormat("$#,##0.00;[Red]-$#,##0.00");
  sheet.getRange(6, balanceValueCol, 3, 1).setNumberFormat("$#,##0.00;[Red]-$#,##0.00");

  let currentRow = totalRowStart + 1;
  for (const expenseSource of expenseSources) {
    // Parent category rows can be selected by the summary filter.
    const checkboxCell = sheet.getRange(currentRow, outputColStart);
    checkboxCell.insertCheckboxes();
    checkboxCell.setValue(false);

    const categoryCell = sheet.getRange(currentRow, outputColStart + 1);
    const budgetCell = sheet.getRange(currentRow, outputColStart + 2);
    const actualCell = sheet.getRange(currentRow, outputColStart + 3);
    const remainingCell = sheet.getRange(currentRow, outputColStart + 5);

    if (expenseSource.expense) {
      categoryCell.setFormula("=" + expenseSource.expense);
    } else {
      categoryCell.setValue(expenseSource.category);
    }

    const categoryRef = categoryCell.getA1Notation();
    const transactionCategoryCol = sheet
      .getRange(1, categoryColIndex + 1)
      .getA1Notation()
      .replace(/\d+/g, "");
    const transactionAmountCol = sheet
      .getRange(1, amountColIndex + 1)
      .getA1Notation()
      .replace(/\d+/g, "");
    const transactionSubcategoryCol = sheet
      .getRange(1, subcategoryColIndex + 1)
      .getA1Notation()
      .replace(/\d+/g, "");

    if (expenseSource.amount) {
      budgetCell.setFormula("=" + expenseSource.amount);
    } else {
      budgetCell.setValue(0);
    }
    actualCell.setFormula(
      "=SUMIFS($" +
        transactionAmountCol +
        ":$" +
        transactionAmountCol +
        ",$" +
        transactionCategoryCol +
        ":$" +
        transactionCategoryCol +
        "," +
        categoryRef +
        ")*-1",
    );
    remainingCell.setFormula("=" + budgetCell.getA1Notation() + "-" + actualCell.getA1Notation());

    currentRow++;

    const subcategories = subcategoriesByCategory.get(expenseSource.category.toLowerCase()) || [];
    for (const subcategory of subcategories) {
      const subcategoryFilterCell = sheet.getRange(currentRow, outputColStart);
      const subcategoryCell = sheet.getRange(currentRow, outputColStart + 1);
      const subcategoryActualCell = sheet.getRange(currentRow, outputColStart + 4);
      subcategoryFilterCell.clearContent();
      subcategoryFilterCell.clearDataValidations();
      subcategoryCell.setValue(subcategory);
      subcategoryCell.setHorizontalAlignment("right");

      const subcategoryRef = subcategoryCell.getA1Notation();
      subcategoryActualCell.setFormula(
        "=SUMIFS($" +
          transactionAmountCol +
          ":$" +
          transactionAmountCol +
          ",$" +
          transactionCategoryCol +
          ":$" +
          transactionCategoryCol +
          "," +
          categoryRef +
          ",$" +
          transactionSubcategoryCol +
          ":$" +
          transactionSubcategoryCol +
          "," +
          subcategoryRef +
          ")*-1",
      );

      currentRow++;
    }
  }

  // set formats to currency
  const numRows = currentRow - (totalRowStart + 1);
  if (numRows > 0) {
    const currencyRange = sheet.getRange(totalRowStart + 1, outputColStart + 2, numRows, 4);
    currencyRange.setNumberFormat("$#,##0.00");
    sheet
      .getRange(totalRowStart + 1, outputColStart + 5, numRows, 1)
      .setNumberFormat("$#,##0.00;[Red]-$#,##0.00");
  }

  setSummaryConditionalFormatting_(
    sheet,
    outputColStart + 1,
    outputColStart + 5,
    balanceValueCol,
  );
}

function setSummaryConditionalFormatting_(sheet, categoryCol, remainingCol, balanceValueCol) {
  const remainingRange = sheet.getRange(1, remainingCol, sheet.getMaxRows(), 1);
  const balanceRanges = [sheet.getRange(4, balanceValueCol), sheet.getRange(8, balanceValueCol)];

  // Replace rules created by this function instead of duplicating them each
  // time the summary is regenerated.
  const isSummaryBalanceRule = (rule) => {
    const ranges = rule.getRanges();
    const hasRemainingRange = ranges.some(
      (range) =>
        range.getColumn() === remainingCol &&
        range.getRow() === 1 &&
        range.getNumColumns() === 1,
    );
    const hasBothBalanceRanges = [4, 8].every((row) =>
      ranges.some(
        (range) =>
          range.getColumn() === balanceValueCol &&
          range.getRow() === row &&
          range.getNumRows() === 1 &&
          range.getNumColumns() === 1,
      ),
    );

    return hasRemainingRange || hasBothBalanceRanges;
  };

  const rules = sheet.getConditionalFormatRules().filter((rule) => !isSummaryBalanceRule(rule));
  const remainingColLetter = remainingRange.getA1Notation().replace(/\d+.*$/, "");
  const categoryColLetter = sheet
    .getRange(1, categoryCol)
    .getA1Notation()
    .replace(/\d+/g, "");
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(
        "=AND(" + remainingColLetter + '1>0,$' + categoryColLetter + '1<>"UNCATEGORIZED")',
      )
      .setBackground("#b7e1cd")
      .setFontColor("#0d652d")
      .setRanges([remainingRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(
        "=AND(" + remainingColLetter + '1<0,$' + categoryColLetter + '1<>"UNCATEGORIZED")',
      )
      .setBackground("#f4c7c3")
      .setFontColor("#b31412")
      .setRanges([remainingRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0)
      .setBackground("#b7e1cd")
      .setFontColor("#0d652d")
      .setRanges(balanceRanges)
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0)
      .setBackground("#f4c7c3")
      .setFontColor("#b31412")
      .setRanges(balanceRanges)
      .build(),
  );
  sheet.setConditionalFormatRules(rules);
}
