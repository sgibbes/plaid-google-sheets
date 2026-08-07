/** @format */

function summarizeByCategory() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  ensureTransactionLayout_(sheet);

  const data = sheet
    .getRange(1, 1, Math.max(sheet.getLastRow(), 1), TRANSACTION_NOTES_COL)
    .getValues();
  const headers = data[0];
  const categoryColIndex = headers.indexOf("Category");
  const subcategoryColIndex = headers.indexOf("SubCategory");
  const amountColIndex = headers.indexOf("Transaction Amount");

  if (categoryColIndex === -1 || subcategoryColIndex === -1 || amountColIndex === -1) {
    throw new Error('Missing required headers: "Category", "SubCategory", or "Transaction Amount"');
  }

  const outputColStart = categoryColIndex + 7; // accounts for SubCategory and Notes after Category
  sheet.getRange(1, outputColStart, sheet.getMaxRows(), 8).clear();

  const subcategoriesByCategory = new Map();
  data.slice(1).forEach((row) => {
    const category = String(row[categoryColIndex] || "").trim();
    const subcategory = String(row[subcategoryColIndex] || "").trim();
    if (!category || !subcategory) {
      return;
    }

    const categoryKey = category.toLowerCase();
    const subcategories = subcategoriesByCategory.get(categoryKey) || new Map();
    subcategories.set(subcategory.toLowerCase(), subcategory);
    subcategoriesByCategory.set(categoryKey, subcategories);
  });

  // Fetch Samaris's budget values from the monthly budget sheet.
  // Column A = category, column D = amount, column H = person.
  const budgetSheetName = "Household Expenses May 2026";
  const budgetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(budgetSheetName);

  if (!budgetSheet) {
    throw new Error('Missing sheet: "' + budgetSheetName + '". Check the tab name.');
  }

  const budgetData = budgetSheet
    .getRange(1, 1, Math.max(budgetSheet.getLastRow(), 1), 8)
    .getDisplayValues();
  const expenseSources = [];

  budgetData.slice(1).forEach((row, rowIndex) => {
    const expense = String(row[0] || "").trim();
    const person = String(row[7] || "").trim().toLowerCase();

    if (expense && person === "samaris") {
      const budgetRow = rowIndex + 2;
      expenseSources.push({
        category: expense,
        expense: "'" + budgetSheetName + "'!A" + budgetRow,
        amount: "'" + budgetSheetName + "'!D" + budgetRow,
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

  sheet.getRange(totalRowStart, outputColStart, 1, 6).setValues([
    ["Filter", "Category Totals", "Budgeted", "Actual", "Subcategory Totals", "Remaining"],
  ]);

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

  const transactionCategoryCol = getColumnLetter_(categoryColIndex + 1);
  const transactionAmountCol = getColumnLetter_(amountColIndex + 1);
  const transactionSubcategoryCol = getColumnLetter_(subcategoryColIndex + 1);
  const summaryCategoryCol = getColumnLetter_(outputColStart + 1);
  const summaryBudgetCol = getColumnLetter_(outputColStart + 2);
  const summaryActualCol = getColumnLetter_(outputColStart + 3);
  const summaryRows = [];
  const checkboxValidations = [];
  const categoryAlignments = [];
  const checkboxRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  let currentRow = totalRowStart + 1;

  for (const expenseSource of expenseSources) {
    const categoryRef = summaryCategoryCol + currentRow;
    summaryRows.push([
      false,
      expenseSource.expense ? "=" + expenseSource.expense : expenseSource.category,
      expenseSource.amount ? "=" + expenseSource.amount : 0,
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
      "",
      "=" + summaryBudgetCol + currentRow + "-" + summaryActualCol + currentRow,
    ]);
    checkboxValidations.push([checkboxRule]);
    categoryAlignments.push([null]);

    currentRow++;

    const subcategoryMap = subcategoriesByCategory.get(expenseSource.category.toLowerCase());
    const subcategories = subcategoryMap ? [...subcategoryMap.values()] : [];
    for (const subcategory of subcategories) {
      const subcategoryRef = summaryCategoryCol + currentRow;
      summaryRows.push([
        "",
        subcategory,
        "",
        "",
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
        "",
      ]);
      checkboxValidations.push([null]);
      categoryAlignments.push(["right"]);

      currentRow++;
    }
  }

  const numRows = summaryRows.length;
  if (numRows > 0) {
    const summaryRange = sheet.getRange(totalRowStart + 1, outputColStart, numRows, 6);
    summaryRange.setValues(summaryRows);
    sheet
      .getRange(totalRowStart + 1, outputColStart, numRows, 1)
      .setDataValidations(checkboxValidations);
    sheet
      .getRange(totalRowStart + 1, outputColStart + 1, numRows, 1)
      .setHorizontalAlignments(categoryAlignments);

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

function getColumnLetter_(column) {
  let result = "";
  for (let value = column; value > 0; value = Math.floor((value - 1) / 26)) {
    result = String.fromCharCode(((value - 1) % 26) + 65) + result;
  }
  return result;
}
