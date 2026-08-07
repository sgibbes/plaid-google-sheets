/** @format */

function categorizeTransactions() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getActiveSheet();
  ensureTransactionLayout_(sheet);

  // Only read the transaction columns. Summary tables to the right can make
  // getDataRange() much larger than the transaction data itself.
  const lastSheetRow = sheet.getLastRow();
  const data = sheet
    .getRange(1, 1, Math.max(lastSheetRow, 1), TRANSACTION_CATEGORY_COL)
    .getValues();
  const headers = data[0];
  const descColNum = headers.indexOf("Transaction Description");
  const amountColumnIndex = headers.indexOf("Transaction Amount");

  if (descColNum === -1 || amountColumnIndex === -1) {
    throw new Error(
      'Missing required headers: "Transaction Description" or "Transaction Amount".',
    );
  }

  // Categories are in column A; each category's matching words are in B onward.
  const catSheet = spreadsheet.getSheetByName("Categories");
  if (!catSheet) {
    throw new Error('Missing sheet: "Categories".');
  }

  const catData = catSheet.getDataRange().getDisplayValues();
  const categories = new Map();

  catData.slice(1).forEach((row) => {
    const category = row[0].trim();
    if (!category) {
      return;
    }

    const keywords = row
      .slice(1)
      .map((keyword) => keyword.trim())
      .filter(Boolean)
      .map((keyword) => {
        const exact = keyword.startsWith('"') && keyword.endsWith('"');
        return {
          exact,
          value: (exact ? keyword.slice(1, -1) : keyword).toLowerCase(),
        };
      });

    if (!categories.has(category)) {
      categories.set(category, []);
    }
    categories.get(category).push(...keywords);
  });

  const fallbackCategory = "UNCATEGORIZED";
  const catNames = [...categories.keys()];
  if (!catNames.some((category) => category.toLowerCase() === fallbackCategory.toLowerCase())) {
    catNames.push(fallbackCategory);
  }

  const catColNum = TRANSACTION_CATEGORY_COL;

  const sortedCats = [...catNames].sort((a, b) => a.toLowerCase().localeCompare(b.toLowerCase()));
  const categoryRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(sortedCats, true)
    .setAllowInvalid(false)
    .build();

  // Ignore summary rows below the transactions. A transaction row has either a
  // description or an amount, which also keeps intentionally blank rows intact.
  let lastTransactionIndex = 0;
  for (let i = 1; i < data.length; i++) {
    if (data[i][descColNum] !== "" || data[i][amountColumnIndex] !== "") {
      lastTransactionIndex = i;
    }
  }

  if (lastTransactionIndex === 0) {
    return;
  }

  const categoryRange = sheet.getRange(2, catColNum, lastTransactionIndex, 1);
  const categoryValues = categoryRange.getValues();

  for (let i = 0; i < categoryValues.length; i++) {
    // Keep categories that were assigned or edited manually.
    if (categoryValues[i][0] !== "") {
      continue;
    }

    const description = String(data[i + 1][descColNum] || "").trim().toLowerCase();
    let category = fallbackCategory;

    for (const [categoryName, keywords] of categories.entries()) {
      if (
        keywords.some((keyword) =>
          keyword.exact
            ? description === keyword.value
            : description.includes(keyword.value),
        )
      ) {
        category = categoryName;
        break;
      }
    }

    categoryValues[i][0] = category;
  }

  // These two bulk writes replace several calls per transaction.
  categoryRange.setDataValidation(categoryRule);
  categoryRange.setValues(categoryValues);
}
