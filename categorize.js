/** @format */

function categorizeTransactions() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  ensureTransactionLayout_(sheet);
  const data = sheet.getDataRange().getValues();

  const descCol = data[0].find((x) => x === "Transaction Description");
  const descColNum = data[0].indexOf(descCol);

  // Categories are in column A; each category's matching words are in B onward.
  const catSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Categories");
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
      .filter(Boolean);

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

  // get column index from transaction amount
  const amountColumnIndex = data[0].indexOf("Transaction Amount");
  const catColNum = TRANSACTION_CATEGORY_COL;

  // Add a 'Amount' column check (assuming amounts are already in the data)
  if (amountColumnIndex === -1) {
    SpreadsheetApp.getUi().alert("No 'Amount' column found. Please add one.");
    return;
  }

  const sortedCats = [...catNames].sort((a, b) => a.toLowerCase().localeCompare(b.toLowerCase()));
  const categoryRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(sortedCats, true)
    .setAllowInvalid(false)
    .build();

  // Categorize each transaction
  for (let i = 1; i < data.length; i++) {
    const cell = sheet.getRange(i + 1, catColNum);
    cell.setDataValidation(categoryRule);

    if (cell.getValue() === "") {
      const description = String(data[i][descColNum] || "").trim().toLowerCase();

      // this is the default category. If no matches, its set to this
      let category = fallbackCategory;
      for (const [key, keywords] of categories.entries()) {
        // if any of the keywords are in the description, mark that the category
        if (
          keywords.some((keyword) => {
            // keywords wrapped in double quotes need an exact match
            const exactMatchWord = keyword.startsWith('"') && keyword.endsWith('"');
            let matchCat = false;
            if (exactMatchWord) {
              const exact = keyword.slice(1, -1).toLowerCase();

              matchCat = description === exact;
            } else {
              matchCat = description.includes(keyword.toLowerCase());
            }

            return matchCat;
          })
        ) {
          category = key;
          break;
        }
      }
      cell.setValue(category);
    }
  }
}
