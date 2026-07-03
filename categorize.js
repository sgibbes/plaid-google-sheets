/** @format */

function categorizeTransactions() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();

  const amountCol = data[0].find((x) => x === "Transaction Amount");

  const amntColNum = data[0].indexOf(amountCol);

  const descCol = data[0].find((x) => x === "Transaction Description");
  const descColNum = data[0].indexOf(descCol);

  // Fetch category and values from the other sheet
  const catSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Categories");
  const catData = catSheet.getDataRange().getValues(); // assuming this includes headers
  const catNames = catData[0];
  const categories = {};
  catNames.forEach((cat) => {
    const colNum = catNames.indexOf(cat) + 1;
    const lastRow = catSheet.getLastRow();
    catValues = catSheet
      .getRange(2, colNum, lastRow)
      .getValues()
      ?.flat()
      ?.filter((value) => value !== "");
    categories[cat] = catValues;
  });

  // get column index from transaction amount
  const amountColumnIndex = data[0].indexOf("Transaction Amount");
  const existingCategoryColIndex = data[0].indexOf("Category");
  const existingNotesColIndex = data[0].indexOf("Notes");
  const accountColIndex = data[0].indexOf("Account");
  const catColNum =
    existingCategoryColIndex !== -1 ? existingCategoryColIndex + 1 : accountColIndex !== -1 ? accountColIndex + 2 : amntColNum + 2;

  // Add a 'Category' header if not present
  if (existingCategoryColIndex === -1) {
    sheet.getRange(1, catColNum).setValue("Category");
  }

  if (existingNotesColIndex === -1) {
    sheet.insertColumnAfter(catColNum);
    sheet.getRange(1, catColNum + 1).setValue("Notes");
  }

  // Add a 'Amount' column check (assuming amounts are already in the data)
  if (amountColumnIndex === -1) {
    SpreadsheetApp.getUi().alert("No 'Amount' column found. Please add one.");
    return;
  }

  // Categorize each transaction
  for (let i = 1; i < data.length; i++) {
    const cell = sheet.getRange(i + 1, catColNum);

    if (cell.getValue() === "") {
      const description = data[i][descColNum].trim().toLowerCase();

      // this is the default category. If no matches, its set to this
      let category = "UNCATEGORIZED";
      for (const [key, keywords] of Object.entries(categories)) {
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

      const sortedCats = [...catNames].sort((a, b) => a.toLowerCase().localeCompare(b.toLowerCase()));
      // Write category to the last column, only if value is blank

      // this creates a new validation rule which makes the cell a drop down?
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(
          sortedCats.sort((a, b) => a - b), // sort alphabetically
          true
        )
        .setAllowInvalid(false)
        .build();

      cell.setDataValidation(rule);
    }
  }
}
