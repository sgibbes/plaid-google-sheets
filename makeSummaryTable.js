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

  const catSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Categories");
  const catData = catSheet.getDataRange().getValues(); // assuming this includes headers
  const catNames = catData[0];

  if (categoryColIndex === -1 || amountColIndex === -1) {
    throw new Error('Missing required headers: "Category" or "Transaction Amount"');
  }

  // set up category totals with all cat names
  const categoryTotals = {};
  catNames.forEach((cat) => {
    categoryTotals[cat] = 0;
  });

  // Start at row 1 to skip headers
  for (let i = 1; i < data.length; i++) {
    // loop over all transactions
    const category = data[i][categoryColIndex]; // get category of that transaction
    const amount = parseFloat(data[i][amountColIndex]); // get amount of that transaction

    if (!isNaN(amount)) {
      if (!categoryTotals[category]) {
        // if we don't have a category
        categoryTotals[category] = 0; // assign total to 0
      }
      // make totals all positive numbers
      if (category.toLowerCase() != "income") {
        categoryTotals[category] -= amount;
      } else {
        categoryTotals[category] += amount;
      }
    }
  }

  // Fetch budget values from the other sheet
  const budgetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Simplified Samaris Acct Budget");
  const budgetData = budgetSheet.getDataRange().getValues(); // assuming this includes headers
  const budgetMap = {};

  for (let i = 1; i < budgetData.length; i++) {
    const category = budgetData[i][0];
    const budgeted = budgetData[i][1];
    budgetMap[category.toLowerCase()] = budgeted;
  }

  const totalRowStart = 1; // leave one blank row after data

  // Write headers
  sheet.getRange(totalRowStart, outputColStart).setValue("Filter");

  sheet.getRange(totalRowStart, outputColStart + 1).setValue("Category Totals");
  sheet.getRange(totalRowStart, outputColStart + 2).setValue("Budgeted");
  sheet.getRange(totalRowStart, outputColStart + 3).setValue("Actual");
  sheet.getRange(totalRowStart, outputColStart + 4).setValue("Remaining");

  let currentRow = totalRowStart + 1;
  for (const category in categoryTotals) {
    const total = categoryTotals[category];
    const budgeted = budgetMap[category.toLowerCase()] || 0;
    const variance = budgeted - total;
    // Insert checkbox in column H (col 8)
    const checkboxCell = sheet.getRange(currentRow, outputColStart);
    checkboxCell.insertCheckboxes();
    checkboxCell.setValue(false);
    sheet.getRange(currentRow, outputColStart + 1).setValue(category);
    sheet.getRange(currentRow, outputColStart + 3).setValue(total);
    sheet.getRange(currentRow, outputColStart + 2).setValue(budgeted);
    const varianceCell = sheet.getRange(currentRow, outputColStart + 4);
    varianceCell.setValue(variance);

    if (variance < 0) {
      varianceCell.setFontColor("red");
    }

    currentRow++;
  }

  // set formats to currency
  const numRows = Object.keys(categoryTotals).length;
  const currencyRange = sheet.getRange(totalRowStart + 1, outputColStart + 1, numRows, 4); // Start row#, col#, # of rows, # of cols
  currencyRange.setNumberFormat("$#,##0.00");
}
