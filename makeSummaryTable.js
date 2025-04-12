/** @format */

const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
const data = sheet.getDataRange().getValues(); // includes headers

const headers = data[0];
const categoryColIndex = headers.indexOf("Category");
const outputColStart = categoryColIndex + 5; // 3 columns to the right of Category column

function summarizeByCategory() {
  const amountColIndex = headers.indexOf("Transaction Amount");
  Logger.log(data);
  // return
  if (categoryColIndex === -1 || amountColIndex === -1) {
    throw new Error('Missing required headers: "Category" or "Transaction Amount"');
  }

  const categoryTotals = {};

  // Start at row 1 to skip headers
  for (let i = 1; i < data.length; i++) {
    const category = data[i][categoryColIndex];
    const amount = parseFloat(data[i][amountColIndex]);

    if (!isNaN(amount)) {
      if (!categoryTotals[category]) {
        categoryTotals[category] = 0;
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
    budgetMap[category] = budgeted;
  }

  const totalRowStart = 1; // leave one blank row after data

  // Write headers
  sheet.getRange(totalRowStart, outputColStart).setValue("Filter");

  sheet.getRange(totalRowStart, outputColStart + 1).setValue("Category Totals");
  sheet.getRange(totalRowStart, outputColStart + 2).setValue("Total");
  sheet.getRange(totalRowStart, outputColStart + 3).setValue("Budgeted");
  sheet.getRange(totalRowStart, outputColStart + 4).setValue("Budgeted - Actual");

  let currentRow = totalRowStart + 1;
  for (const category in categoryTotals) {
    const total = categoryTotals[category];
    const budgeted = budgetMap[category] || 0;
    const variance = budgeted - total;
    // Insert checkbox in column H (col 8)
    const checkboxCell = sheet.getRange(currentRow, outputColStart);
    checkboxCell.insertCheckboxes();
    checkboxCell.setValue(false);
    sheet.getRange(currentRow, outputColStart + 1).setValue(category);
    sheet.getRange(currentRow, outputColStart + 2).setValue(total);
    sheet.getRange(currentRow, outputColStart + 3).setValue(budgeted);
    const varianceCell = sheet.getRange(currentRow, outputColStart + 4);
    varianceCell.setValue(variance);
    if (variance < 0) {
      varianceCell.setFontColor("red");
    }

    currentRow++;
  }

  // set formats to currency
  const numRows = Object.keys(categoryTotals).length;
  const currencyRange = sheet.getRange(totalRowStart + 1, outputColStart + 1, numRows, 3); // Total, Budgeted, Variance
  currencyRange.setNumberFormat("$#,##0.00");

  // add the unhide box at the end:
  sheet.getRange("N1").setValue("Clear Filter");
  // Set O1 as a checkbox
  sheet.getRange("O1").insertCheckboxes(); // Adds the checkbox
  sheet.getRange("O1").setValue(false); // Optional: uncheck it
}
