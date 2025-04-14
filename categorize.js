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
    console.log(catNames.indexOf(cat));
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
  const catColNum = amntColNum + 2;
  // Add a 'Category' header if not present
  if (data[0].indexOf("Category") === -1) {
    sheet.getRange(1, amntColNum + 2).setValue("Category");
  }
  // Add a 'Amount' column check (assuming amounts are already in the data)
  if (amountColumnIndex === -1) {
    SpreadsheetApp.getUi().alert("No 'Amount' column found. Please add one.");
    return;
  }

  // Categorize each transaction
  for (let i = 1; i < data.length; i++) {
    const description = data[i][descColNum].toLowerCase();

    const amount = parseFloat(data[i][amountColumnIndex]) || 0;

    let category = "UNCATEGORIZED";
    for (const [key, keywords] of Object.entries(categories)) {
      Logger.log(keywords);
      if (keywords.some((keyword) => description.includes(keyword.toString().toLowerCase()))) {
        category = key;
        break;
      }
    }

    // Write category to the last column, only is value is blank
    const cell = sheet.getRange(i + 1, catColNum);
    if (cell.getValue() === "") {
      cell.setValue(category);
    }
  }
}
