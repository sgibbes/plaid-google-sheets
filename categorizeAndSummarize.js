function categorizeTransactions() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();

  const amountCol = data[0].find((x) => x === 'Transaction Amount')

  const amntColNum = data[0].indexOf(amountCol)

  const descCol = data[0].find((x) => x === 'Transaction Description')
  const descColNum = data[0].indexOf(descCol)

  // first convert values to pos/neg:
  for (let i = 1; i < data.length; i++) {
    let amount = data[i][amntColNum]; // Amount in column A
    let transactionType = data[i][3]; // Transaction type in column B

    // Check if the transaction type is "debit"
    if (transactionType?.toLowerCase() === "debit") {
      data[i][4] = -Math.abs(amount); // Convert to negative if debit
    }
  }

  // Update the sheet with the modified data
  sheet.getDataRange().setValues(data);

  // Fetch category and values from the other sheet
  const catSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Categories');
  const catData = catSheet.getDataRange().getValues(); // assuming this includes headers
  const catNames = catData[0]
  const categories = {}
  catNames.forEach((cat) => {
    console.log(catNames.indexOf(cat))
    const colNum = catNames.indexOf(cat) + 1
    const lastRow = catSheet.getLastRow();
    catValues = catSheet.getRange(2, colNum, lastRow).getValues()?.flat()?.filter(value => value !== '');
    categories[cat] = catValues
  })
  Logger.log(categories)
  // get column index from transaction amount
  const amountColumnIndex = data[0].indexOf("Transaction Amount");
  const catColNum = amntColNum + 2
  // Add a 'Category' header if not present
  if (data[0].indexOf("Category") === -1) {
    sheet.getRange(1, amntColNum + 2).setValue("Category");
  }
  // Add a 'Amount' column check (assuming amounts are already in the data)
  if (amountColumnIndex === -1) {
    SpreadsheetApp.getUi().alert("No 'Amount' column found. Please add one.");
    return;
  }
  const categoryColumnIndex = data[0].length; // Last column for Category

  const totals = {}; // To store category totals
  // Categorize each transaction
  for (let i = 1; i < data.length; i++) {
    const description = data[i][descColNum].toLowerCase();

    const amount = parseFloat(data[i][amountColumnIndex]) || 0;

    let category = "UNCATEGORIZED";
    for (const [key, keywords] of Object.entries(categories)) {
      Logger.log(keywords)
      if (keywords.some((keyword) => description.includes(keyword.toString().toLowerCase()))) {
        category = key;
        break;
      }
    }

    // Write category to the last column
    sheet.getRange(i + 1, catColNum).setValue(category);

    // Calculate totals
    if (!totals[category]) totals[category] = 0;
    totals[category] += amount;
  }
  // summarizeByCategory();
}