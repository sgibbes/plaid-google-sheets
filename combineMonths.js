function addToSummarySheet() {
  // // Prompt the user for input
  // const ui = SpreadsheetApp.getUi();
  // const promptResult = ui.prompt("Please enter a comma-separated list:");
  // Logger.log(promptResult)

  // Fetch budget values from the other sheet
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const budgetSheet = spreadsheet.getSheetByName("Simplified Samaris Acct Budget");
  const budgetData = budgetSheet.getDataRange().getValues(); // assuming this includes headers
  const budgetMap = {};

  for (let i = 1; i < budgetData.length; i++) {
    const category = budgetData[i][0];
    const budgeted = budgetData[i][1];
    budgetMap[category] = budgeted;
  }
  const sheetNames = ["Jan 2025", "Feb 2025", "March 2025"];
  // const sheetNames =promptResult.getResponseText().split(',');
  // Logger.log(promptResult.getResponseText())
  Logger.log(sheetNames)
  const summaryData = {};
  for (const sheet in sheetNames) {
    const sheetToGet = spreadsheet.getSheetByName(sheetNames[sheet]);
    const lastRow = sheetToGet.getLastRow();

    const range = sheetToGet.getRange(1, 11, lastRow, 2);
    const values = range.getValues();
    const nonEmptyRows = values.filter((row) => row[0] !== "" || row[1] !== "");
    summaryData[sheetNames[sheet]] = nonEmptyRows;

    const summaryMap = {};
    for (let i = 1; i < nonEmptyRows.length; i++) {
      const category = nonEmptyRows[i][0];
      const amount = nonEmptyRows[i][1];
      summaryMap[category] = amount;
    }
    summaryData[sheetNames[sheet]] = summaryMap;
  }
  //   {March 2025=[[Income, 6027.86], [Discovery Plus, 19.98], [Groceries, 1005.6199999999999], [Sam Moving Money, 640.1499999999999], [Uncategorized, 969.69], [Gas, 349.56999999999994], [Tolls, 46.0], [Washington Gas, 92.0], [Verizon Cable/Internet, 180.7], [Dominion Electric, 114.13], [Netflix, 17.99], [Mortgage, 2677.73], [Sportrock, 105.0]], Jan 2025=[[I
  // get the summary sheet and add a the first column with all the data from the budgets sheet
  const summarySheet = spreadsheet.getSheetByName("All Months");
  const categories = Object.keys(budgetMap);
  const output = [
    ["Category", "Budgeted", ...sheetNames],
    ...categories.map((category) => [
      category,
      budgetMap[category],
      ...sheetNames.map((monthName) => summaryData[monthName][category] ?? ""),
    ]),
  ];
  summarySheet.getRange(1, 1, output.length, output[0].length).setValues(output);
}
