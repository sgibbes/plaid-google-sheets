/** @format */

function getRealTransactionsWrapper() {
  getRealTransactions(); // optionally pass fixed args here if you want
}

function getRealTransactions(userMonth = null, userYear = null) {
  const accessToken = PropertiesService.getScriptProperties().getProperty("PLAID_ACCESS_TOKEN");
  if (!accessToken) {
    throw new Error("No access token found. Run launchPlaidLink() first.");
  }

  // const today = new Date();
  // const startDate = new Date(today);
  // Logger.log(startDate.getMonth())
  // Logger.log(startDate) //Fri Apr 11 12:36:09 GMT-04:00 2025
  // const daysToGoBack = 3
  // startDate.setDate(today.getDate() - daysToGoBack);

  // // if I want to get the current month of data:
  const formatMonth = (datesMonth) => {
    const m = datesMonth.getMonth() + 1;
    return m < 10 ? `0${m}` : `${m}`;
  };
  // const startYear = startDate.getFullYear()
  // const currentMonthStartDate = `${startYear}-${startMonth()}-01`

  const getEndDayNum = (year, month) => {
    const d = new Date(year, month, 0);
    return d.getDate();
  };

  // Logger.log(getEndDayNum())
  // const currentMonthEndDate = `${startYear}-${startMonth}-${getEndDayNum(startYear, startDate.getMonth())}`
  // const startDateFormatted = startDate.toISOString().slice(0, 10) // like '2025-04-11'
  // const endDateFormatted = today.toISOString().slice(0, 10)
  // Logger.log(currentMonthStartDate)
  // Logger.log(endDateFormatted)

  let month = userMonth ? userMonth : "01";
  let year = userYear ? userYear : "2025";

  // set up default date range
  let startDate = `${month}-01-${year}`;
  let endDate = `${month}-02-${year}`;

  // if no user defined month/year, get all data available for this month
  if (!userMonth && !userYear) {
    const today = new Date();
    endDate = today.toISOString().slice(0, 10); // like '2025-04-11'
    const thisMonth = formatMonth(today);
    const thisYear = today.getFullYear();
    year = thisYear;
    month = thisMonth;
    startDate = `${thisYear}-${thisMonth}-01`;
  }

  // if user define month/year, get month's data. this will error if month end date is in the future
  if (userMonth && userYear) {
    startDate = `${year}-${month}-01`;
    endDate = `${year}-${month}-${getEndDayNum(year, month)}`;
  }
  // startDate = '2025-01-12'
  // endDate = '2025-01-13'
  // SpreadsheetApp.getActiveSpreadsheet().toast(endDate)
  const response = UrlFetchApp.fetch("https://production.plaid.com/transactions/get", {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({
      client_id: "67f7d02c204d3500278eb772",
      secret: "22426845ed29506ee59faef7d0bbee",
      access_token: accessToken,
      start_date: startDate,
      end_date: endDate,
    }),
  });
  const transactionsResponse = JSON.parse(response.getContentText());

  Logger.log("Total transactions: " + transactionsResponse.total_transactions);
  Logger.log("Transactions returned: " + transactionsResponse.transactions.length);
  Logger.log("Earliest date returned: " + (transactionsResponse.transactions.at(-1)?.date || "No transactions"));
  // return
  let transactions = JSON.parse(response.getContentText()).transactions;
  Logger.log(startDate);
  Logger.log(transactions);

  const total_transactions = JSON.parse(response.getContentText()).total_transactions;
  // Manipulate the offset parameter to paginate
  // transactions and retrieve all available data
  while (transactions.length < total_transactions) {
    const paginatedRequest = UrlFetchApp.fetch("https://production.plaid.com/transactions/get", {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({
        client_id: "67f7d02c204d3500278eb772",
        secret: "22426845ed29506ee59faef7d0bbee",
        access_token: accessToken,
        start_date: startDate,
        end_date: endDate,
        options: {
          offset: transactions.length,
        },
      }),
    });

    transactions = transactions.concat(JSON.parse(paginatedRequest.getContentText()).transactions);
  }
  const newSheetName = `${month}-${year}`;
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheetToCheck = spreadsheet.getSheetByName(newSheetName);

  if (sheetToCheck) {
    const response = Browser.msgBox(
      `Sheet ${newSheetName} already exists. Overwrite sheet?`, // message
      Browser.Buttons.OK_CANCEL // button set
    );

    if (response === "ok") {
      spreadsheet.deleteSheet(sheetToCheck);
    } else {
      return;
    }
  }
  const sheet = spreadsheet.insertSheet(); // Creates a sheet with a default name
  sheet.setName(newSheetName); // names new sheet

  sheet.clearContents();
  sheet.appendRow(["Date", "Transaction Description", "Transaction Amount"]);
  Logger.log(transactions);
  transactions.forEach((tx) => {
    sheet.appendRow([tx.date, tx.name, tx.amount]);
  });

  sheet.getRange(2, 6).setValue("Run Categories");
  sheet.getRange(2, 7).insertCheckboxes();

  sheet.getRange(3, 6).setValue("Create Summary Table");
  sheet.getRange(3, 7).insertCheckboxes();

  sheet.getRange(4, 6).setValue("Download Data Again");
  sheet.getRange(4, 7).insertCheckboxes();
  // categorizeTransactions()
}

// list all the accounts linked
function getAccounts() {
  const accessToken = PropertiesService.getScriptProperties().getProperty("PLAID_ACCESS_TOKEN");

  const response = UrlFetchApp.fetch("https://production.plaid.com/accounts/get", {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({
      client_id: "67f7d02c204d3500278eb772",
      secret: "22426845ed29506ee59faef7d0bbee",
      access_token: accessToken,
    }),
  });
  const data = JSON.parse(response.getContentText());
  Logger.log(data.accounts);

  Logger.log(JSON.stringify(data.accounts, null, 2));
}

function overwriteJustNewData(transactions) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("01-2025");
  // sheet.appendRow(['Date', 'Transaction Description', 'Transaction Amount']);
  let currentRow = 2;

  transactions.forEach((tx) => {
    sheet.getRange(currentRow, 1).setValue(tx.date);
    sheet.getRange(currentRow, 2).setValue(tx.name);
    sheet.getRange(currentRow, 3).setValue(tx.amount);
    currentRow++;
  });

  sheet.getRange(2, 6).setValue("Run Categories");
  sheet.getRange(2, 7).insertCheckboxes();

  sheet.getRange(3, 6).setValue("Create Summary Table");
  sheet.getRange(3, 7).insertCheckboxes();

  sheet.getRange(4, 6).setValue("Download Data Again");
  sheet.getRange(4, 7).insertCheckboxes();
}
