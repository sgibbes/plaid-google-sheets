/** @format */

function getRealTransactions(userMonth = null, userYear = null) {
  const accessToken = PropertiesService.getScriptProperties().getProperty("PLAID_ACCESS_TOKEN");
  if (!accessToken) {
    throw new Error("No access token found. Run launchPlaidLink() first.");
  }

  // // if I want to get the current month of data:
  const formatMonth = (datesMonth) => {
    const m = datesMonth.getMonth() + 1;
    return m < 10 ? `0${m}` : `${m}`;
  };

  const getEndDayNum = (year, month) => {
    const d = new Date(year, month, 0);
    return d.getDate();
  };

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

  const response = UrlFetchApp.fetch("https://production.plaid.com/transactions/get", {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({
      client_id: x,
      secret: x,
      access_token: accessToken,
      start_date: startDate,
      end_date: endDate,
    }),
  });

  let transactions = JSON.parse(response.getContentText()).transactions;

  const total_transactions = JSON.parse(response.getContentText()).total_transactions;

  while (transactions.length < total_transactions) {
    const paginatedRequest = UrlFetchApp.fetch("https://production.plaid.com/transactions/get", {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({
        client_id: x,
        secret: "x",
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
  const txSorted = [...transactions];
  txSorted.sort((a, b) => {
    const aDate = new Date(a.date);
    const bDate = new Date(b.date);
    return bDate - aDate;
  });
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

  txSorted.forEach((tx) => {
    sheet.appendRow([tx.date, tx.name, tx.amount]);
  });

  sheet.getRange(2, 6).setValue("Run Categories");
  sheet.getRange(2, 7).insertCheckboxes();

  sheet.getRange(3, 6).setValue("Create Summary Table");
  sheet.getRange(3, 7).insertCheckboxes();

  sheet.getRange(4, 6).setValue("Download Data Again");
  sheet.getRange(4, 7).insertCheckboxes();
}

// list all the accounts linked
function getAccounts() {
  const accessToken = PropertiesService.getScriptProperties().getProperty("PLAID_ACCESS_TOKEN");

  const response = UrlFetchApp.fetch("https://production.plaid.com/accounts/get", {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({
      client_id: x,
      secret: "x",
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
    const adjustedAmount = tx.amount < 0 ? Math.abs(tx.amount) : -tx.amount;

    sheet.getRange(currentRow, 3).setValue(adjustedAmount);
    currentRow++;
  });

  sheet.getRange(2, 6).setValue("Run Categories");
  sheet.getRange(2, 7).insertCheckboxes();

  sheet.getRange(3, 6).setValue("Create Summary Table");
  sheet.getRange(3, 7).insertCheckboxes();

  sheet.getRange(4, 6).setValue("Download Data Again");
  sheet.getRange(4, 7).insertCheckboxes();
}
