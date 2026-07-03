/** @format */
const TRANSACTION_HEADERS = ["Date", "Transaction Description", "Transaction Amount", "Account", "Category", "Notes"];
const CONTROL_LABEL_COL = 7;
const CONTROL_CHECKBOX_COL = 8;

function createDataInSheet(newSheetName, txAdjusted) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.insertSheet(); // Creates a sheet with a default name
  sheet.setName(newSheetName); // names new sheet

  sheet.clearContents();
  sheet.appendRow(TRANSACTION_HEADERS);
  txAdjusted.forEach((tx) => {
    sheet.appendRow(getTransactionRow_(tx, {}));
  });

  sheet.getRange(1, CONTROL_LABEL_COL).setValue("Clear Filter");
  sheet.getRange(1, CONTROL_CHECKBOX_COL).insertCheckboxes();

  sheet.getRange(2, CONTROL_LABEL_COL).setValue("Run Categories");
  sheet.getRange(2, CONTROL_CHECKBOX_COL).insertCheckboxes();

  sheet.getRange(3, CONTROL_LABEL_COL).setValue("Create Summary Table");
  sheet.getRange(3, CONTROL_CHECKBOX_COL).insertCheckboxes();

  sheet.getRange(4, CONTROL_LABEL_COL).setValue("Re-Download Data");
  sheet.getRange(4, CONTROL_CHECKBOX_COL).insertCheckboxes();

  sheet.getRange(5, CONTROL_LABEL_COL).setValue("Create Charts");
  sheet.getRange(5, CONTROL_CHECKBOX_COL).insertCheckboxes();

  sheet.autoResizeColumns(1, 10);
}

function getTransactionRow_(tx, existingNotes) {
  const key = getTransactionNoteKey_({
    date: tx.date,
    name: tx.name,
    amount: tx.amount,
    accountName: tx.accountName,
  });

  return [tx.date, tx.name, tx.amount, tx.accountName, "", existingNotes[key] || ""];
}

function getTransactionNoteKey_(tx) {
  return [
    getTransactionNoteKeyValue_(tx.date, true),
    getTransactionNoteKeyValue_(tx.name),
    getTransactionNoteKeyValue_(tx.amount),
    getTransactionNoteKeyValue_(tx.accountName),
  ].join("|");
}

function getTransactionNoteKeyValue_(value, isDate) {
  if (isDate && value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }

  return String(value || "").trim();
}

function getExistingNotes_(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    return {};
  }

  const headers = data[0];
  const dateCol = headers.indexOf("Date");
  const descriptionCol = headers.indexOf("Transaction Description");
  const amountCol = headers.indexOf("Transaction Amount");
  const accountCol = headers.indexOf("Account");
  const notesCol = headers.indexOf("Notes");

  if ([dateCol, descriptionCol, amountCol, accountCol, notesCol].some((col) => col === -1)) {
    return {};
  }

  return data.slice(1).reduce((notes, row) => {
    const note = row[notesCol];
    if (!note) {
      return notes;
    }

    const key = getTransactionNoteKey_({
      date: row[dateCol],
      name: row[descriptionCol],
      amount: row[amountCol],
      accountName: row[accountCol],
    });
    notes[key] = note;
    return notes;
  }, {});
}

function ensureTransactionLayout_(sheet) {
  let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  let accountCol = headers.indexOf("Account") + 1;
  let categoryCol = headers.indexOf("Category") + 1;
  let notesCol = headers.indexOf("Notes") + 1;

  if (!categoryCol) {
    categoryCol = accountCol ? accountCol + 1 : 5;
    sheet.getRange(1, categoryCol).setValue("Category");
    headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    notesCol = headers.indexOf("Notes") + 1;
  }

  if (!notesCol) {
    sheet.insertColumnAfter(categoryCol);
    sheet.getRange(1, categoryCol + 1).setValue("Notes");
  }
}

function getAccountDisplayName_(account) {
  const name = account.official_name || account.name || account.subtype || account.account_id;

  return account.mask ? `${name} (${account.mask})` : name;
}

function getAccountLookupForAccessToken_(accessToken) {
  const { clientId, secret } = getPlaidCredentials_();
  const response = UrlFetchApp.fetch("https://production.plaid.com/accounts/get", {
    method: "post",
    contentType: "application/json",
    muteHttpExceptions: true,
    payload: JSON.stringify({
      client_id: clientId,
      secret: secret,
      access_token: accessToken,
    }),
  });

  const data = JSON.parse(response.getContentText());
  if (response.getResponseCode() >= 400 || !data.accounts) {
    throw new Error("Plaid accounts/get failed: " + response.getContentText());
  }

  return data.accounts.reduce((lookup, account) => {
    lookup[account.account_id] = getAccountDisplayName_(account);
    return lookup;
  }, {});
}

function getTransactionsForAccessToken_(accessToken, startDate, endDate) {
  const { clientId, secret } = getPlaidCredentials_();
  const accountLookup = getAccountLookupForAccessToken_(accessToken);
  const transactions = [];
  let totalTransactions = null;

  while (totalTransactions === null || transactions.length < totalTransactions) {
    const response = UrlFetchApp.fetch("https://production.plaid.com/transactions/get", {
      method: "post",
      contentType: "application/json",
      muteHttpExceptions: true,
      payload: JSON.stringify({
        client_id: clientId,
        secret: secret,
        access_token: accessToken,
        start_date: startDate,
        end_date: endDate,
        options: {
          offset: transactions.length,
        },
      }),
    });

    const data = JSON.parse(response.getContentText());
    if (response.getResponseCode() >= 400 || !data.transactions) {
      throw new Error("Plaid transactions/get failed: " + response.getContentText());
    }

    transactions.push(
      ...data.transactions.map((transaction) => ({
        ...transaction,
        accountName: accountLookup[transaction.account_id] || transaction.account_id || "",
      }))
    );
    totalTransactions = data.total_transactions;
  }

  return transactions;
}

function getRealTransactions(userMonth = null, userYear = null, append = false) {
  Logger.log(append);
  const accessTokens = getPlaidAccessTokens_();
  if (accessTokens.length === 0) {
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

  const transactions = accessTokens.flatMap((accessToken) => getTransactionsForAccessToken_(accessToken, startDate, endDate));
  const txSorted = [...transactions].sort((a, b) => {
    const dateA = new Date(a.datetime || a.date);
    const dateB = new Date(b.datetime || b.date);

    if (dateA.getTime() !== dateB.getTime()) {
      return dateA - dateB;
    }

    return (a.transaction_id || "").localeCompare(b.transaction_id || "");
  });
  Logger.log(txSorted.filter((x) => x.pending));
  const txAdjusted = txSorted.map((tx) => {
    const adjustedAmount = tx.amount < 0 ? Math.abs(tx.amount) : -tx.amount;

    return { ...tx, amount: adjustedAmount };
  });

  const newSheetName = `${month}-${year}`;
  Logger.log({ newSheetName });

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheetToCheck = spreadsheet.getSheetByName(newSheetName);
  Logger.log({ sheetToCheck });
  Logger.log({ append });

  if (sheetToCheck && !append) {
    Logger.log("here");

    const response = Browser.msgBox(
      `Sheet ${newSheetName} already exists. Overwrite sheet?`, // message
      Browser.Buttons.OK_CANCEL // button set
    );

    if (response === "ok") {
      spreadsheet.deleteSheet(sheetToCheck);

      createDataInSheet(newSheetName, txAdjusted);
    } else if (!sheetToCheck) {
      createDataInSheet(newSheetName, txAdjusted);
    } else {
      return;
    }
  }

  // the sheet does not already exist
  if (!sheetToCheck) {
    const sheet = spreadsheet.insertSheet(); // Creates a sheet with a default name
    sheet.setName(newSheetName); // names new sheet

    sheet.clearContents();
    sheet.appendRow(TRANSACTION_HEADERS);
    txAdjusted.forEach((tx) => {
      sheet.appendRow(getTransactionRow_(tx, {}));
    });

    sheet.getRange(1, CONTROL_LABEL_COL).setValue("Clear Filter");
    sheet.getRange(1, CONTROL_CHECKBOX_COL).insertCheckboxes();

    sheet.getRange(2, CONTROL_LABEL_COL).setValue("Run Categories");
    sheet.getRange(2, CONTROL_CHECKBOX_COL).insertCheckboxes();

    sheet.getRange(3, CONTROL_LABEL_COL).setValue("Create Summary Table");
    sheet.getRange(3, CONTROL_CHECKBOX_COL).insertCheckboxes();

    sheet.getRange(4, CONTROL_LABEL_COL).setValue("Re-Download Data");
    sheet.getRange(4, CONTROL_CHECKBOX_COL).insertCheckboxes();

    sheet.getRange(5, CONTROL_LABEL_COL).setValue("Create Charts");
    sheet.getRange(5, CONTROL_CHECKBOX_COL).insertCheckboxes();

    sheet.autoResizeColumns(1, 10);
  }

  if (append) {
    const sheet = sheetToCheck || spreadsheet.getActiveSheet();
    ensureTransactionLayout_(sheet);
    const existingNotes = getExistingNotes_(sheet);
    const values = txAdjusted.map((tx) => getTransactionRow_(tx, existingNotes));

    sheet.getRange(1, 1, 1, TRANSACTION_HEADERS.length).setValues([TRANSACTION_HEADERS]);
    sheet.getRange(2, 1, sheet.getMaxRows() - 1, TRANSACTION_HEADERS.length).clearContent();

    if (values.length > 0) {
      sheet.getRange(2, 1, values.length, TRANSACTION_HEADERS.length).setValues(values); //row, col, numRows, num cols
    }
  }
}

// list all the accounts linked
function getAccounts() {
  const { clientId, secret } = getPlaidCredentials_();
  const accessTokens = getPlaidAccessTokens_();

  if (accessTokens.length === 0) {
    throw new Error("No access token found. Run launchPlaidLink() first.");
  }

  accessTokens.forEach((accessToken) => {
    const response = UrlFetchApp.fetch("https://production.plaid.com/accounts/get", {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({
        client_id: clientId,
        secret: secret,
        access_token: accessToken,
      }),
    });
    const data = JSON.parse(response.getContentText());
    Logger.log(data.accounts);

    Logger.log(JSON.stringify(data.accounts, null, 2));
  });
}
