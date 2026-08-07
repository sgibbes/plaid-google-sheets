/** @format */
const TRANSACTION_HEADERS = [
  "Date",
  "Transaction Description",
  "Transaction Amount",
  "Account",
  "Category",
  "SubCategory",
  "Notes",
];
const TRANSACTION_CATEGORY_COL = 5;
const TRANSACTION_SUBCATEGORY_COL = 6;
const TRANSACTION_NOTES_COL = 7;
const CONTROL_LABEL_COL = 8;
const CONTROL_CHECKBOX_COL = 9;

function createDataInSheet(newSheetName, txAdjusted) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.insertSheet(newSheetName);
  const transactionRows = txAdjusted.map((tx) => getTransactionRow_(tx, {}));
  const allRows = [TRANSACTION_HEADERS, ...transactionRows];

  sheet.getRange(1, 1, allRows.length, TRANSACTION_HEADERS.length).setValues(allRows);
  sheet.getRange(1, CONTROL_LABEL_COL, 5, 2).setValues([
    ["Clear Filter", false],
    ["Run Categories", false],
    ["Create Summary Table", false],
    ["Re-Download Data", false],
    ["Create Charts", false],
  ]);
  sheet.getRange(1, CONTROL_CHECKBOX_COL, 5, 1).insertCheckboxes();

  sheet.autoResizeColumns(1, 10);
  return sheet;
}

function getTransactionRow_(tx, existingNotes = {}, existingSubcategories = {}) {
  const key = getTransactionNoteKey_({
    date: tx.date,
    name: tx.name,
    amount: tx.amount,
    accountName: tx.accountName,
  });

  return [
    tx.date,
    tx.name,
    tx.amount,
    tx.accountName,
    "",
    existingSubcategories[key] || "",
    existingNotes[key] || "",
  ];
}

function shouldExcludeTransaction_(tx) {
  const description = String(tx.name || "").trim().toLowerCase();
  const amountInCents = Math.round(Number(tx.amount) * 100);

  return description === "best buy" && amountInCents === -121794;
}

function isMonthlyInterestEndingBalanceTransaction_(tx) {
  const description = String(tx.name || "").trim().toLowerCase();
  const accountName = String(tx.accountName || "").trim().toLowerCase();

  return description === "monthly interest paid" && accountName === "360 checking (1147)";
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

function getExistingTransactionValues_(sheet, valueColumnName) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    return {};
  }

  const headers = data[0];
  const dateCol = headers.indexOf("Date");
  const descriptionCol = headers.indexOf("Transaction Description");
  const amountCol = headers.indexOf("Transaction Amount");
  const accountCol = headers.indexOf("Account");
  const valueCol = headers.indexOf(valueColumnName);

  if ([dateCol, descriptionCol, amountCol, accountCol, valueCol].some((col) => col === -1)) {
    return {};
  }

  return data.slice(1).reduce((values, row) => {
    const value = row[valueCol];
    if (!value) {
      return values;
    }

    const key = getTransactionNoteKey_({
      date: row[dateCol],
      name: row[descriptionCol],
      amount: row[amountCol],
      accountName: row[accountCol],
    });
    values[key] = value;
    return values;
  }, {});
}

function getExistingNotes_(sheet) {
  return getExistingTransactionValues_(sheet, "Notes");
}

function getExistingSubcategories_(sheet) {
  return getExistingTransactionValues_(sheet, "SubCategory");
}

function getTransactionAppendState_(sheet) {
  const rowCount = Math.max(sheet.getLastRow() - 1, 0);
  if (rowCount === 0) {
    return { lastTransactionRow: 1, latestDate: null };
  }

  const dateValues = sheet.getRange(2, 1, rowCount, 1).getValues();
  let lastTransactionRow = 1;
  let latestDate = null;

  dateValues.forEach((row, index) => {
    const value = row[0];
    if (value === "") {
      return;
    }

    lastTransactionRow = index + 2;
    const isoDate = getIsoDateValue_(value);

    if (isoDate && (!latestDate || isoDate > latestDate)) {
      latestDate = isoDate;
    }
  });

  return { lastTransactionRow, latestDate };
}

function getIsoDateValue_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }

  const match = String(value || "").match(/^(\d{4}-\d{2}-\d{2})/);
  return match ? match[1] : null;
}

function getNextIsoDate_(isoDate) {
  const [year, month, day] = isoDate.split("-").map(Number);
  return new Date(Date.UTC(year, month - 1, day + 1)).toISOString().slice(0, 10);
}

function ensureTransactionLayout_(sheet) {
  let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  let categoryCol = headers.indexOf("Category") + 1;
  let notesCol = headers.indexOf("Notes") + 1;

  if (!categoryCol) {
    categoryCol = TRANSACTION_CATEGORY_COL;
    sheet.getRange(1, TRANSACTION_CATEGORY_COL).setValue("Category");
    headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    notesCol = headers.indexOf("Notes") + 1;
  }

  // Upgrade the old layout by inserting SubCategory at F exactly once.
  if (notesCol === TRANSACTION_SUBCATEGORY_COL) {
    sheet.insertColumnAfter(TRANSACTION_CATEGORY_COL);
    notesCol++;
    headers.splice(TRANSACTION_SUBCATEGORY_COL - 1, 0, "");
  }

  if (headers[TRANSACTION_SUBCATEGORY_COL - 1] !== "SubCategory") {
    sheet.getRange(1, TRANSACTION_SUBCATEGORY_COL).setValue("SubCategory");
  }

  if (!notesCol) {
    sheet.getRange(1, TRANSACTION_NOTES_COL).setValue("Notes");
  }
}

function getAccountDisplayName_(account) {
  const name = account.official_name || account.name || account.subtype || account.account_id;

  return account.mask ? `${name} (${account.mask})` : name;
}

function getPlaidAccountConnectionByName_(targetAccountName) {
  const { clientId, secret } = getPlaidCredentials_();
  const accessTokens = getPlaidAccessTokens_();
  if (accessTokens.length === 0) {
    throw new Error("No access token found. Run launchPlaidLink() first.");
  }

  const responses = UrlFetchApp.fetchAll(
    accessTokens.map((accessToken) => ({
      url: "https://production.plaid.com/accounts/get",
      method: "post",
      contentType: "application/json",
      muteHttpExceptions: true,
      payload: JSON.stringify({
        client_id: clientId,
        secret,
        access_token: accessToken,
      }),
    })),
  );
  const normalizedTarget = String(targetAccountName).trim().toLowerCase();

  for (let index = 0; index < responses.length; index++) {
    const response = responses[index];
    const text = response.getContentText();
    const data = JSON.parse(text);
    if (response.getResponseCode() >= 400 || !data.accounts) {
      throw new Error("Plaid accounts/get failed: " + text);
    }

    const account = data.accounts.find(
      (candidate) => getAccountDisplayName_(candidate).trim().toLowerCase() === normalizedTarget,
    );
    if (account) {
      return { accessToken: accessTokens[index], account };
    }
  }

  return null;
}

function getHistoricalBalanceForAccountOnDate_(targetAccountName, targetDate) {
  const connection = getPlaidAccountConnectionByName_(targetAccountName);
  if (!connection || !connection.account.balances) {
    return null;
  }

  const currentBalance = connection.account.balances.current;
  const isoTargetDate = getIsoDateValue_(targetDate);
  if (currentBalance === null || currentBalance === undefined || !isoTargetDate) {
    return null;
  }

  const today = new Date().toISOString().slice(0, 10);
  if (isoTargetDate >= today) {
    return currentBalance;
  }

  const transactionsAfterTarget = getTransactionsForAccessTokens_(
    [connection.accessToken],
    getNextIsoDate_(isoTargetDate),
    today,
  );
  const laterPostedTotal = transactionsAfterTarget
    .filter(
      (transaction) =>
        transaction.account_id === connection.account.account_id && !transaction.pending,
    )
    .reduce((total, transaction) => total + Number(transaction.amount || 0), 0);

  // Plaid transaction amounts are positive for outflows and negative for
  // inflows, so adding later amounts rolls the current balance backward.
  return currentBalance + laterPostedTotal;
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

function getTransactionsForAccessTokens_(accessTokens, startDate, endDate) {
  const { clientId, secret } = getPlaidCredentials_();
  const pageSize = 500;
  const request = (url, accessToken, options) => ({
    url,
    method: "post",
    contentType: "application/json",
    muteHttpExceptions: true,
    payload: JSON.stringify({
      client_id: clientId,
      secret,
      access_token: accessToken,
      ...(url.endsWith("/transactions/get")
        ? { start_date: startDate, end_date: endDate, options }
        : {}),
    }),
  });
  const parseResponse = (response, operation) => {
    const text = response.getContentText();
    const data = JSON.parse(text);
    if (response.getResponseCode() >= 400) {
      throw new Error("Plaid " + operation + " failed: " + text);
    }
    return data;
  };

  // Account metadata and the first transaction page for every linked item can
  // be fetched concurrently.
  const initialRequests = accessTokens.flatMap((accessToken) => [
    request("https://production.plaid.com/accounts/get", accessToken),
    request("https://production.plaid.com/transactions/get", accessToken, {
      offset: 0,
      count: pageSize,
    }),
  ]);
  const initialResponses = UrlFetchApp.fetchAll(initialRequests);
  const itemResults = accessTokens.map((accessToken, index) => {
    const accountData = parseResponse(initialResponses[index * 2], "accounts/get");
    const transactionData = parseResponse(
      initialResponses[index * 2 + 1],
      "transactions/get",
    );
    if (!accountData.accounts || !transactionData.transactions) {
      throw new Error("Plaid returned an incomplete accounts or transactions response.");
    }

    const accountLookup = accountData.accounts.reduce((lookup, account) => {
      lookup[account.account_id] = getAccountDisplayName_(account);
      return lookup;
    }, {});
    return {
      accessToken,
      accountLookup,
      total: transactionData.total_transactions,
      transactions: transactionData.transactions,
    };
  });

  const additionalPages = [];
  itemResults.forEach((item, itemIndex) => {
    for (let offset = pageSize; offset < item.total; offset += pageSize) {
      additionalPages.push({
        itemIndex,
        request: request("https://production.plaid.com/transactions/get", item.accessToken, {
          offset,
          count: pageSize,
        }),
      });
    }
  });

  if (additionalPages.length > 0) {
    const responses = UrlFetchApp.fetchAll(additionalPages.map((page) => page.request));
    responses.forEach((response, index) => {
      const data = parseResponse(response, "transactions/get");
      if (!data.transactions) {
        throw new Error("Plaid returned an incomplete transactions response.");
      }
      itemResults[additionalPages[index].itemIndex].transactions.push(...data.transactions);
    });
  }

  return itemResults.flatMap((item) =>
    item.transactions.map((transaction) => ({
      ...transaction,
      accountName:
        item.accountLookup[transaction.account_id] || transaction.account_id || "",
    })),
  );
}

function getRealTransactions(userMonth = null, userYear = null, append = false) {
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

  const newSheetName = `${month}-${year}`;
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheetToCheck = spreadsheet.getSheetByName(newSheetName);
  let appendState = null;

  if (append && sheetToCheck) {
    ensureTransactionLayout_(sheetToCheck);
    appendState = getTransactionAppendState_(sheetToCheck);
    const today = new Date().toISOString().slice(0, 10);
    endDate = endDate > today ? today : endDate;

    if (appendState.latestDate) {
      startDate = getNextIsoDate_(appendState.latestDate);
    }

    if (startDate > endDate) {
      spreadsheet.toast("No new transaction dates to download.");
      return;
    }
  }

  const transactions = getTransactionsForAccessTokens_(accessTokens, startDate, endDate);
  const txSorted = [...transactions].sort((a, b) => {
    const dateA = new Date(a.datetime || a.date);
    const dateB = new Date(b.datetime || b.date);

    if (dateA.getTime() !== dateB.getTime()) {
      return dateA - dateB;
    }

    return (a.transaction_id || "").localeCompare(b.transaction_id || "");
  });
  const txAdjusted = txSorted
    .map((tx) => {
      const adjustedAmount = tx.amount < 0 ? Math.abs(tx.amount) : -tx.amount;

      return { ...tx, amount: adjustedAmount };
    })
    .filter((tx) => !shouldExcludeTransaction_(tx));

  if (sheetToCheck && !append) {
    const response = Browser.msgBox(
      `Sheet ${newSheetName} already exists. Overwrite sheet?`, // message
      Browser.Buttons.OK_CANCEL // button set
    );

    if (response === "ok") {
      spreadsheet.deleteSheet(sheetToCheck);
      return createDataInSheet(newSheetName, txAdjusted);
    } else {
      return null;
    }
  }

  // the sheet does not already exist
  if (!sheetToCheck) {
    return createDataInSheet(newSheetName, txAdjusted);
  }

  if (append) {
    const values = txAdjusted.map((tx) => getTransactionRow_(tx));

    if (values.length > 0) {
      sheetToCheck
        .getRange(
          appendState.lastTransactionRow + 1,
          1,
          values.length,
          TRANSACTION_HEADERS.length,
        )
        .setValues(values);
      spreadsheet.toast(values.length + " new transactions appended.");
    } else {
      spreadsheet.toast("No new transactions found.");
    }

    return sheetToCheck;
  }

  return null;
}

// list all the accounts linked
function getAccounts() {
  const { clientId, secret } = getPlaidCredentials_();
  const accessTokens = getPlaidAccessTokens_();

  if (accessTokens.length === 0) {
    throw new Error("No access token found. Run launchPlaidLink() first.");
  }

  const responses = UrlFetchApp.fetchAll(
    accessTokens.map((accessToken) => ({
      url: "https://production.plaid.com/accounts/get",
      method: "post",
      contentType: "application/json",
      muteHttpExceptions: true,
      payload: JSON.stringify({
        client_id: clientId,
        secret: secret,
        access_token: accessToken,
      }),
    })),
  );

  responses.forEach((response) => {
    const data = JSON.parse(response.getContentText());
    if (response.getResponseCode() >= 400 || !data.accounts) {
      throw new Error("Plaid accounts/get failed: " + response.getContentText());
    }
    Logger.log(data.accounts);
    Logger.log(JSON.stringify(data.accounts, null, 2));
  });
}
