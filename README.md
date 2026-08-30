<!-- @format -->

# Plaid Google Sheets budget tracker

This repository contains a personal budgeting workflow built with Google Apps Script, Google Sheets, and Plaid. It connects one or more bank accounts through Plaid Link, downloads transactions into monthly tabs, categorizes them from a keyword table, and compares actual spending with a household budget.

The code currently targets a specific spreadsheet and contains custom sheet names, account names, categories, and formulas. See [Personal assumptions to customize](#personal-assumptions-to-customize) before reusing it.

## What the code does

### 1. Connects bank accounts with Plaid

`launchPlaidLink()` opens `plaidLink.html` in a Google Sheets dialog. The dialog:

1. asks Apps Script to create a Plaid Link token;
2. opens Plaid Link so a bank can be selected;
3. exchanges the returned public token for an access token; and
4. saves the access token in Apps Script's Script Properties.

Multiple Plaid items are supported. Access tokens are stored as a JSON array in `PLAID_ACCESS_TOKENS`; the older single-value `PLAID_ACCESS_TOKEN` property is also read for compatibility. `fullyUnlinkAccount()` removes every stored Plaid item through Plaid and then deletes both token properties.

All Plaid requests currently use the production API (`https://production.plaid.com`).

### 2. Downloads monthly transactions

`getRealTransactions(month, year, append)` fetches accounts and transactions for every saved access token. Requests for different Plaid items and transaction pages are batched where possible.

The function then:

- gives each transaction an account label such as `My Checking (1234)`;
- sorts transactions from oldest to newest, using the Plaid transaction ID as a tie-breaker;
- reverses Plaid's amount convention so money in is positive and spending is negative;
- writes the result to a tab named `MM-YYYY`.

A new transaction tab has these columns:

| Column | Contents                      |
| ------ | ----------------------------- |
| A      | Date                          |
| B      | Transaction Description       |
| C      | Transaction Amount            |
| D      | Account                       |
| E      | Category                      |
| F      | SubCategory                   |
| G      | Notes                         |
| H-I    | Command labels and checkboxes |

The command checkboxes are **Clear Filter**, **Run Categories**, **Create Summary Table**, **Re-Download Data**, and **Create Charts**.

When `append` is `true`, downloading starts on the day after the latest date already in column A and adds rows below the existing transactions. This preserves existing rows, categories, subcategories, and notes. It does not update transactions on an already-downloaded date. A non-append download asks before deleting and rebuilding an existing monthly tab; rebuilding removes manual edits on that tab.

Calling `getRealTransactions()` with no arguments downloads the current month through today. Supplying a month and year requests that entire calendar month.

### 3. Categorizes transactions

`categorizeTransactions()` reads rules from a sheet named `Categories`:

- column A contains the category name;
- columns B onward contain description keywords;
- ordinary keywords match anywhere in the transaction description, without regard to case; and
- a keyword wrapped in double quotes, such as `"whole foods"`, must match the complete description.

Categories are checked in sheet order and the first matching category wins. Unmatched transactions receive `UNCATEGORIZED`. Existing category values are left unchanged, and the Category column receives a dropdown containing the known categories.

### 4. Builds a budget summary

`summarizeByCategory()` writes a summary table to columns K-P of the active transaction tab. It:

- reads budget categories and amounts from the configured household-expenses sheet;
- includes only budget rows
- calculates actual spending per category with spreadsheet formulas;
- adds totals for any manually entered subcategories;
- calculates the remaining amount for each budget category;
- adds `SumZero` and `UNCATEGORIZED` rows; and
- applies green/red conditional formatting to positive and negative balances.

The function also writes household balance and split calculations to columns R-S. If it finds a `Monthly Interest Paid` transaction for the configured checking account, it estimates that account's balance on the transaction date by starting with Plaid's current balance and rolling back all later posted transactions.

### 5. Handles checkbox actions

`onEditFunctions.js` routes checked boxes to the appropriate action:

- **Run Categories** categorizes blank Category cells.
- **Create Summary Table** rebuilds the budget summary.
- summary-table checkboxes hide transaction rows outside the selected categories.
- **Clear Filter** shows every row and clears the summary filter boxes.
- **Re-Download Data** appends newer transaction dates to the current monthly tab.
- **Create Charts** replaces existing charts with doughnut charts for `groceries`, `discretionary`, `tolls`, and `gas` when those categories exist.

A separate `runScript` sheet can provide a one-click monthly workflow. Put a checkbox in `B1` and a value in `B2` formatted as `MM-YYYY`. Checking `B1` downloads or rebuilds that month, categorizes its transactions, and creates its summary.

Because downloads call an external API, run `installOnEditTrigger()` once to create an authorized installable edit trigger. The simple `onEdit` handler deliberately leaves Plaid-powered actions to that trigger.

## Setup

1. Create or open the Google Sheet that will own this script.
2. Open **Extensions → Apps Script** and add the `.js`, `.html`, and `appsscript.json` files from this repository to the bound Apps Script project. The file names do not affect Apps Script execution, but both HTML files have distinct purposes: `plaidLink.html` is used by the spreadsheet dialog, while `index.html` is only a simple informational landing page.
3. In **Project Settings → Script Properties**, add:
   - `PLAID_CLIENT_ID`: your Plaid client ID
   - `PLAID_SECRET`: your Plaid production secret
4. Create a `Categories` tab using the rule layout described above.
5. Create or rename the budget tabs expected by `makeSummaryTable.js`, or update the hard-coded names and cell references listed below.
6. From the Apps Script editor, run `launchPlaidLink()` and authorize the script. Repeat it to connect additional Plaid items.
7. Run `installOnEditTrigger()` once.
8. Either call `getRealTransactions()` from the editor for the current month, call it with a month and year, or configure the optional `runScript` sheet.

Do not commit Plaid secrets or access tokens to this repository. They belong in Script Properties and grant access to sensitive financial data.

## Personal assumptions to customize

The current implementation is tailored to the author's spreadsheet:

| File                                       | Current assumption                                                                                            |
| ------------------------------------------ | ------------------------------------------------------------------------------------------------------------- |
| `makeSummaryTable.js`                      | Budget source tab is `Household Expenses May 2026`; category is in column A, amount in D, and person in H.    |
| `makeSummaryTable.js`                      | Household formulas refer to `Budget May 2026!B13`, `C4`, and `D4`                                             |
| `makeSummaryTable.js` / `bankFunctions.js` | Ending-balance logic expects account `My Account (1234)` and transaction description `Monthly Interest Paid`. |
| `onEditFunctions.js`                       | Charts are created only for groceries, discretionary, tolls, and gas.                                         |

Plaid Link requests the `transactions` product for US institutions and asks for 120 days of history. The app name shown inside Plaid Link is `My Google Sheets Bank App`.

## Main functions

| Function                                   | Purpose                                                      |
| ------------------------------------------ | ------------------------------------------------------------ |
| `launchPlaidLink()`                        | Opens the bank-connection dialog.                            |
| `getRealTransactions(month, year, append)` | Creates or extends a monthly transaction tab.                |
| `categorizeTransactions()`                 | Applies keyword rules to uncategorized rows.                 |
| `summarizeByCategory()`                    | Builds the budget, subcategory, balance, and split summary.  |
| `installOnEditTrigger()`                   | Installs the authorized checkbox handler.                    |
| `getAccounts()`                            | Logs accounts returned by Plaid for troubleshooting.         |
| `checkTokenInfo()`                         | Logs Plaid item information for saved connections.           |
| `fullyUnlinkAccount()`                     | Removes all connected Plaid items and locally stored tokens. |
| `addToSummarySheet()`                      | Runs the older, hard-coded multi-month comparison.           |
| `testOnEdit()`                             | Simulates checking cell `I3` on the active sheet.            |

## Repository layout

| File                  | Responsibility                                                   |
| --------------------- | ---------------------------------------------------------------- |
| `setupPlaid.js`       | Plaid credentials, Link setup, token storage, and unlinking.     |
| `bankFunctions.js`    | Plaid account/transaction requests and monthly sheet creation.   |
| `categorize.js`       | Keyword-based transaction categorization.                        |
| `makeSummaryTable.js` | Budget summaries, balance calculations, and formatting.          |
| `onEditFunctions.js`  | Checkbox automation and installable trigger setup.               |
| `createChart.js`      | Per-category spent-versus-remaining doughnut charts.             |
| `combineMonths.js`    | Legacy multi-month budget comparison.                            |
| `testOnEdit.js`       | Manual edit-event test helper.                                   |
| `plaidLink.html`      | Plaid Link dialog used inside Google Sheets.                     |
| `index.html`          | Standalone informational page, suitable for GitHub Pages.        |
| `appsscript.json`     | Apps Script manifest using V8 and the America/New_York timezone. |

## Known limitations

- The project uses Plaid's legacy `/transactions/get` flow rather than Transactions Sync or webhooks, so updates are manual.
- Append mode works by date. Late-posting or modified transactions for dates already present will not be picked up by an append.
- Rebuilding an existing month deletes that sheet and its manual categories, subcategories, notes, formatting, and charts.
- A full-month request whose end date is in the future may be rejected by Plaid; the no-argument current-month path ends at today.
- There is no automated test suite. `testOnEdit()` is an Apps Script helper and acts on the active spreadsheet.
- The project has no custom spreadsheet menu; primary actions are run from Apps Script or through sheet checkboxes.

This is a private personal-finance tool, not a hosted financial service or a general-purpose Plaid integration.
