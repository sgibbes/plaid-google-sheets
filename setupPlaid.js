/** @format */

function createLinkToken() {
  const user = {
    client_user_id: Utilities.getUuid(), // A unique ID for the current user
  };

  const response = UrlFetchApp.fetch("https://production.plaid.com/link/token/create", {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({
      x: x,
      secret: "x",
      user: user,
      client_name: "My Google Sheets Bank App",
      products: ["transactions"],
      transactions: {
        days_requested: 120,
      },
      country_codes: ["US"],
      language: "en",
    }),
  });

  const token = JSON.parse(response.getContentText()).link_token;
  Logger.log(token);

  return token;
}

function launchPlaidLink() {
  const html = HtmlService.createHtmlOutputFromFile("plaidLink").setWidth(600).setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, "Connect Your Bank");
}

function handlePublicToken(publicToken) {
  const response = UrlFetchApp.fetch("https://production.plaid.com/item/public_token/exchange", {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({
      x: x,
      secret: "x",
      public_token: publicToken,
    }),
  });

  const data = JSON.parse(response.getContentText());
  const accessToken = data.access_token;

  PropertiesService.getScriptProperties().setProperty("PLAID_ACCESS_TOKEN", accessToken);
}

function checkTokenInfo() {
  const accessToken = PropertiesService.getScriptProperties().getProperty("PLAID_ACCESS_TOKEN");

  const res = UrlFetchApp.fetch("https://production.plaid.com/item/get", {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({
      x: x,
      secret: "x",
      access_token: accessToken,
    }),
  });

  Logger.log(res.getContentText());
}

function fullyUnlinkAccount() {
  const accessToken = PropertiesService.getScriptProperties().getProperty("PLAID_ACCESS_TOKEN");
  if (!accessToken) {
    Logger.log("No access token found to unlink.");
    return;
  }

  const url = "https://production.plaid.com/item/remove";
  const payload = {
    x: x,
    secret: "x",
    access_token: accessToken,
  };

  const response = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
  });

  const result = JSON.parse(response.getContentText());
  Logger.log("Unlink response: " + JSON.stringify(result));

  PropertiesService.getScriptProperties().deleteProperty("PLAID_ACCESS_TOKEN");
}
