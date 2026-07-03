/** @format */

function getPlaidCredentials_() {
  const props = PropertiesService.getScriptProperties();
  const clientId = props.getProperty("PLAID_CLIENT_ID");
  const secret = props.getProperty("PLAID_SECRET");

  if (!clientId || !secret) {
    throw new Error("Missing PLAID_CLIENT_ID or PLAID_SECRET in Script Properties.");
  }

  return { clientId, secret };
}

function getPlaidAccessTokens_() {
  const props = PropertiesService.getScriptProperties();
  const tokensJson = props.getProperty("PLAID_ACCESS_TOKENS");
  const legacyToken = props.getProperty("PLAID_ACCESS_TOKEN");
  let tokens = [];

  if (tokensJson) {
    try {
      tokens = JSON.parse(tokensJson);
    } catch (error) {
      throw new Error("PLAID_ACCESS_TOKENS is not valid JSON: " + error.message);
    }
  }

  if (legacyToken) {
    tokens.push(legacyToken);
  }

  return [...new Set(tokens.filter(Boolean))];
}

function savePlaidAccessToken_(accessToken) {
  const props = PropertiesService.getScriptProperties();
  const tokens = getPlaidAccessTokens_();

  if (!tokens.includes(accessToken)) {
    tokens.push(accessToken);
  }

  props.setProperty("PLAID_ACCESS_TOKENS", JSON.stringify(tokens));
  props.setProperty("PLAID_ACCESS_TOKEN", accessToken);
}

function createLinkToken() {
  const { clientId, secret } = getPlaidCredentials_();
  const user = {
    client_user_id: Utilities.getUuid(), // A unique ID for the current user
  };

  const response = UrlFetchApp.fetch("https://production.plaid.com/link/token/create", {
    method: "post",
    contentType: "application/json",
    muteHttpExceptions: true,
    payload: JSON.stringify({
      client_id: clientId,
      secret: secret,
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

  const data = JSON.parse(response.getContentText());
  if (response.getResponseCode() >= 400 || !data.link_token) {
    throw new Error("Plaid link/token/create failed: " + response.getContentText());
  }

  return data.link_token;
}

function launchPlaidLink() {
  const html = HtmlService.createHtmlOutputFromFile("plaidLink").setWidth(600).setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, "Connect Your Bank");
}

function handlePublicToken(publicToken) {
  const { clientId, secret } = getPlaidCredentials_();

  const response = UrlFetchApp.fetch("https://production.plaid.com/item/public_token/exchange", {
    method: "post",
    contentType: "application/json",
    muteHttpExceptions: true,
    payload: JSON.stringify({
      client_id: clientId,
      secret: secret,
      public_token: publicToken,
    }),
  });

  const data = JSON.parse(response.getContentText());
  if (response.getResponseCode() >= 400 || !data.access_token) {
    throw new Error("Plaid public_token exchange failed: " + response.getContentText());
  }

  const accessToken = data.access_token;

  savePlaidAccessToken_(accessToken);
}

function checkTokenInfo() {
  const { clientId, secret } = getPlaidCredentials_();
  const accessTokens = getPlaidAccessTokens_();

  if (accessTokens.length === 0) {
    throw new Error("No access token found. Run launchPlaidLink() first.");
  }

  accessTokens.forEach((accessToken) => {
    const res = UrlFetchApp.fetch("https://production.plaid.com/item/get", {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({
        client_id: clientId,
        secret: secret,
        access_token: accessToken,
      }),
    });

    Logger.log(res.getContentText());
  });
}

function fullyUnlinkAccount() {
  const { clientId, secret } = getPlaidCredentials_();
  const props = PropertiesService.getScriptProperties();
  const accessTokens = getPlaidAccessTokens_();

  if (accessTokens.length === 0) {
    Logger.log("No access token found to unlink.");
    return;
  }

  const url = "https://production.plaid.com/item/remove";
  accessTokens.forEach((accessToken) => {
    const payload = {
      client_id: clientId,
      secret: secret,
      access_token: accessToken,
    };

    const response = UrlFetchApp.fetch(url, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
    });

    const result = JSON.parse(response.getContentText());
    Logger.log("Unlink response: " + JSON.stringify(result));
  });

  props.deleteProperty("PLAID_ACCESS_TOKEN");
  props.deleteProperty("PLAID_ACCESS_TOKENS");
}
