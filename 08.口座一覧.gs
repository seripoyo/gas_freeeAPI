/******************************************************************
summary       |口座一覧取得
requestUrl    |https://api.freee.co.jp/api/1/walletables?company_id={companyId}&with_balance=true
method        |GET
******************************************************************/

function getAllPaymentAccounts() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("売上履歴");
  var lastRow = sheet.getLastRow();
  
  // P列のデータを取得（2行目から最終行まで）
  var paymentAccountsRange = sheet.getRange(2, 16, lastRow - 1);
  var paymentAccountsData = paymentAccountsRange.getValues();

  // 配列を平坦化し、空ではない値のみを抽出
  var paymentAccounts = paymentAccountsData.flat().filter(function(account) {
    return account !== "";
  });
    Logger.log(paymentAccounts);
  return paymentAccounts;
}

/******************************************************************
概要：スプシP列に存在する口座を取得し、API取得の一覧と照合し存在しない口座は新規作成。
　　　その後、新規作成口座も含め再度口座一覧を取得し、各口座を保存する。
method        |GET
******************************************************************/
function manageWalletables() {
  var freeeApp = getService();
  var accessToken = freeeApp.getAccessToken();
  var companyId = getSelectedCompanyId();
  var headers = { "Authorization": "Bearer " + accessToken };
  
  // 既存の口座一覧を取得
  var walletables = getExistingWalletables(companyId, headers);

  // 売上履歴シートから決済口座名を取得
  var paymentAccounts = getAllPaymentAccounts();

  // 存在しない口座を登録
  paymentAccounts.forEach(function(accountName) {
    if (!walletables.some(w => w.name === accountName)) {
      registerNewWalletable(companyId, accountName, accessToken);
    }
  });

  // 更新された口座一覧を取得
  var updatedWalletables = getExistingWalletables(companyId, headers);

  // 口座情報を保存
  saveWalletableData(updatedWalletables);
}
// =========================================================
// 既存の口座一覧を取得
// =========================================================
function getExistingWalletables(companyId, headers) {
  var requestUrl = "https://api.freee.co.jp/api/1/walletables?company_id=" + companyId;
  var options = { "method": "get", "headers": headers };
  var response = UrlFetchApp.fetch(requestUrl, options).getContentText();
  var walletablesResponse = JSON.parse(response);
  return walletablesResponse.walletables;
}
// =========================================================
// 新しい口座を登録
// =========================================================
function registerNewWalletable(companyId, walletableName, accessToken) {
  var requestUrl = "https://api.freee.co.jp/api/1/walletables";
  var requestBody = {
    "company_id": companyId,
    "name": walletableName,
    "type": "wallet",
    "is_asset": true
  };
  var options = {
    "method": "post",
    "headers": { "Authorization": "Bearer " + accessToken },
    "contentType": "application/json",
    "payload": JSON.stringify(requestBody),
    "muteHttpExceptions": true
  };

  var response = UrlFetchApp.fetch(requestUrl, options);
  var statusCode = response.getResponseCode();
  var responseContent = JSON.parse(response.getContentText());

  if (statusCode !== 201) {
    Logger.log("口座の登録に失敗しました: " + walletableName + " - " + responseContent.errors.map(e => e.messages).join("; "));
  }
}
// =========================================================
// 口座情報を保存
// =========================================================
function saveWalletableData(walletables) {
  var walletablesData = walletables.map(function(w) {
    return { id: w.id, name: w.name, type: w.type };
  });

  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty("walletablesData", JSON.stringify(walletablesData));
  Logger.log("Saved walletables data: " + JSON.stringify(walletablesData));
}

// テスト用関数
function testManageWalletables() {
  try {
    manageWalletables();
    Logger.log("口座管理処理が正常に完了しました。");
  } catch (e) {
    Logger.log("口座管理処理中にエラーが発生しました: " + e.message);
  }
}
