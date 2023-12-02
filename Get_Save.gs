
/******************************************************************
function name |getWalletables
summary       |口座一覧取得
requestUrl    |https://api.freee.co.jp/api/1/walletables?company_id={companyId}&with_balance=true
method        |GET
******************************************************************/
function getWalletables() {
  var freeeApp = getService();
  var accessToken = freeeApp.getAccessToken();
  var companyId = getSelectedCompanyId();
  var requestUrl = "https://api.freee.co.jp/api/1/walletables?company_id=" + companyId + "&with_balance=true";
  var headers = { "Authorization": "Bearer " + accessToken };

  var options = {
    "method": "get",
    "headers": headers
  };

  var res = UrlFetchApp.fetch(requestUrl, options).getContentText();
  var walletablesResponse = JSON.parse(res);
  var walletables = walletablesResponse.walletables;

  // IDを整数に変換して保存
  var processedWalletables = walletables.map(function(walletable) {
    return {
      id: parseInt(walletable.id), // IDを整数に変換
      name: walletable.name,
      type: walletable.type,
      bank_id: walletable.bank_id,
      walletable_balance: walletable.walletable_balance,
      last_balance: walletable.last_balance
    };
  });

  // 変換したデータをプロパティサービスに保存
  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty("walletablesData", JSON.stringify(processedWalletables));
}

/******************************************************************
処理：getWalletablesで保存した関数の取得
******************************************************************/
function getSavedWalletablesData() {
  var userProperties = PropertiesService.getUserProperties();
  var walletablesDataString = userProperties.getProperty("walletablesData");
  
  if (walletablesDataString) {
    var walletablesData = JSON.parse(walletablesDataString);

    // 数値が含まれるプロパティを整数に変換
    var processedWalletables = walletablesData.map(function(walletable) {
      return {
        ...walletable, // その他のプロパティを維持
        id: parseInt(walletable.id, 10), // IDを整数に変換
        walletable_balance: parseInt(walletable.walletable_balance, 10), // walletable_balanceを整数に変換
        last_balance: parseInt(walletable.last_balance, 10) // last_balanceを整数に変換
      };
    });

    Logger.log(processedWalletables);
    return processedWalletables; // 変換後のデータを返す

  } else {
    Logger.log("No walletables data found.");
    return []; // データがない場合は空の配列を返す
  }
}

/******************************************************************
function name |getTaxes
summary       |税区分一覧取得&保存
requestUrl   |https://api.freee.co.jp/api/1/taxes/companies/
method        |GET
******************************************************************/
function getAndSaveMatchingTaxes() {
  var freeeApp = getService();
  var accessToken = freeeApp.getAccessToken();
  var companyId = getSelectedCompanyId();
  var requestUrl = "https://api.freee.co.jp/api/1/taxes/companies/" + companyId;
  var headers = { "Authorization": "Bearer " + accessToken };
  var options = { "method": "get", "headers": headers };

  var response = UrlFetchApp.fetch(requestUrl, options).getContentText();
  var taxesResponse = JSON.parse(response);
  var taxes = taxesResponse.taxes;

  // 税区分一覧のID（整数に変換）と名前を配列に格納
  var taxesData = taxes.map(function (tax) {
    return { id: parseInt(tax.code, 10), name_ja: tax.name_ja };
  });

  Logger.log(taxesData);
  // スプレッドシートのG列のデータを取得
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("売上履歴");
  var lastRowInGColumn = getLastRowInColumn("売上履歴", 7);
  var gColumnData = sheet.getRange(2, 7, lastRowInGColumn).getValues();
  Logger.log(gColumnData);
  // 合致する税区分を検索し、保存する配列
  var matchingTaxes = [];

  gColumnData.forEach(function (gValue) {
    var matchingTax = taxesData.find(tax => tax.name_ja === gValue[0]);
    if (matchingTax) {
      matchingTaxes.push(matchingTax);
    }
  });
  Logger.log(matchingTaxes);
  saveMatchingTaxesData(matchingTaxes);
}


function getAndSaveMatchingTaxes() {
  var freeeApp = getService();
  var accessToken = freeeApp.getAccessToken();
  var companyId = 2146764;
  var requestUrl = "https://api.freee.co.jp/api/1/taxes/companies/" + companyId;
  var headers = { "Authorization": "Bearer " + accessToken };
  var options = { "method": "get", "headers": headers };

  var response = UrlFetchApp.fetch(requestUrl, options).getContentText();
  var taxesResponse = JSON.parse(response);
  var taxes = taxesResponse.taxes;

  // 税区分一覧のID（整数に変換）と名前を配列に格納
  var taxesData = taxes.map(function (tax) {
    return { id: parseInt(tax.code, 10), name_ja: tax.name_ja };
  });

  Logger.log(taxesData);
}
