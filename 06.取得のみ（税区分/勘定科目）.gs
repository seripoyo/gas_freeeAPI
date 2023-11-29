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
  var taxesData = taxes.map(function(tax) {
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

  gColumnData.forEach(function(gValue) {
    var matchingTax = taxesData.find(tax => tax.name_ja === gValue[0]);
    if (matchingTax) {
      matchingTaxes.push(matchingTax);
    }
  });

 saveMatchingTaxesData(matchingTaxes);
}

/******************************************************************
function name |getAccountItems
summary       |勘定科目一覧取得
requestUrl    |https://api.freee.co.jp/api/1/account_items?company_id={companyId}
method        |GET
******************************************************************/
function getAndSaveMatchingAccountItems() {
  var freeeApp = getService();
  var accessToken = freeeApp.getAccessToken();
  var companyId = getSelectedCompanyId();
  var requestUrl = "https://api.freee.co.jp/api/1/account_items?company_id=" + companyId;
  var headers = { "Authorization" : "Bearer " + accessToken };
  var options = { "method": "get", "headers": headers };

  var response = UrlFetchApp.fetch(requestUrl, options).getContentText();
  var accountItemsResponse = JSON.parse(response);
  var accountItems = accountItemsResponse.account_items;

  // スプレッドシートのF列のデータを取得
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("売上履歴");
  var lastRowInFColumn = getLastRowInColumn("売上履歴", 6);
  var fColumnData = sheet.getRange(2, 6, lastRowInFColumn - 1).getValues();

  // 合致する勘定科目のnameとidを保存する配列
  var matchingAccountItems = [];

  fColumnData.forEach(function(fValue) {
    var matchingAccountItem = accountItems.find(accountItem => accountItem.name === fValue[0]);
    if (matchingAccountItem) {
      matchingAccountItems.push({
        id: matchingAccountItem.id,
        name: matchingAccountItem.name
      });
    }
  });

  // 合致する勘定科目をプロパティサービスに保存
  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty("matchingAccountItems", JSON.stringify(matchingAccountItems));

  // 確認のためにログに出力
  Logger.log("保存した勘定科目: " + JSON.stringify(matchingAccountItems));
}

// 指定したシートと列番号に基づいて情報が入力されている最終行を取得する関数
function getLastRowInColumn(sheetName, columnNumber) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var columnData = sheet.getRange(1, columnNumber, sheet.getMaxRows(), 1).getValues();
  var lastRow = columnData.filter(String).length;
  return lastRow;
}

