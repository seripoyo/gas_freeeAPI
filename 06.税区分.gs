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


