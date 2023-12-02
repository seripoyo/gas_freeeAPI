



/******************************************************************
function name |getAccountItems
summary       |勘定科目一覧取得
requestUrl    |https://api.freee.co.jp/api/1/account_items?company_id={companyId}
method        |GET
******************************************************************/
function getAccountItems() {
  var freeeApp = getService();
  var accessToken = freeeApp.getAccessToken();
  var requestUrl = "https://api.freee.co.jp/api/1/account_items?company_id={companyId}";
    var companyId = getSelectedCompanyId();
  var requestUrl = requestUrl.replace("{companyId}", companyId);
  var headers = { "Authorization": "Bearer " + accessToken };

  var options =
  {
    "method": "get",
    "headers": headers
  };
  var res = UrlFetchApp.fetch(requestUrl, options).getContentText();

  //勘定科目一覧を取得
  var accountItemsResponse = JSON.parse(res);
  var accountItems = accountItemsResponse.account_items;

  //項目ごとに配列を作成し、要素を格納
  var id = [];
  var name = [];
  var defaultTaxId = [];

  for (var i = 0; i < accountItems.length; i++) {
    id.push([accountItems[i].id]);
    name.push([accountItems[i].name]);
        defaultTaxCode.push([accountItems[i].default_tax_code]);//もし税区分が入力されていなかった場合はdefaultTaxIdを
  };

  //取得した勘定科目一覧を”勘定科目一覧”シートに反映
  var sheet = ss.getSheetByName("勘定科目一覧");
  sheet.clear();
  sheet.getRange(1, 1, id.length, 1).setValues(id);
  sheet.getRange(1, 2, name.length, 1).setValues(name);
  sheet.getRange(1, 3, defaultTaxCode.length, 1).setValues(defaultTaxCode);
}

/******************************************************************
function name |getPartners
summary       |取引先一覧取得
requestUrl   |https://api.freee.co.jp/api/1/partners?company_id={companyId}
method        |GET
******************************************************************/
function getPartners() {
  var freeeApp = getService();
  var accessToken = freeeApp.getAccessToken();
  var requestUrl = "https://api.freee.co.jp/api/1/partners?company_id={companyId}";
    var companyId = getSelectedCompanyId();
  var requestUrl = requestUrl.replace("{companyId}", companyId);
  var headers = { "Authorization": "Bearer " + accessToken };

  var options =
  {
    "method": "get",
    "headers": headers
  };

  var res = UrlFetchApp.fetch(requestUrl, options).getContentText();

  //取引先一覧を取得
  var partnersResponse = JSON.parse(res);
  var partners = partnersResponse.partners;

  //項目ごとに配列を作成し、要素を格納
  var id = [];
  var name = [];


  for (var i = 0; i < partners.length; i++) {
    id.push([partners[i].id]);
    name.push([partners[i].name]);
  };

  //取得した取引先一覧を”取引先一覧”シートに反映
  var sheet = ss.getSheetByName("取引先一覧");
  sheet.clear();

  try {
    sheet.getRange(1, 1, id.length, 1).setValues(id);
    sheet.getRange(1, 2, name.length, 1).setValues(name);
  } catch (e) {
    return; //取引先がなかった時の処理
  };

}


/******************************************************************
function name |getItems
summary       |品目一覧取得
requestUrl    |https://api.freee.co.jp/api/1/items?company_id={companyId}
method        |GET
******************************************************************/
function getItems() {
  var freeeApp = getService();
  var accessToken = freeeApp.getAccessToken();
  var requestUrl = "https://api.freee.co.jp/api/1/items?company_id={companyId}&limit=3000";
    var companyId = getSelectedCompanyId();
  var requestUrl = requestUrl.replace("{companyId}", companyId);
  var headers = { "Authorization": "Bearer " + accessToken };

  var options =
  {
    "method": "get",
    "headers": headers
  };
  var res = UrlFetchApp.fetch(requestUrl, options).getContentText();

  //品目一覧を取得
  var itemsResponse = JSON.parse(res);
  var items = itemsResponse.items;

  //項目ごとに配列を作成し、要素を格納
  var id = [];
  var name = [];

  for (var i = 0; i < items.length; i++) {
    id.push([items[i].id]);
    name.push([items[i].name]);
  };

  //取得した品目一覧を”品目一覧”シートに反映
  var sheet = ss.getSheetByName("品目一覧");
  sheet.clear();

  try {
    sheet.getRange(1, 1, id.length, 1).setValues(id);
    sheet.getRange(1, 2, name.length, 1).setValues(name);

  } catch (e) {
    return; //品目がなかった時の処理
  };

}

