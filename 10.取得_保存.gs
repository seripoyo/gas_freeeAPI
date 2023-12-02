/******************************************************************
 * 各関数の共通変数
******************************************************************/


/******************************************************************
function name |get_Walletables
summary       |口座一覧取得
requestUrl    |https://api.freee.co.jp/api/1/walletables?company_id={companyId}&with_balance=true
method        |GET
******************************************************************/
function get_Walletables() {
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
  var processedWalletables = walletables.map(function (walletable) {
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
処理：get_Walletablesで保存した関数の取得
******************************************************************/
function saved_Walletables() {
  var userProperties = PropertiesService.getUserProperties();
  var walletablesDataString = userProperties.getProperty("walletablesData");

  if (walletablesDataString) {
    var walletablesData = JSON.parse(walletablesDataString);

    var processedWalletables = walletablesData.map(function(walletable) {
      // IDを整数に変換
      var id = String(walletable.id).replace('.0', '');
      // walletable_balanceとlast_balanceを文字列に変換して末尾の".0"を取り除く
      var walletable_balance = String(walletable.walletable_balance).replace('.0', '');
      var last_balance = String(walletable.last_balance).replace('.0', '');

      return {
        ...walletable,
        id: id,
        walletable_balance: walletable_balance,
        last_balance: last_balance
      };
    });

    Logger.log(processedWalletables);
    return processedWalletables;

  } else {
    Logger.log("No walletables data found.");
    return [];
  }
}


/******************************************************************
function name |getTaxes
summary       |税区分一覧取得&保存
requestUrl   |https://api.freee.co.jp/api/1/taxes/companies/
method        |GET
******************************************************************/
function get_Taxes() {
  var freeeApp = getService();
  var accessToken = freeeApp.getAccessToken();
  var companyId = getSelectedCompanyId();
  var requestUrl = "https://api.freee.co.jp/api/1/taxes/companies/" + companyId;
  var headers = { "Authorization": "Bearer " + accessToken };
  var options = { "method": "get", "headers": headers };

  var response = UrlFetchApp.fetch(requestUrl, options).getContentText();
  var taxesResponse = JSON.parse(response);
  var taxes = taxesResponse.taxes;

  // 税区分一覧のIDを整数に変換して配列に格納
  var taxesData = taxes.map(function (tax) {
    return {
      id: parseInt(tax.code, 10).toString(), // IDを整数に変換して文字列化
      name_ja: tax.name_ja
    };
  });

  Logger.log(taxesData);

  // 税区分データをユーザープロパティに保存
  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty("taxesData", JSON.stringify(taxesData));
}


/******************************************************************
function name |get_AccountItems
summary       |勘定科目一覧取得
requestUrl    |https://api.freee.co.jp/api/1/account_items?company_id={companyId}
method        |GET
******************************************************************/
function get_AccountItems() {
  var freeeApp = getService();
  var accessToken = freeeApp.getAccessToken();
  var companyId = getSelectedCompanyId();
  var requestUrl = "https://api.freee.co.jp/api/1/account_items?company_id=" + companyId;
  var headers = { "Authorization": "Bearer " + accessToken };
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

  fColumnData.forEach(function (fValue) {
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

/******************************************************************
function name |getPartners
summary       |取引先一覧取得
requestUrl   |https://api.freee.co.jp/api/1/partners?company_id={companyId}
method        |GET
******************************************************************/
function get_Partners() {
  var freeeApp = getService();
  var accessToken = freeeApp.getAccessToken();
  var companyId = getSelectedCompanyId();
  var requestUrl = "https://api.freee.co.jp/api/1/partners?company_id=" + companyId;
  var headers = { "Authorization": "Bearer " + accessToken };
  var options = { "method": "get", "headers": headers };

  var response = UrlFetchApp.fetch(requestUrl, options).getContentText();
  var partnersResponse = JSON.parse(response);
  var partners = partnersResponse.partners;

  // 配列を作成し、要素を格納（IDを整数に変換）
  var partnersData = partners.map(function (partner) {
    return {
      id: parseInt(partner.id, 10).toString(), // IDを整数に変換して文字列化
      name: partner.name
    };
  });

  // 取引先データをユーザープロパティに保存
  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty("partnersData", JSON.stringify(partnersData));
}


/******************************************************************
 * 取引先データの呼び出し関数 
 ******************************************************************/

function saved_PartnersData() {
  var userProperties = PropertiesService.getUserProperties();
  var partnersDataString = userProperties.getProperty("partnersData");

  if (partnersDataString) {
    var partnersData = JSON.parse(partnersDataString);
    Logger.log(partnersData);
    return partnersData;
  } else {
    Logger.log("No partners data found.");
    return []; // データがない場合は空の配列を返す
  }
}


/******************************************************************
function name |get_Items
summary       |品目一覧取得
requestUrl    |https://api.freee.co.jp/api/1/items?company_id={companyId}
method        |GET
******************************************************************/
function get_Items() {
  var freeeApp = getService();
  var accessToken = freeeApp.getAccessToken();
  var companyId = getSelectedCompanyId();
  var requestUrl = "https://api.freee.co.jp/api/1/items?company_id=" + companyId + "&limit=3000";
  var headers = { "Authorization": "Bearer " + accessToken };
  var options = { "method": "get", "headers": headers };

  var response = UrlFetchApp.fetch(requestUrl, options).getContentText();
  var itemsResponse = JSON.parse(response);
  var items = itemsResponse.items;

  // 配列を作成し、要素を格納（IDを整数に変換）
  var itemsData = items.map(function (item) {
    return {
      id: parseInt(item.id, 10).toString(), // IDを整数に変換して文字列化
      name: item.name
    };
  });

  // 品目データをユーザープロパティに保存
  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty("itemsData", JSON.stringify(itemsData));
}


function getSavedItemsData() {
  var userProperties = PropertiesService.getUserProperties();
  var itemsDataString = userProperties.getProperty("itemsData");

  if (itemsDataString) {
    var itemsData = JSON.parse(itemsDataString);
    Logger.log(itemsData);
    return itemsData;
  } else {
    Logger.log("No items data found.");
    return []; // データがない場合は空の配列を返す
  }
}
