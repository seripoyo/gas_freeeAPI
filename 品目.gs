/******************************************************************
 * 
 * 品目の一覧を取得しスプシのL列に存在する項目が合致しなかったら
 * 新しく登録した上で一覧を再取得して名称とIDを保存するよ
 * 
******************************************************************/


/******************************************************************
関数：get_Items_Register
概要：品目一覧取得&登録→再度一覧取得→データを保存
******************************************************************/
// 既存の品目の一覧を取得し、必要に応じて新しい品目を登録
function get_Items_Register() {
  var freeeApp = getService();
  var accessToken = freeeApp.getAccessToken();
  var companyId = getSelectedCompanyId();
  var requestUrl = "https://api.freee.co.jp/api/1/items?company_id=" + companyId + "&limit=3000";
  var headers = { "Authorization": "Bearer " + accessToken };
  var options = { "method": "get", "headers": headers };

  // 既存の品目一覧を取得
  var response = UrlFetchApp.fetch(requestUrl, options).getContentText();
  var itemsResponse = JSON.parse(response);
  var items = itemsResponse.items;
  var itemsMap = new Map(items.map(item => [item.name.trim(), item.id]));

  // スプレッドシートからL列のデータを取得
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("取引一覧");
  var lastRowInLColumn = getLastRowInColumn("取引一覧", 12);
  var lColumnData = sheet.getRange(2, 12, lastRowInLColumn - 1).getValues();

  var newlyRegisteredItems = []; // 新しく登録された品目を格納する配列

  lColumnData.forEach(function (lValue) {
    var itemName = lValue[0].trim();
    if (itemName && !itemsMap.has(itemName)) {
      // 新しい品目を登録
      try {
        var newItemId = register_Save_New_Item(companyId, itemName, accessToken);
        newlyRegisteredItems.push({ id: newItemId, name: itemName });
        Logger.log("新しい品目を登録しました: " + itemName);
      } catch (e) {
        Logger.log("新しい品目の登録に失敗しました: " + itemName);
      }
    }
  });

  // 新しい品目を含む品目一覧を更新
  var updatedItems = items.concat(newlyRegisteredItems);

  // 品目データを保存
  var itemsData = updatedItems.map(function (item) {
    return { item_id: item.id.toString(), name: item.name };
  });

  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty("itemsData", JSON.stringify(itemsData));
  Logger.log("品目データは保存されています");
}

/******************************************************************
関数：register_Save_New_Item
概要：品目の新規登録
******************************************************************/
function register_Save_New_Item(companyId, itemName, accessToken) {
  var requestUrl = "https://api.freee.co.jp/api/1/items";
  var headers = { "Authorization": "Bearer " + accessToken };
  var options;

  // 新しい品目を登録
  var requestBody = { "company_id": companyId, "name": itemName };
  options = {
    "method": "post",
    "headers": headers,
    "contentType": "application/json",
    "payload": JSON.stringify(requestBody),
    "muteHttpExceptions": true
  };

  var response = UrlFetchApp.fetch(requestUrl, options);
  var responseCode = response.getResponseCode();

  if (responseCode >= 200 && responseCode < 300) {
    // 品目登録に成功した場合、更新された品目一覧を取得
    options = { "method": "get", "headers": headers };
    response = UrlFetchApp.fetch(requestUrl + "?company_id=" + companyId, options).getContentText();
    var updatedItems = JSON.parse(response).items;

    // 品目データを保存
    var itemsData = updatedItems.map(function (item) {
      return { item_id: item.id.toString(), name: item.name };
    });

    var userProperties = PropertiesService.getUserProperties();
    userProperties.setProperty("itemsData", JSON.stringify(itemsData));
    Logger.log("新しい品目を含む保存された品目データ: " + JSON.stringify(itemsData));
  } else {
    // 品目登録に失敗した場合
    Logger.log("品目登録エラー: " + response.getContentText());
    throw new Error("品目の登録に失敗しました: " + itemName);
  }
}



// 最終行を取得する関数
function getLastRowInColumn(sheetName, columnNumber) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var data = sheet.getRange(2, columnNumber, sheet.getLastRow()).getValues();
  var lastRow = data.filter(String).length;
  return lastRow;
}

// テスト用の関数
function testGetItemsAndRegister() {
  get_Items_Register();
  var itemsData = getSavedItemsData();
  Logger.log("取得した品目データ: " + JSON.stringify(itemsData));
}

// 保存した品目データを取得する関数
function getSavedItemsData() {
  var userProperties = PropertiesService.getUserProperties();
  var itemsDataString = userProperties.getProperty("itemsData");
  return itemsDataString ? JSON.parse(itemsDataString) : [];
  
}
