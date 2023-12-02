/******************************************************************
function name |get_Items_Register
summary       |品目一覧取得&登録→再度一覧取得→データを保存
******************************************************************/
// 既存の品目の一覧を取得し、必要に応じて新しい品目を登録
function get_Items_Register() {
  var companyId = getSelectedCompanyId();
  var accessToken = getService().getAccessToken();
  var requestUrl = "https://api.freee.co.jp/api/1/items?company_id=" + companyId + "&limit=3000";
  var headers = { "Authorization": "Bearer " + accessToken };
  var options = { "method": "get", "headers": headers };

  // 既存の品目一覧を取得
  var response = UrlFetchApp.fetch(requestUrl, options).getContentText();
  var itemsResponse = JSON.parse(response);
  var items = itemsResponse.items;
  var itemsMap = new Map(items.map(item => [item.name.trim(), item.id]));

  // スプレッドシートからL列のデータを取得
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("売上履歴");
  var lastRowInLColumn = getLastRowInColumn("売上履歴", 12);
  var lColumnData = sheet.getRange(2, 12, lastRowInLColumn - 1).getValues();

  var newlyRegisteredItems = []; // 新しく登録された品目を格納する配列

  lColumnData.forEach(function(lValue) {
    var itemName = lValue[0].trim();
    if (itemName && !itemsMap.has(itemName)) {
      // 新しい品目を登録
      try {
        var newItemId = register_NewItem(companyId, itemName, accessToken);
        newlyRegisteredItems.push({ id: newItemId, name: itemName });
        Logger.log("新しい品目を登録しました: " + itemName);
      } catch (e) {
        Logger.log("新しい品目の登録に失敗しました: " + itemName);
      }
    }
  });
/******************************************************************
function name |register_NewItem
summary       |品目の新規登録
******************************************************************/
function register_NewItem(companyId, itemName, accessToken) {
  var requestUrl = "https://api.freee.co.jp/api/1/items";
  var requestBody = {
    "company_id": companyId,
    "name": itemName
  };

  var options = {
    "method": "post",
    "headers": { "Authorization": "Bearer " + accessToken },
    "contentType": "application/json",
    "payload": JSON.stringify(requestBody),
    "muteHttpExceptions": true // HTTP例外をミュートに設定
  };

  var response = UrlFetchApp.fetch(requestUrl, options);
  var responseCode = response.getResponseCode();
  var responseContent = JSON.parse(response.getContentText());

  if (responseCode >= 200 && responseCode < 300) {
    // 正常に登録された場合
    var newItem = responseContent.item;
    return newItem.id.toString(); // 新しく登録された品目のIDを返す
  } else {
    // エラーが発生した場合、詳細をログに出力
    Logger.log("品目登録エラー: " + response.getContentText());
    throw new Error("品目の登録に失敗しました: " + itemName);
  }
}
  // 更新された品目一覧を取得
  response = UrlFetchApp.fetch(requestUrl, options).getContentText();
  var updatedItems = JSON.parse(response).items;
  saveItemsData(updatedItems);

  // 新しく登録された品目のリストをログに出力
  if (newlyRegisteredItems.length > 0) {
    Logger.log("新しく登録された品目: " + JSON.stringify(newlyRegisteredItems));
  } else {
    Logger.log("新しい品目は登録されませんでした。");
  }
}
// 品目のデータを保存する関数
function saveItemsData(items) {
  var itemsData = items.map(function(item) {
    return { id: item.id.toString(), name: item.name };
  });

  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty("itemsData", JSON.stringify(itemsData));
  Logger.log("保存した品目データ: " + JSON.stringify(itemsData));
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
  get_ItemsAndRegister();
  var itemsData = getSavedItemsData();
  Logger.log("取得した品目データ: " + JSON.stringify(itemsData));
}

// 保存した品目データを取得する関数
function getSavedItemsData() {
  var userProperties = PropertiesService.getUserProperties();
  var itemsDataString = userProperties.getProperty("itemsData");
  return itemsDataString ? JSON.parse(itemsDataString) : [];
}