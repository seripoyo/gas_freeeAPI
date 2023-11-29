
/******************************************************************
内容：税区分一覧取得→既存を保存→新規を登録
使用API：https://api.freee.co.jp/api/1/taxes/companies/
実行内容：GET・POST
******************************************************************/

function getItemsAndRegisterNew() {
  // Freee APIから既存の品目を取得
  var freeeApp = getService();
  var accessToken = freeeApp.getAccessToken();
  var companyId = getSelectedCompanyId();
  var requestUrl = "https://api.freee.co.jp/api/1/items?company_id=" + companyId + "&limit=3000";
  var headers = { "Authorization": "Bearer " + accessToken };
  var options = { "method": "get", "headers": headers };

  var response = UrlFetchApp.fetch(requestUrl, options).getContentText();
  var itemsResponse = JSON.parse(response);
  var items = itemsResponse.items;

  // 既存の品目の名前とIDをマッピング
  var itemsMap = new Map(items.map(item => [item.name.trim(), item.id]));

  // スプレッドシートからL列のデータを取得
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("売上履歴");
  var lastRowInLColumn = getLastRowInColumn("売上履歴", 12);
  var lColumnData = sheet.getRange(2, 12, lastRowInLColumn - 1).getValues();

  var itemsData = [];

  lColumnData.forEach(function(lValue) {
    var itemName = lValue[0].trim();
    if (itemName) {
      var itemId = getExistingItemId(itemName, itemsMap);
      
      if (itemId === null) {
        // 新しい品目を登録
        try {
          itemId = registerNewItem(companyId, itemName, accessToken);
        } catch (e) {
          Logger.log("新しい品目の登録に失敗しました: " + itemName + " - " + e.message);
        }
      }

      // 品目データを配列に追加
      if (itemId !== null) {
        itemsData.push({ id: itemId.toString(), name: itemName });
      } else {
        Logger.log("品目IDが見つからない: " + itemName);
      }
    }
  });

  saveItemsData(itemsData);
}

// =============================================================-
// 内容：新規登録された品目のIDを取得しIDを整数・文字列に変換
// 使用API：https://api.freee.co.jp/api/1/taxes/companies/
// 実行内容：GET・POST
// -=============================================================

// 新規登録された品目のIDを取得し、そのIDを文字列として返します。
function registerNewItem(companyId, itemName, accessToken) {
  // 品目名の前後の空白を取り除く
  var trimmedItemName = itemName.trim();
  // 新しい品目をFreee APIに登録
  var requestUrl = "https://api.freee.co.jp/api/1/items";
  var requestBody = {
    "company_id": companyId,
    "name": trimmedItemName
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
    throw new Error(responseContent.errors.map(error => error.messages).join("; "));
  }
  var newItem = responseContent.item;
  return newItem.id.toString();
}

// =============================================================-
// 内容：既存品目のIDを取得
// 使用API：https://api.freee.co.jp/api/1/taxes/companies/
// 実行内容：GET・POST
// -=============================================================
function getExistingItemId(itemName, itemsMap) {
  if (itemsMap.has(itemName)) {
    return itemsMap.get(itemName);
  }
  return null; // 品目が見つからなかった場合はnullを返す
}

