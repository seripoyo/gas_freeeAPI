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

