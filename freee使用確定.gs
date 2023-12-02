
/******************************************************************
function name |getItems
summary       |品目一覧取得
requestUrl    |https://api.freee.co.jp/api/1/items?company_id={companyId}
method        |GET
******************************************************************/
function getItems() {
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
  var itemsData = items.map(function(item) {
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
