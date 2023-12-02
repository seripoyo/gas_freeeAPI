// Freee APIから既存の品目を取得
var freeeApp = getService();
var accessToken = freeeApp.getAccessToken();
var companyId = getSelectedCompanyId();
var headers = { "Authorization": "Bearer " + accessToken };


/******************************************************************
function name |get_PartnersAndRegister
内容：取引先一覧を取得→合致しない情報を新規を登録
******************************************************************/
function get_PartnersAndRegister() {
  var requestUrl = "https://api.freee.co.jp/api/1/partners?company_id=" + companyId;
  var options = { "method": "get", "headers": headers };

  // 既存の取引先を取得
  var response = UrlFetchApp.fetch(requestUrl, options).getContentText();
  var existingPartners = JSON.parse(response).partners;
  var partnersMap = new Map(existingPartners.map(p => [p.name, p.id]));

  var salesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("売上履歴");
  var salesData = salesSheet.getDataRange().getValues();
  var newPartnersRegistered = false;

  // 新しい取引先を登録
  for (var i = 1; i < salesData.length; i++) {
    var partnerName = salesData[i][4];
    if (partnerName && !partnersMap.has(partnerName)) {
      try {
        var newPartnerId = create_NewPartner(companyId, partnerName, accessToken);
        partnersMap.set(partnerName, newPartnerId);
        newPartnersRegistered = true;
        Logger.log("新しい取引先を登録しました: " + partnerName);
      } catch (e) {
        Logger.log("既に登録されています: " + partnerName);
      }
    }
  }

  // 更新された取引先一覧を取得
  response = UrlFetchApp.fetch(requestUrl, options).getContentText();
  var updatedPartners = JSON.parse(response).partners;
  savePartnersData(updatedPartners);

  if (!newPartnersRegistered) {
    Logger.log("新しい取引先は見つかりませんでした。");
  } else {
    Logger.log("更新された取引先一覧を取得しました。");
  }
}

function create_NewPartner(companyId, partnerName, accessToken) {
  var requestUrl = "https://api.freee.co.jp/api/1/partners";
  var requestBody = {
    "company_id": companyId,
    "name": partnerName,
    "available": true,
    "shortcut1": "",
    "shortcut2": "",
    "org_code": 1,
    "country_code": "JP"
  };

  var options = {
    "method": "post",
    "headers": { "Authorization" : "Bearer " + accessToken },
    "contentType": "application/json",
    "payload": JSON.stringify(requestBody)
  };

  var response = UrlFetchApp.fetch(requestUrl, options);
  var newPartner = JSON.parse(response.getContentText()).partner;
  return newPartner.id; // 新しく登録された取引先のIDを返す
}

function save_PartnersData(partners) {
  // partnersが配列であることを確認
  if (!Array.isArray(partners)) {
    Logger.log("エラー: partnersは配列ではありません。");
    return;
  }

  // partnersが空でないことを確認
  if (partners.length === 0) {
    Logger.log("エラー: partners配列が空です。");
    return;
  }

  // partnersデータを処理
  var partnersData = partners.map(function(partner) {
    // 各partnerが必要なプロパティを持っていることを確認
    if (!partner.id || !partner.name) {
      Logger.log("エラー: 不完全なpartnerデータ - " + JSON.stringify(partner));
      return; // 不完全なデータは無視
    }
    return { id: partner.id, name: partner.name };
  });

  // フィルタリングされたデータのみを保存
  var filteredPartnersData = partnersData.filter(function(data) {
    return data !== undefined;
  });

  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty("partnersData", JSON.stringify(filteredPartnersData));

  // 保存した取引先情報をログに出力
  Logger.log("保存した取引先情報: " + JSON.stringify(filteredPartnersData));
}

