function get() {
  // getSelectedCompanyId関数を呼び出して事業所IDを取得
  var selectedCompanyId = getSelectedCompanyId();

  var freeeApp = getService();
  var accessToken = freeeApp.getAccessToken();
  Logger.log(accessToken);
}

function managePartners() {
  var freeeApp = getService();
  var accessToken = freeeApp.getAccessToken();
  var companyId = getSelectedCompanyId();
  var requestUrl = "https://api.freee.co.jp/api/1/partners?company_id=" + companyId;
  var headers = { "Authorization" : "Bearer " + accessToken };
  var options = { "method": "get", "headers": headers };

  // 既存の取引先を取得
  var response, existingPartners;
  try {
    response = UrlFetchApp.fetch(requestUrl, options).getContentText();
    existingPartners = JSON.parse(response).partners;
    if (existingPartners.length === 0) {
      Logger.log("既存の取引先が見つかりませんでした。");
    }
  } catch (e) {
    Logger.log("既存の取引先の取得に失敗しました: " + e.message);
    return;
  }

  var salesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("売上履歴");
  var salesData = salesSheet.getDataRange().getValues();
  var partnersMap = new Map(existingPartners.map(p => [p.name, p.id]));
  var newPartnersRegistered = false;

  // 新しい取引先を登録
  for (var i = 1; i < salesData.length; i++) {
    var partnerName = salesData[i][4];
    if (partnerName && !partnersMap.has(partnerName)) {
      try {
        createNewPartner(companyId, partnerName, accessToken);
        newPartnersRegistered = true;
        Logger.log("新しい取引先を登録しました: " + partnerName);
      } catch (e) {
        Logger.log("新しい取引先の登録に失敗しました: " + partnerName + "; Error: " + e.message);
      }
    }
  }

  if (!newPartnersRegistered) {
    Logger.log("新しい取引先は見つかりませんでした。");
  }

  // 更新された取引先一覧を取得
  try {
    response = UrlFetchApp.fetch(requestUrl, options).getContentText();
    var updatedPartners = JSON.parse(response).partners;
    savePartnersData(updatedPartners);
    Logger.log("更新された取引先一覧を取得しました。");
  } catch (e) {
    Logger.log("更新された取引先一覧の取得に失敗しました: " + e.message);
  }
}


// 新しい取引先を登録する関数
function createNewPartner(companyId, partnerName, accessToken) {
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

  UrlFetchApp.fetch(requestUrl, options);
}

// 新しい取引先をログに出力する関数
function logNewPartners(registeredPartners) {
  if (registeredPartners.length > 0) {
    Logger.log(registeredPartners.length + "件を登録しました: " + registeredPartners.join(", "));
  } else {
    Logger.log("登録していない取引先はありません");
  }
}
