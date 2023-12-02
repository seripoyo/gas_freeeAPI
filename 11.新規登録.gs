function manage_Partners() {
  var freeeApp = getService();
  var accessToken = freeeApp.getAccessToken();
  var companyId = getSelectedCompanyId();
  var requestUrl = "https://api.freee.co.jp/api/1/partners?company_id=" + companyId + "&limit=3000";
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

  var newPartners = [];

  // 新しい取引先を登録
  for (var i = 1; i < salesData.length; i++) {
    var partnerName = salesData[i][4];
    if (partnerName && !partnersMap.has(partnerName)) {
      try {
        var newPartnerId = createNewPartner(companyId, partnerName, accessToken);
        if (newPartnerId) {  // 登録成功時のみログに出力
          newPartners.push({ name: partnerName, id: newPartnerId });
          Logger.log("登録しました: " + partnerName + " (ID: " + newPartnerId + ")");
        } else {
          Logger.log("登録に失敗しました: " + partnerName);
        }
      } catch (e) {
        Logger.log("エラー発生: " + e.message);
      }
    }
  }

  // 新しい取引先が登録されなかった場合
  if (newPartners.length === 0) {
    Logger.log("新しい取引先は見つかりませんでした。");
  } else {
    logNewPartners(newPartners);
  }

  // 更新された取引先一覧を取得して保存
  try {
    response = UrlFetchApp.fetch(requestUrl, options).getContentText();
    var updatedPartners = JSON.parse(response).partners;
    Logger.log("更新された取引先一覧を取得しました。");
  } catch (e) {
    Logger.log("更新された取引先一覧の取得に失敗しました: " + e.message);
  }

  // 取得した一覧を savePartnersData に渡す
  savePartnersData(updatedPartners);
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

  var response = UrlFetchApp.fetch(requestUrl, options).getContentText();
  var result = JSON.parse(response);

  // 正常に登録された場合はIDを返す
  return result.partner ? result.partner.id : null;
}
function logNewPartners(registeredPartners) {
  registeredPartners.forEach(p => {
    Logger.log("登録済み取引先: " + p.name + " (ID: " + p.id + ")");
  });
}

function savePartnersData(partners) {
  // 取引先のIDと名前をログに出力
  partners.forEach(partner => {
    Logger.log("取引先: " + partner.name + " (ID: " + partner.id + ")");
  });

  // Google スプレッドシートに保存
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("取引先一覧") || spreadsheet.insertSheet("取引先一覧");
  sheet.clear(); // 既存のデータをクリア

  // ヘッダー行を追加
  sheet.appendRow(["ID", "取引先名"]);

  // 各取引先データを行として追加
  partners.forEach(partner => {
    sheet.appendRow([partner.id, partner.name]);
  });
}
