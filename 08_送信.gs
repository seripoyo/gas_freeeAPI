/******************************************************************
 * 関数：postDeals
 * 概要：取引の送信
 * request_url   |https://api.freee.co.jp/api/1/deals
 * method        |POST
 ******************************************************************/
function postDeals() {
  // Freee APIへのアクセス情報を取得
  var freeeApp = getService();
  var accessToken = freeeApp.getAccessToken();

  // APIエンドポイントURL
  var requestUrl = "https://api.freee.co.jp/api/1/deals";

  // スプレッドシートの情報を取得
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var headers = { "Authorization": "Bearer " + accessToken };

  // スプレッドシートから取引データを取得
  var dealsSheet = ss.getSheetByName("取引");
  var dealsColumnLastRow = getLastRowNumber(2, "取引");
  var detailsColumnLastRow = getLastRowNumber(7, "取引");
  var paymentsColumnLastRow = getLastRowNumber(14, "取引");
  var dealsLastRow = Math.max(dealsColumnLastRow, detailsColumnLastRow, paymentsColumnLastRow);
  var dealsValues = dealsSheet.getRange(1, 1, dealsLastRow, 17).getValues();

  // 各種カウンターおよびデータ格納用の配列を初期化
  var countDealsRowSkip = 0; // 行スキップのカウンター
  var countPostedDeals = 0; // 送信成功取引のカウンター
  var countErrorDeals = 0; // 送信エラー取引のカウンター
  var details = []; // 取引詳細データの格納用配列
  var payments = []; // 取引支払データの格納用配列

  // 取引データをループ処理
  for (var i = 1; i < dealsLastRow; i++) {
    // 取引詳細データおよび支払データの行を取得
    var detailsRow = String(dealsValues[i][6]);
    var paymentsRow = String(dealsValues[i][13]);

    // 最終行の場合、次の行は「取引を作成する」
    if (i + 1 == dealsLastRow) {
      var nextdealsRow = "取引を作成する";
    } else {
      var nextdealsRow = String(dealsValues[i + 1][1]);
    }

    // 取引詳細データを作成
    if (detailsRow != "") {
      // 取引詳細データの各フィールドを取得
      var accountItemId = parseInt(dealsValues[i][6]);
      var taxCode = parseInt(dealsValues[i][7]);
      var itemId = parseInt( dealsValues[ i ][ 8 ] );
      var vat = parseInt(dealsValues[i][12]);
      var sectionId = parseInt(dealsValues[i][9]);
      var tagIds = String(dealsValues[i][10]).split(",");
      
      // 空の場合、空の配列にする
      if (tagIds == "") {
        tagIds = [];
      }

      var amountDetails = parseInt(dealsValues[i][11]);

      // 取引詳細データを配列に追加
      details.push({
        "account_item_id": accountItemId,
        "tax_code": taxCode,
        "item_id": isNaN(itemId) ? undefined : itemId,
        "section_id": isNaN(sectionId) ? undefined : sectionId,
        "tag_ids": tagIds,
        "amount": amountDetails,
        "vat": vat
      });
    }

    // 取引支払データを作成
    if (paymentsRow != "") {
      // 取引支払データの各フィールドを取得
      var date = Utilities.formatDate(dealsValues[i][13], "JST", "yyyy-MM-dd");
      var fromWalletableType = String(dealsValues[i][14]);
      var fromWalletableId = parseInt(dealsValues[i][15]);
      var amountPayments = parseInt(dealsValues[i][16]);

      // 取引支払データを配列に追加
      payments.push({
        "date": date,
        "from_walletable_type": fromWalletableType,
        "from_walletable_id": fromWalletableId,
        "amount": amountPayments
      });
    }

    // 取引を作成
    if (nextdealsRow == "") {
      countDealsRowSkip++;
    } else {
      var companyId = parseInt(dealsValues[i - countDealsRowSkip][0]);

      if (dealsValues[i - countDealsRowSkip][1] != "") {
        var issueDate = Utilities.formatDate(dealsValues[i - countDealsRowSkip][1], "JST", "yyyy-MM-dd");
      }

      if (dealsValues[i - countDealsRowSkip][2] != "") {
        var dueDate = Utilities.formatDate(dealsValues[i - countDealsRowSkip][2], "JST", "yyyy-MM-dd");
      }

      var type = String(dealsValues[i - countDealsRowSkip][3]);
      var partnerId = parseInt(dealsValues[i - countDealsRowSkip][4], 10);
      var refNumber = String(dealsValues[i - countDealsRowSkip][5]);

      // リクエストボディの作成
      var requestBody = {
        "company_id": companyId,
        "issue_date": issueDate,
        "due_date": dueDate,
        "type": type,
        "partner_id": isNaN(partnerId) ? undefined : partnerId,
        "ref_number": refNumber,
        "details": details,
        "payments": payments
      };

      // Logger.log(JSON.stringify(requestBody));

      // POSTオプション
      var options = {
        "method": "POST",
        "contentType": "application/json",
        "headers": headers,
        "payload": JSON.stringify(requestBody),
        "muteHttpExceptions": true
      };

      // POSTリクエスト
      var res = UrlFetchApp.fetch(requestUrl, options);

      // レスポンスコードに応じて処理を分岐
      if (res.getResponseCode() == 201) {
        countPostedDeals++;
      } else {
        Logger.log("Error response: " + res.getContentText());
        countErrorDeals++;
      }

      countDealsRowSkip = 0;
      details.length = 0;
      payments.length = 0;
    }
  }

  // 結果のアラート表示
  SpreadsheetApp.getUi().alert(countPostedDeals + "件の取引を送信しました");
  if (countErrorDeals != 0) {
    SpreadsheetApp.getUi().alert(countErrorDeals + "件の取引の送信に失敗しました");
  }
}
