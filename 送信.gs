
/******************************************************************
function name |postDeals
summary       |取引の送信
request_url   |https://api.freee.co.jp/api/1/deals
method        |POST
******************************************************************/
function postDeals() {
  var freeeApp = getService();
  var accessToken = freeeApp.getAccessToken();
  var requestUrl = "https://api.freee.co.jp/api/1/deals";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var headers = { "Authorization": "Bearer " + accessToken };
  var dealsSheet = ss.getSheetByName("取引");
  var dealsColumnLastRow = getLastRowNumber(2, "取引");
  var detailsColumnLastRow = getLastRowNumber(7, "取引");
  var paymentsColumnLastRow = getLastRowNumber(14, "取引");
  var dealsLastRow = Math.max(dealsColumnLastRow, detailsColumnLastRow, paymentsColumnLastRow);
  var dealsValues = dealsSheet.getRange(1, 1, dealsLastRow, 17).getValues();
  var countDealsRowSkip = 0;
  var countPostedDeals = 0;
  var countErrorDeals = 0;
  var details = [];
  var payments = [];

  for (var i = 1; i < dealsLastRow; i++) {
    var detailsRow = String(dealsValues[i][6]);
    var paymentsRow = String(dealsValues[i][13]);

    if (i + 1 == dealsLastRow) {
      var nextdealsRow = "取引を作成する";  //最終行に到達したら強制的に取引を作成

    } else {
      var nextdealsRow = String(dealsValues[i + 1][1]);
    };

    //detailsの作成
    if (detailsRow != "") {
      var accountItemId = parseInt(dealsValues[i][6]);
      var taxCode = parseInt(dealsValues[i][7]);
      var itemId = parseInt(dealsValues[i][8]);
      var sectionId = parseInt(dealsValues[i][9]);
      var tagIds = String(dealsValues[i][10]).split(",");

      if (tagIds == "") {
        tagIds = [];
      };

      var amountDetails = parseInt(dealsValues[i][11]);
      var description = String(dealsValues[i][12]);

      details.push({
        "account_item_id": accountItemId,
        "tax_code": taxCode,
        "item_id": isNaN(itemId) ? undefined : itemId,
        "section_id": isNaN(sectionId) ? undefined : sectionId,
        "tag_ids": tagIds,
        "amount": amountDetails,
        "description": description
      });
    };

    //paymentsの作成
    if (paymentsRow != "") {
      var date = Utilities.formatDate(dealsValues[i][13], "JST", "yyyy-MM-dd");
      var fromWalletableType = String(dealsValues[i][14]);
      var fromWalletableId = parseInt(dealsValues[i][15]);
      var amountPayments = parseInt(dealsValues[i][16]);

      payments.push({
        "date": date,
        "from_walletable_type": fromWalletableType,
        "from_walletable_id": fromWalletableId,
        "amount": amountPayments
      });
    };

    //取引を作成
    if (nextdealsRow == "") {
      countDealsRowSkip++;

    } else {
      var companyId = parseInt(dealsValues[i - countDealsRowSkip][0]);

      if (dealsValues[i - countDealsRowSkip][1] != "") {
        var issueDate = Utilities.formatDate(dealsValues[i - countDealsRowSkip][1], "JST", "yyyy-MM-dd");
      };

      if (dealsValues[i - countDealsRowSkip][2] != "") {
        var dueDate = Utilities.formatDate(dealsValues[i - countDealsRowSkip][2], "JST", "yyyy-MM-dd");
      };

      var type = String(dealsValues[i - countDealsRowSkip][3]);
      var partnerId = parseInt(dealsValues[i - countDealsRowSkip][4], 10);
      var refNumber = String(dealsValues[i - countDealsRowSkip][5]);

      var requestBody =
      {
        "company_id": companyId,
        "issue_date": issueDate,
        "due_date": dueDate,
        "type": type,
        "partner_id": isNaN(partnerId) ? undefined : partnerId,
        "ref_number": refNumber,
        "details": details,
        "payments": payments
      };

      Logger.log(JSON.stringify(requestBody));

      // POSTオプション
      var options = {
        "method": "POST",
        "contentType": "application/json",
        "headers": headers,
        "payload": JSON.stringify(requestBody),
        muteHttpExceptions: true
      };

      // POSTリクエスト
      var res = UrlFetchApp.fetch(requestUrl, options);

      if (res.getResponseCode() == 201) {
        countPostedDeals++;
      } else {
        Logger.log("Error response: " + res.getContentText());
        countErrorDeals++;
      }

      countDealsRowSkip = 0;
      details.length = 0;
      payments.length = 0;
    };
  };
  // 結果のアラート表示
  SpreadsheetApp.getUi().alert( countPostedDeals + "件の取引を送信しました" );
  if ( countErrorDeals != 0 ) {
    SpreadsheetApp.getUi().alert( countErrorDeals + "件の取引の送信に失敗しました" );
  };
}