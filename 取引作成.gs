
/******************************************************************
function name |dealsTranscription
summary       |入力された取引をPOST用の取引シートに転写する
******************************************************************/

function dealsTranscription() {
  var sourceSheetName = "売上履歴";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName(sourceSheetName);
  var targetSheet = ss.getSheetByName("取引");

  // 各列の最終行を確認
  var lastRow = sourceSheet.getLastRow();
  var numRows = lastRow - 1; // 1行目はヘッダー行として除外
  if (numRows <= 0) {
    Logger.log("データがありません。");
    return;
  }

  // ソースシートからデータを取得
  var sourceSheetValues = sourceSheet.getRange(2, 1, numRows, 17).getValues();
  var companyId = getSelectedCompanyId();

  // PropertiesServiceから保存されたデータを取得
  var savedPartnersData = saved_PartnersData();
  var savedItemsData = getSavedItemsData();
  var savedAccountItems = getSavedAccountItemsData();
  var savedTaxesData = getSavedTaxesData();
  var savedWalletablesData = getSavedWalletablesData();

  // データの初期化
  targetSheet.getRange(2, 1, targetSheet.getLastRow(), targetSheet.getLastColumn()).clear();


  //転記
  // 以後の転記処理...
  for (var i = 0; i < sourceSheetValues.length; i++) {
    var row = sourceSheetValues[i];

    // companyId
    targetSheet.getRange(i + 2, 1).setValue(companyId);

    // typeを設定（収入/支出に基づいてincome/expenseを設定）
    var typeValue = row[0] == "収入" ? "income" : "expense";
    targetSheet.getRange(i + 2, 4).setValue(typeValue);

    //発生日
    targetSheet.getRange(i + 2, 2).setValue(sourceSheetValues[i][2]);

    //決済期日
    targetSheet.getRange(i + 2, 3).setValue(sourceSheetValues[i][3]);


    //取引先
    var partnerName = row[4]; // 取引先名

    // 取引先名からpartner_idを検索する関数
    function findPartnerIdByName(partnersData, partnerName) {
      for (var i = 0; i < partnersData.length; i++) {
        if (partnersData[i].name === partnerName) {
          return partnersData[i].partner_id; // 該当するpartner_idを返す
        }
      }
      return ''; // 該当する取引先がなければ空文字を返す
    }

    var partnerId = findPartnerIdByName(savedPartnersData, partnerName);
    // 取引シートにpartner_idを設定
    if (partnerId) {
      targetSheet.getRange(i + 2, 5).setValue(partnerId);
    }

    //勘定科目
    // var accountItemName = sourceSheetValues[i][5];
    // for (var e = 0; e < accountItemsLastRow; e++) {
    //   if (accountItemsValues[e][1] == accountItemName) {
    //     targetSheet.getRange(i + 2, 7).setValue(accountItemsValues[e][0]);
    //     targetSheet.getRange(i + 2, 8).setValue(accountItemsValues[e][5]);
    //     break;
    //   };
    // };

    //税計算区分
    // var taxCodeName = sourceSheetValues[i][6];
    // if (taxCodeName != "") {
    //   for (var e = 0; e < taxesCodesLastRow; e++) {
    //     if (taxesCodesValues[e][2] == taxCodeName) {
    //       targetSheet.getRange(i + 2, 8).setValue(taxesCodesValues[e][0]);
    //       break;
    //     };
    //   };
    // };

    //品目
    // var itemName = sourceSheetValues[i][11];
    // for (var e = 0; e < itemsLastRow; e++) {
    //   if (itemsValues[e][2] == itemName) {
    //     targetSheet.getRange(i + 2, 9).setValue(itemsValues[e][0]);
    //     break;
    //   };
    // };


    // 金額(amount)
    targetSheet.getRange(i + 2, 12).setValue(sourceSheetValues[i][7]);

    // 税額(vat)
    targetSheet.getRange(i + 2, 13).setValue(sourceSheetValues[i][9]);

    // 決済金額(amount)
    targetSheet.getRange(i + 2, 18).setValue(sourceSheetValues[i][16]);

    var issueDate = formatDate(row[2]);
    targetSheet.getRange(i + 2, 2).setValue(issueDate);

    // targetSheet.getRange(i + 2, 12).setValue(sourceSheetValues[i][7]);

    //備考
    // targetSheet.getRange(i + 2, 13).setValue(sourceSheetValues[i][10]);

    //決済日
    // targetSheet.getRange(i + 2, 15).setValue(sourceSheetValues[i][14]);
    var settlementDate = formatDate(row[14]);
    targetSheet.getRange(i + 2, 15).setValue(settlementDate);


    //from_walletable_type,from_walletable_id
    // var walletableName = sourceSheetValues[i][15];
    // for (var e = 0; e < walletablesLastRow; e++) {
    //   if (walletable_values[e][1] == walletableName) {
    //     targetSheet.getRange(i + 2, 15).setValue(walletable_values[e][3]);
    //     targetSheet.getRange(i + 2, 16).setValue(walletable_values[e][0]);
    //     break;
    //   };
    // };

    //amount_payments
    targetSheet.getRange(i + 2, 17).setValue(sourceSheetValues[i][16]);
  }

}

/******************************************************************
function name |formatDate()
summary       |日付をyyyy-mm-dd形式にフォーマット
******************************************************************/
function formatDate(date) {
  if (date) {
    return Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return '';
}

/******************************************************************
function name |getLastRowNumber()
summary       |値の入っている最終行を返す（なぜかgetLastRowだとが関係ない最終行が返る）
******************************************************************/

function getLastRowNumber(column, sheetname) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetname);
  var lastRow = sheet.getLastRow();
  var values = sheet.getRange(1, column, lastRow, 1).getValues();

  for (var i = values.length - 1; i >= 0; i--) {
    if (values[i][0] != "") {
      return i + 1;
    }
  }
  return 0;
}

// 各種保存されたデータを取得する関数
function getSavedAccountItemsData() {
  var userProperties = PropertiesService.getUserProperties();
  var dataString = userProperties.getProperty("matchingAccountItems");
  return dataString ? JSON.parse(dataString) : [];
}

function getSavedTaxesData() {
  var userProperties = PropertiesService.getUserProperties();
  var dataString = userProperties.getProperty("taxesData");
  return dataString ? JSON.parse(dataString) : [];
}

function getSavedWalletablesData() {
  var userProperties = PropertiesService.getUserProperties();
  var dataString = userProperties.getProperty("walletablesData");
  return dataString ? JSON.parse(dataString) : [];
}