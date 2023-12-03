
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

    // 勘定科目名からaccount_item_idを検索する関数
    function findAccountIdByName(accountItemsData, accountItemName) {
      for (var i = 0; i < accountItemsData.length; i++) {
        if (accountItemsData[i].name === accountItemName) {
          return accountItemsData[i].account_item_id; // 該当するaccount_item_idを返す
        }
      }
      return ''; // 該当する勘定科目がなければ空文字を返す
    }
    var savedAccountItemsData = JSON.parse(PropertiesService.getUserProperties().getProperty("matchingAccountItems"));

    // 勘定科目名に基づいてaccount_item_idを検索
    var accountItemName = row[5]; // 勘定科目名
    var accountId = findAccountIdByName(savedAccountItemsData, accountItemName);

    // 取引シートにaccount_item_idを設定
    if (accountId) {
      targetSheet.getRange(i + 2, 7).setValue(accountId);
    }

    //税計算区分
    var savedTaxesData = JSON.parse(PropertiesService.getUserProperties().getProperty("taxesData"));
    // 税区分名からtax_codeを検索する関数
    function findTaxCodeByName(taxesData, taxName) {
      for (var i = 0; i < taxesData.length; i++) {
        if (taxesData[i].name_ja === taxName) {
          return taxesData[i].tax_code; // 該当するtax_codeを返す
        }
      }
      return ''; // 該当する税区分がなければ空文字を返す
    }

    // 税区分名に基づいてtax_codeを検索
    var taxName = row[6]; // 税区分名
    var taxCode = findTaxCodeByName(savedTaxesData, taxName);

    // 取引シートにtax_codeを設定
    if (taxCode) {
      targetSheet.getRange(i + 2, 8).setValue(taxCode);
    }

    // var taxCodeName = sourceSheetValues[i][6];
    // if (taxCodeName != "") {
    //   for (var e = 0; e < taxesCodesLastRow; e++) {
    //     if (taxesCodesValues[e][2] == taxCodeName) {
    //       targetSheet.getRange(i + 2, 8).setValue(taxesCodesValues[e][0]);
    //       break;
    //     };
    //   };
    // };

    //品目  品目でaccount_item_id,
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

    function findWalletableIdByName(walletablesData, walletableName) {
      for (var i = 0; i < walletablesData.length; i++) {
        if (walletablesData[i].name === walletableName) {
          return walletablesData[i].from_walletable_id; // 該当するfrom_walletable_idを返す
        }
      }
      return ''; // 該当する口座がなければ空文字を返す
    }
  var savedWalletablesData = JSON.parse(PropertiesService.getUserProperties().getProperty("walletablesData"));

    // 口座名に基づいてfrom_walletable_idを検索
    var walletableName = row[15]; // 口座名
    var walletableId = findWalletableIdByName(savedWalletablesData, walletableName);
      var walletableType = savedWalletablesData.find(walletable => walletable.from_walletable_id === walletableId)?.from_walletable_type;


    // 取引シートにfrom_walletable_idとfrom_walletable_typeを設定
    if (walletableId) {
      targetSheet.getRange(i + 2, 17).setValue(walletableId);
      targetSheet.getRange(i + 2, 16).setValue(walletableType);
      // targetSheet.getRange(i + 2, 15).setValue

    }

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