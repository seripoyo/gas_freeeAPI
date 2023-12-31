/******************************************************************
 * 
 * 今まで取得した情報を取引作成のための専用シートに出力するよ！
 * 情報を保存していなくとも問題無い日程・金額・支出はそのまま置換
 * それ以外は保存してあるやつから情報を引っ張るよ！というやつ
 * それぞれ全部引っ張ったら取引シートに入力出来るよ。
 * 
******************************************************************/

/******************************************************************
関数：dealsTranscription
概要：入力された取引をPOST用の取引シートに転写する
******************************************************************/

function dealsTranscription() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var targetSheet = ss.getSheetByName("取引");

  // 転写するシート名
  var sourceSheetNames = ["取引一覧", "過去の取引一覧"];

  // データの初期化
  targetSheet.getRange(2, 1, targetSheet.getLastRow(), targetSheet.getLastColumn()).clear();

  // データ転写処理
  sourceSheetNames.forEach(function (sourceSheetName) {
    var sourceSheet = ss.getSheetByName(sourceSheetName);
    var lastRow = sourceSheet.getLastRow();
    var numRows = lastRow - 1;  // 1行目はヘッダー行として除外

    if (numRows > 0) {
      // データ転写処理を行う
      transferDataFromSheet(sourceSheet, targetSheet);
    }
  });

  // スプレッドシートの特定の列を中央揃えに設定
  var centerAlignedColumns = ['A', 'B', 'D', 'E', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'O', 'P', 'Q', 'R'];
  centerAlignedColumns.forEach(function (column) {
    var range = targetSheet.getRange(column + "2:" + column + targetSheet.getLastRow());
    range.setHorizontalAlignment("center");
  });
}

function transferDataFromSheet(sourceSheet, targetSheet) {
  // PropertiesServiceから保存されたデータを取得
  var companyId = getSelectedCompanyId();
  var savedPartnersData = saved_PartnersData();
  var savedItemsData = getSavedItemsData();
  var savedAccountItemsData = getSavedAccountItemsData();
  var savedTaxesData = getSavedTaxesData();
  var savedWalletablesData = getSavedWalletablesData();

  // ソースシートからデータを取得
  var sourceSheetValues = sourceSheet.getRange(2, 1, sourceSheet.getLastRow() - 1, 17).getValues();

  sourceSheetValues.forEach(function (row, index) {
    var currentRow = targetSheet.getLastRow() + 1;  // 現在の最終行に追加

    // companyId
    targetSheet.getRange(currentRow, 1).setValue(companyId);

    // typeを設定（収入/支出に基づいてincome/expenseを設定）
    var typeValue = row[0] == "収入" ? "income" : "expense";
    targetSheet.getRange(currentRow, 4).setValue(typeValue);

    // 発生日
    targetSheet.getRange(currentRow, 2).setValue(formatDate(row[2]));

    // 決済期日
    targetSheet.getRange(currentRow, 3).setValue(formatDate(row[3]));

    // 取引先
    var partnerName = row[4];
    // 取引先名からpartner_idを検索する関数
    function findPartnerIdByName(partnersData, partnerName) {
      partnerName = partnerName.trim(); // 余分な空白を削除
      for (var i = 0; i < partnersData.length; i++) {
        if (partnersData[i].name.trim() === partnerName) { // こちらも余分な空白を削除
          return partnersData[i].partner_id;
        }
      }
      return ''; // 該当する取引先がなければ空文字を返す
    }
    var partnerId = findPartnerIdByName(savedPartnersData, partnerName);
    if (partnerId) {
      targetSheet.getRange(currentRow, 5).setValue(partnerId);
    }

    // 勘定科目
    var accountItemName = row[5];
    //         // 勘定科目名からidを検索する関数
    function findAccountIdByName(accountItemsData, accountItemName) {
      for (var i = 0; i < accountItemsData.length; i++) {
        if (accountItemsData[i].name === accountItemName) {
          return accountItemsData[i].id; // 該当するidを返す
        }
      }
      return ''; // 該当する勘定科目がなければ空文字を返す
    }
    var accountId = findAccountIdByName(savedAccountItemsData, accountItemName);
    if (accountId) {
      targetSheet.getRange(currentRow, 7).setValue(accountId);
    }

    // 税計算区分
    var taxName = row[6];
    //         //税計算区分
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
    var taxCode = findTaxCodeByName(savedTaxesData, taxName);
    if (taxCode) {
      targetSheet.getRange(currentRow, 8).setValue(taxCode);
    }

    // 品目
    var itemName = row[11];
    //         // 品目名からitem_idを検索する関数
function findItemIdByName(itemsData, itemName) {
  itemName = itemName.trim(); // ここで品目名から余分な空白を除去
  for (var i = 0; i < itemsData.length; i++) {
    if (itemsData[i].name === itemName) {
      return itemsData[i].item_id; // 該当するitem_idを返す
    }
  }
  return ''; // 該当する品目がなければ空文字を返す
}
    var itemId = findItemIdByName(savedItemsData, itemName);
    if (itemId) {
      targetSheet.getRange(currentRow, 9).setValue(itemId);
    }

    // 金額
    targetSheet.getRange(currentRow, 12).setValue(row[7]);

    // 税額
    // targetSheet.getRange(currentRow, 13).setValue(row[9]);
 var taxAmount = row[9] ? row[9] : 0;
    targetSheet.getRange(currentRow, 13).setValue(taxAmount);

    // 決済金額
    targetSheet.getRange(currentRow, 17).setValue(row[16]);

    // 決済日
    var settlementDate = formatDate(row[14]);
    targetSheet.getRange(currentRow, 14).setValue(settlementDate);

    // 口座情報
    function findWalletableIdByName(walletablesData, walletableName) {
      for (var i = 0; i < walletablesData.length; i++) {
        if (walletablesData[i].name === walletableName) {
          return walletablesData[i].from_walletable_id; // 該当するfrom_walletable_idを返す
        }
      }
      return ''; // 該当する口座がなければ空文字を返す
    }
    var savedWalletablesData = JSON.parse(PropertiesService.getUserProperties().getProperty("walletablesData"));

    var walletableName = row[15];
    var walletableId = findWalletableIdByName(savedWalletablesData, walletableName);
    var walletableType = savedWalletablesData.find(walletable => walletable.from_walletable_id === walletableId)?.from_walletable_type;

    var walletableId = findWalletableIdByName(savedWalletablesData, walletableName);
    if (walletableId) {
      targetSheet.getRange(currentRow, 16).setValue(walletableId); // 口座ID-from_walletable_id
      targetSheet.getRange(currentRow, 15).setValue(walletableType); //口座種別-from_walletable_type

    }
    var walletableType = savedWalletablesData.find(walletable => walletable.from_walletable_id === walletableId)?.from_walletable_type;
  });
}


/******************************************************************
 * 関数：getLastRowNumber
 * 概要：指定されたシートと列で最終行を返す
 * 
 * @param {number} column - チェックする列の列番号（1から始まる）
 * @param {string} sheetname - チェックするシートの名前
 * @return {number} - 最終の有効な行の行番号
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
  return 0; // 有効な行が見つからない場合、0を返す
}


/******************************************************************
 * 関数：getSavedAccountItemsData
 * 概要：ユーザープロパティから保存された勘定科目データを取得する関数
 * 
 * @return {Array} - 勘定科目データの配列
 ******************************************************************/
function getSavedAccountItemsData() {
  var userProperties = PropertiesService.getUserProperties();
  var dataString = userProperties.getProperty("matchingAccountItems");
  return dataString ? JSON.parse(dataString) : []; // データが存在しない場合は空の配列を返す
}

/******************************************************************
 * 関数：getSavedTaxesData
 * 概要：ユーザープロパティから保存された税データを取得する関数
 * 
 * @return {Array} - 税データの配列
 ******************************************************************/
function getSavedTaxesData() {
  var userProperties = PropertiesService.getUserProperties();
  var dataString = userProperties.getProperty("taxesData");
  return dataString ? JSON.parse(dataString) : []; // データが存在しない場合は空の配列を返す
}

/******************************************************************
 * 関数：getSavedWalletablesData
 * 概要：ユーザープロパティから保存された口座データを取得する関数
 * 
 * @return {Array} - 口座データの配列
 ******************************************************************/
function getSavedWalletablesData() {
  var userProperties = PropertiesService.getUserProperties();
  var dataString = userProperties.getProperty("walletablesData");
  return dataString ? JSON.parse(dataString) : []; // データが存在しない場合は空の配列を返す
}