/******************************************************************
 * 新規登録をする必要がない各情報の取得＆保存
 * このファイルでは口座・税区分・勘定科目を一覧で取得＆保存する
 * 品目と取引先は登録も行い長くなるので別ファイルに格納
******************************************************************/


/******************************************************************
関数：manage_Walletables
概要：勘定科目一覧を取得し、整形して保存する
******************************************************************/
function manage_Walletables() {
  var freeeApp = getService(); // freeeサービスを取得
  var accessToken = freeeApp.getAccessToken(); // アクセストークンを取得
  var companyId = getSelectedCompanyId(); // 選択中の会社のIDを取得
  var requestUrl = "https://api.freee.co.jp/api/1/walletables?company_id=" + companyId + "&with_balance=true";
  var headers = { "Authorization": "Bearer " + accessToken };

  var options = {
    "method": "get",
    "headers": headers
  };

  // APIからデータ取得
  var res = UrlFetchApp.fetch(requestUrl, options).getContentText();
  var walletablesResponse = JSON.parse(res);
  var walletables = walletablesResponse.walletables;

  // データ処理
  var processedWalletables = walletables.map(function (walletable) {
    return {
      from_walletable_id: parseInt(walletable.id, 10).toString(), // IDを整数に変換して文字列化
      name: walletable.name,
      from_walletable_type: walletable.type,
      bank_id: walletable.bank_id ? parseInt(walletable.bank_id, 10).toString() : null, // bank_idがnullでなければ整数に変換
      walletable_balance: parseInt(walletable.walletable_balance, 10), // walletable_balanceを整数に変換
      last_balance: parseInt(walletable.last_balance, 10) // last_balanceを整数に変換
    };
  });

  // データをユーザープロパティに保存
  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty("walletablesData", JSON.stringify(processedWalletables));
  Logger.log("口座一覧を保存しました");

  return processedWalletables; // 結果の配列を返す
}

/******************************************************************
関数：get_Taxes
概要：選択中の会社の税区分一覧を取得し、整形して保存する
******************************************************************/
function get_Taxes() {
  var freeeApp = getService(); // freeeサービスを取得
  var accessToken = freeeApp.getAccessToken(); // アクセストークンを取得
  var companyId = getSelectedCompanyId(); // 選択中の会社のIDを取得
  var requestUrl = "https://api.freee.co.jp/api/1/taxes/companies/" + companyId;
  var headers = { "Authorization": "Bearer " + accessToken };
  var options = { "method": "get", "headers": headers };

  var response = UrlFetchApp.fetch(requestUrl, options).getContentText();
  var taxesResponse = JSON.parse(response);
  var taxes = taxesResponse.taxes;

  // 税区分一覧のIDを整数に変換して配列に格納
  var taxesData = taxes.map(function (tax) {
    return {
      tax_code: parseInt(tax.code, 10).toString(), // IDを整数に変換して文字列化
      name_ja: tax.name_ja,
      default_Tax_Ccode: 34,
    };
  });

  // 税区分データをユーザープロパティに保存
  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty("taxesData", JSON.stringify(taxesData));
  Logger.log("税区分を保存しました");

  return taxesData; // 結果の配列を返す
}

/******************************************************************
 * 関数：get_AccountItems
 * 概要：取引一覧のスプレッドシートから勘定科目情報を取得し、プロパティサービスに保存する関数
 ******************************************************************/
function get_AccountItems() {
  var freeeApp = getService();
  var accessToken = freeeApp.getAccessToken();
  var companyId = getSelectedCompanyId();
  var requestUrl = "https://api.freee.co.jp/api/1/account_items?company_id=" + companyId;
  var headers = { "Authorization": "Bearer " + accessToken };
  var options = { "method": "get", "headers": headers };

  var response = UrlFetchApp.fetch(requestUrl, options).getContentText();
  var accountItemsResponse = JSON.parse(response);
  var accountItems = accountItemsResponse.account_items;

  // スプレッドシートのF列のデータを取得
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("取引一覧");
  var lastRowInFColumn = getLastRowInColumn("取引一覧", 6);

  // lastRowInFColumnが1未満の場合は処理をスキップ
  if (lastRowInFColumn < 1) {
    Logger.log("F列にデータがありません。");
    return;
  }
  var fColumnData = sheet.getRange(2, 6, lastRowInFColumn).getValues();

  // 合致する勘定科目のnameとidを保存する配列
  var matchingAccountItems = [];

  fColumnData.forEach(function (fValue) {
    var matchingAccountItem = accountItems.find(accountItem => accountItem.name === fValue[0]);
    if (matchingAccountItem) {
      matchingAccountItems.push({
        id: matchingAccountItem.id,
        name: matchingAccountItem.name,
        defaultTaxId: matchingAccountItem.default_tax_id
      });
    }
  });

  // 合致する勘定科目をプロパティサービスに保存
  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty("matchingAccountItems", JSON.stringify(matchingAccountItems));

  // 確認のためにログに出力
  Logger.log("勘定科目を保存しました");

  return matchingAccountItems; // 結果の配列を返す
}
/** --------------------------------------------------------------------
 * 関数：getLastRowInColumn
 * 概要：指定したシートと列番号に基づいて情報が入力されている最終行を取得する関数
 * 
 * @param {string} sheetName - 検索対象のシート名
 * @param {number} columnNumber - 情報が入力されている列の列番号（1から始まる）
 * @return {number} - 情報が入力されている最終行の行番号
---------------------------------------------------------------------- **/
function getLastRowInColumn(sheetName, columnNumber) {
  // 指定したシート名からシートを取得
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  
  // 指定した列のデータを取得し、2次元の配列として格納
  var columnData = sheet.getRange(1, columnNumber, sheet.getMaxRows(), 1).getValues();
  
  // 列のデータから空でないセルが現れる最後の行を取得
  var lastRow = columnData.filter(String).length;
  
  // 最終行の行番号を返す
  return lastRow;
}