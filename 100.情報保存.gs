// ------------------------------------------------------------------------------------------
// 処理：選択した事業所IDをプロパティに保存
/** @param {string} companyId - 選択された事業所のID  */
// ------------------------------------------------------------------------------------------
function setSelectedCompanyId(companyId) {
  try {
    var userProperties = PropertiesService.getUserProperties();
    userProperties.setProperty("selectedCompanyId", companyId);
  } catch (e) {
    // 事業所IDの保存に失敗した場合のアラート表示
    SpreadsheetApp.getUi().alert("選択した事業所IDを保存できませんでした。エラー: " + e.message);
  }
}
// ------------------------------------------------------------------------------------------
// 処理：保存された事業所IDを取得する関数
/**  @return {string} - 保存された事業所ID  */
// ------------------------------------------------------------------------------------------
function getSelectedCompanyId() {
  var userProperties = PropertiesService.getUserProperties();
  var selectedCompanyId = userProperties.getProperty('selectedCompanyId');
    // 取得したIDを整数に変換
  return parseInt(selectedCompanyId, 10);
}
// ------------------------------------------------------------------------------------------
// 処理：新規登録も含め、取引先情報のIDと名前を保存
/**  @return {string} - 保存された事業所ID  */
// ------------------------------------------------------------------------------------------
// 取引先データを保存する関数
function savePartnersData(partners) {
  var partnersData = partners.map(p => ({id: p.id, name: p.name}));
  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty("partnersData", JSON.stringify(partnersData));
  
  Logger.log("保存した取引先: " + JSON.stringify(partnersData));
}


/**
 * 指定されたシートの特定の列で情報が入力されている最終行を取得する関数
 * @param {string} sheetName - 取得するシートの名前
 * @param {number} column - 情報を確認する列の番号 (例: A列は1, B列は2)
 * @return {number} - 情報が入力されている最終行の行番号
 */
function getLastRowInColumn(sheetName, column) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var columnData = sheet.getRange(1, column, sheet.getMaxRows(), 1).getValues();
  var lastRow = columnData.findIndex(row => row[0] === "") + 1; // 最終行を見つける

  // 空のセルが見つかった場合は、その一つ前の行が最終行
  // 空のセルが見つからない場合は、最後の行が最終行
  return lastRow === 0 ? columnData.length : lastRow;
}
// ------------------------------------------------------------------------------------------
// 税区分一覧と取得した税区分が一致する情報を保存
/**  @return {string} - 保存された事業所ID  */
// ------------------------------------------------------------------------------------------

function saveMatchingTaxesData(matchingTaxes) {
  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty("matchingTaxes", JSON.stringify(matchingTaxes));
  Logger.log("保存した税区分: " + JSON.stringify(matchingTaxes));
}
// ------------------------------------------------------------------------------------------
// 作成する関数：showAlertWithAccessToken
// 実行内容：GET/ポップアップ出力
// 処理：メニュー選択時にアクセストークンをアラートで出力
// ------------------------------------------------------------------------------------------
function showAlertWithAccessToken() {
  var service = getService();
  if (service.hasAccess()) {
    var accessToken = service.getAccessToken();
    SpreadsheetApp.getUi().alert("アクセストークン: " + accessToken);
  } else {
    SpreadsheetApp.getUi().alert("アクセストークンを取得できませんでした。");
  }
}
// =============================================================-
// 内容：情報を保存
// 使用API：https://api.freee.co.jp/api/1/taxes/companies/
// 実行内容：GET・POST
// -=============================================================
function saveItemsData(items) {
  var itemsData = items.map(function(item) {
    return { id: item.id, name: item.name };
  });

  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty("itemsData", JSON.stringify(itemsData));
  Logger.log("Saved items data: " + JSON.stringify(itemsData));
}