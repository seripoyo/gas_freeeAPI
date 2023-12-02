// メニュー作成
// ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝

// スプシ起動時に実行
function onOpen() {
  updateMenu();
}
function updateMenu() {
  var ui = SpreadsheetApp.getUi();

  // GAS操作メニューを作成
  var gasMenu = ui.createMenu('GAS操作メニュー');
  var folderUrl = PropertiesService.getUserProperties().getProperty('folderUrl');
  Logger.log(folderUrl); // デバッグ情報
  if (folderUrl) {
    gasMenu.addItem('Googleドライブでフォルダを開く', 'openFolder');
  }
  gasMenu.addItem('売上履歴をインポートする', 'copyDataFromMultipleSheets');
  gasMenu.addItem('売上履歴シートをリセットする', 'confirmAndResetSalesHistorySheet');
  gasMenu.addToUi();

  // freee連携のヒントメニューを作成
  var freeeMenu = ui.createMenu('freee連携のヒント');
    freeeMenu.addItem('①freeeと連携', 'alertAuth');
      freeeMenu.addItem('②事業所を選択', 'GetMyCompaniesID');
  freeeMenu.addItem('コールバックURLはこちら', 'showCallbackUrl');
  freeeMenu.addItem('クライアントID＆シークレットの入力', 'inputClientInfo');
    freeeMenu.addItem('売上データ送信', 'postDealsToFreee');
    freeeMenu.addItem('アクセストークン', 'showAlertWithAccessToken');

  freeeMenu.addToUi();
}
// 売上履歴シートの項目以外を削除
// ------------------------------------------------------------------------------------------
function confirmAndResetSalesHistorySheet() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('売上履歴シートをリセットしますか？', ui.ButtonSet.YES_NO);

  if (response == ui.Button.YES) {
    resetSalesHistorySheet();
  }
}

function resetSalesHistorySheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("売上履歴");
  if (sheet) {
    sheet.getRange("2:247").clearContent();
    SpreadsheetApp.getUi().alert("売上履歴シートがリセットされました。");
  } else {
    SpreadsheetApp.getUi().alert("「売上履歴」シートが見つかりません。");
  }
}

// Googleドライブに作成したフォルダのページを開く
// ------------------------------------------------------------------------------------------
function openFolder() {
  // フォルダを開く関数
  var folderUrl = PropertiesService.getUserProperties().getProperty('folderUrl');
  if (folderUrl) {
    var html = HtmlService.createHtmlOutput('<script>window.open("' + folderUrl + '");google.script.host.close();</script>');
    SpreadsheetApp.getUi().showModalDialog(html, '別タブでフォルダを開いています');
  }
}

function getAll(){
  getPartners();
  getWalletables();
  getAccountItems();
  getItems();
  getSections();
  getTags();
}
