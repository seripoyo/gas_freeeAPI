
/*********************************************************
 * カスタムメニューgasMenuとfreeeMenuを作成
 * freeeMenuに追加される項目はfreee_menu.gsに記載
******************************************************** */


/******************************************************************
関数：onOpen
概要：スプシ起動時にmenuを作成
******************************************************************/
function onOpen() {
  menu();
}

/******************************************************************
関数：menu
概要：onOpen()でgasMenuとfreeeMenuを作成
******************************************************************/

function menu() {
  var ui = SpreadsheetApp.getUi();

/** gasMenuを作成
**************************************************************/
  var gasMenu = ui.createMenu('GAS操作メニュー');
  var folderUrl = PropertiesService.getUserProperties().getProperty('folderUrl');
  Logger.log(folderUrl); // デバッグ情報
  if (folderUrl) {
    gasMenu.addItem('Googleドライブでフォルダを開く', 'openFolder');
  }
  gasMenu.addItem('売上履歴をインポートする', 'copyDataFromMultipleSheets'); //請求書出力gs
  gasMenu.addItem('売上履歴シートをリセットする', 'reset_Sheet');
  gasMenu.addToUi();
  
/** freeeMenuを作成
**************************************************************/
  var freeeMenu = ui.createMenu('freee連携機能');
  freeeMenu.addItem('①freeeと連携', 'alertAuth');
  freeeMenu.addItem('②事業所を選択', 'GetMyCompaniesID');
  freeeMenu.addItem('コールバックURLはこちら', 'showCallbackUrl');
  freeeMenu.addItem('クライアントID＆シークレットの入力', 'inputClientInfo');
  freeeMenu.addItem('売上データ送信', 'postDealsToFreee');
  freeeMenu.addItem('アクセストークン', 'showAlertWithAccessToken');

  freeeMenu.addToUi();
}


/******************************************************************
追加されるメニュー：gasMenu
関数：openFolder
概要：Googleドライブに作成したフォルダのページを開く
******************************************************************/

function openFolder() {
  // フォルダを開く関数
  var folderUrl = PropertiesService.getUserProperties().getProperty('folderUrl');
  if (folderUrl) {
    var html = HtmlService.createHtmlOutput('<script>window.open("' + folderUrl + '");google.script.host.close();</script>');
    SpreadsheetApp.getUi().showModalDialog(html, '別タブでフォルダを開いています');
  }
}

/******************************************************************
追加されるメニュー：gasMenu
関数：reset_Sheet
概要：売上履歴シートのヘッダー以外を削除してリセット
******************************************************************/
function reset_Sheet() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('売上履歴シートをリセットしますか？', ui.ButtonSet.YES_NO);

  if (response == ui.Button.YES) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("売上履歴");
    if (sheet) {
      sheet.getRange("2:247").clearContent();
      ui.alert("売上履歴シートがリセットされました。");
    } else {
      ui.alert("「売上履歴」シートが見つかりません。");
    }
  }
}
