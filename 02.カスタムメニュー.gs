
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

/** freeeMenuを作成
**************************************************************/
  var freeeMenu = ui.createMenu('freee連携メニュー');
  freeeMenu .addItem('①free連携用アプリを作成', 'openLink')
  freeeMenu.addItem('③コールバックURLの表示', 'show_CallbackUrl_and_Applink');
  freeeMenu.addItem('④クライアントID＆シークレットの入力', 'inputClientInfo');
  freeeMenu.addItem('②アプリ連携&事業所選択', 'auth_Setting');
  freeeMenu.addItem('⑥請求書をインポート', 'invoice_import'); //請求書出力gs
  freeeMenu.addItem('⑦売上データを送信', 'submit_freee');
  // freeeMenu.addItem('アクセストークン', 'showAlertWithAccessToken');

  freeeMenu.addToUi();

/** gasMenuを作成
**************************************************************/
  var gasMenu = ui.createMenu('オプション操作');
  var folderUrl = PropertiesService.getUserProperties().getProperty('folderUrl');
  // Logger.log(folderUrl); // デバッグ情報
  if (folderUrl) {
    gasMenu.addItem('Googleドライブでフォルダを開く', 'openFolder');
  }
  // gasMenu.addItem('請求書をインポートする', 'copyDataFromMultipleSheets'); //請求書出力gs
  gasMenu.addItem('取引一覧シートをリセット', 'reset_Sheet');
  gasMenu.addItem('アクセストークン表示', 'showAlertWithAccessToken');
  gasMenu.addToUi();
  

}
// リンクを開く関数
function openLink() {
  var url = 'https://app.secure.freee.co.jp/developers/companies';
  var html = HtmlService.createHtmlOutput('<html><script>window.open("' + url + '");google.script.host.close();</script></html>')
    .setWidth(400)
    .setHeight(30);
  SpreadsheetApp.getUi().showModalDialog(html, '連携アプリ作成画面を開いています');
}

/******************************************************************
追加されるメニュー：gasMenu
関数：openFolder
概要：Googleドライブに作成したフォルダのページを開く
******************************************************************/



function invoice_import() {
  copyDataFromMultipleSheets(); //請求書インポート
  // 少し遅延を入れる（必要に応じて）
  Utilities.sleep(3000);
  manage_Walletables(); //口座
  get_Taxes(); //税区分
  get_AccountItems();//勘定科目
  manage_Partners();//取引先
  get_Items_Register();//品目
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
概要：取引一覧シートのヘッダー以外を削除してリセット
******************************************************************/
function reset_Sheet() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('取引一覧シートをリセットしますか？', ui.ButtonSet.YES_NO);

  if (response == ui.Button.YES) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("取引一覧");
    if (sheet) {
      sheet.getRange("2:247").clearContent();
      ui.alert("取引一覧シートがリセットされました。");
    } else {
      ui.alert("「取引一覧」シートが見つかりません。");
    }
  }
}
