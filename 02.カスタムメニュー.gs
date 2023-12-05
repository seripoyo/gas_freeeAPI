
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
  freeeMenu.addItem('①free連携用アプリを作成', 'openLink')
  freeeMenu.addItem('③コールバックURLの表示', 'show_CallbackUrl_and_Applink');
  //input_ClientInfo()で情報保存→alertAuthが実行→ウィンドウを閉じる際に自動でgetMyCompaniesIDが実行され事業所IDが保存される
  freeeMenu.addItem('④クライアントID＆シークレットの入力', 'input_ClientInfo'); 
  freeeMenu.addItem('⑦売上データを送信', 'submit_freee');
  // freeeMenu.addItem('アクセストークン', 'showAlertWithAccessToken');

  freeeMenu.addToUi();

  /** gasMenuを作成
  **************************************************************/
  var gasMenu = ui.createMenu('共通メニュー');
  gasMenu.addItem('最初の認証', 'alertAuth_First')
  gasMenu.addItem('フォルダ＆サンプル作成', 'create_Folder_And_Update?Menu')
  gasMenu.addItem('⑥請求書をインポート', 'invoice_import')

  var folderUrl = PropertiesService.getUserProperties().getProperty('folderUrl');
  if (folderUrl) {
    gasMenu.addItem('Googleドライブでフォルダを開く', 'openFolder');
  }

  gasMenu.addItem('取引一覧シートをリセット', 'reset_Sheet');
  gasMenu.addItem('アクセストークン表示', 'showAlertWithAccessToken');
  gasMenu.addToUi();


}

/******************************************************************
追加されるメニュー：gasMenu
関数：openFolder
概要：Googleドライブに作成したフォルダのページを開く
******************************************************************/



function invoice_import() {
  copy_Data_From_MultipleSheets(); //請求書インポート
  // 少し遅延を入れる（必要に応じて）
  // Utilities.sleep(3000);
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
