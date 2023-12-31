
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
  var freeeMenu = ui.createMenu('基本設定');
  freeeMenu.addItem('①最初の認証を行う', 'alertAuth_First')

  // ユーザープロパティからフォルダURLを取得
  var folderUrl = PropertiesService.getUserProperties().getProperty('folderUrl');

  // フォルダURLが存在しない場合のみ、「専用フォルダ＆サンプルを作る」を追加
  if (!folderUrl) {
    freeeMenu.addItem('②専用フォルダ＆サンプルを作る', 'create_Folder_And_Update_Menu');
  }

  // その他のメニュー項目を追加
  if (folderUrl) {
    freeeMenu.addItem('作成したフォルダを別タブで開く', 'openFolder');
  }

  freeeMenu.addItem('請求書を取引一覧に読み込む', 'copy_Data_From_MultipleSheets')
  freeeMenu.addItem('読み込んだ取引一覧を削除する', 'reset_Sheet');

  // freeeMenu.addItem('アクセストークン表示', 'showAlertWithAccessToken');
  freeeMenu.addToUi();

  /** gasMenuを作成
**************************************************************/
  var gasMenu = ui.createMenu('過去請求書の読み込み');
  // ユーザープロパティからフォルダURLを取得
  var folderUrl = PropertiesService.getUserProperties().getProperty('folderUrl');

  gasMenu.addItem('過去の請求書を一覧から抽出する', 'listSpreadsheetsByYear')
  gasMenu.addItem('選択した過去の請求書を読み込む', 'past_invoice_setting')
  gasMenu.addItem('過去の取引一覧をリセットする', 'past_invoice_setting')

  gasMenu.addToUi();


  /** freeeMenu2を作成
  **************************************************************/

  var freeeMenu2 = ui.createMenu('freee連携');

  freeeMenu2.addItem('④freeeと連携する', 'show_CallbackUrl_and_Applink');
  freeeMenu2.addItem('⑤売上データを送信する', 'submit_freee');
  // freeeMenu.addItem('アクセストークン表示', 'showAlertWithAccessToken');
  freeeMenu2.addToUi();
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
/******************************************************************
追加されるメニュー：gasMenu
関数：reset_bef_Sheet
概要：移行用スプシ一覧シートのヘッダー以外を削除してリセット
******************************************************************/
function reset_bef_Sheet() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('移行用スプシ一覧シートをリセットしますか？', ui.ButtonSet.YES_NO);

  if (response == ui.Button.YES) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("移行用スプシ一覧");
    if (sheet) {
      sheet.getRange("3:247").clearContent();
      ui.alert("移行用スプシ一覧シートがリセットされました。");
    } else {
      ui.alert("「移行用スプシ一覧」シートが見つかりません。");
    }
  }
}

/******************************************************************
関数：showAlertWithAccessToken
概要：メニュー選択時にアクセストークンをアラートで出力
******************************************************************/

function showAlertWithAccessToken() {
  var service = getService();
  if (service.hasAccess()) {
    var accessToken = service.getAccessToken();
    SpreadsheetApp.getUi().alert("アクセストークン: " + accessToken);
  } else {
    SpreadsheetApp.getUi().alert("アクセストークンを取得できませんでした。");
  }
}

/******************************************************************
関数：submit_freee
概要：SelectModal事業所がで選択され、そのIDを取得したら連動して実行
******************************************************************/

function submit_freee() {
  // getMyCompaniesID();
  manage_Walletables(); //口座
  get_Taxes(); //税区分
  get_AccountItems();//勘定科目
  manage_Partners();//取引先
  get_Items_Register();//品目
  dealsTranscription(); //取引データを作成して
  postDeals(); //送信！

}