/*
***********************************************************************************
参照ライブラリ
title               |OAuth2
project_key  |1B7FSrk5Zi6L1rSxxTDgDEUsPzlukDsi4KGuTMorsTQHhGBzBkMun4iDF
***********************************************************************************
*/

// 各自で作成したアプリのアプリ詳細画面からClient IDとClient Secretをカスタムメニューへコピペ
// ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
function inputClientInfo() {
  //　まとめて入力したいのと、placeholderを付けたい
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('freee API設定', 'Client IDを入力してください', ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() == ui.Button.OK) {
    var clientId = response.getResponseText();
    PropertiesService.getUserProperties().setProperty('freeeClientId', clientId);

    response = ui.prompt('freee API設定', 'Client Secretを入力してください', ui.ButtonSet.OK_CANCEL);
    if (response.getSelectedButton() == ui.Button.OK) {
      var clientSecret = response.getResponseText();
      PropertiesService.getUserProperties().setProperty('freeeClientSecret', clientSecret);

      // 認証URLの取得と表示
      alertAuth();
    }
  }
}
// inputClientInfo()実行後に認証URL（認証のエンドポイント）をポップアップ表示
// ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
function alertAuth() {
  var service = getService();
  var authorizationUrl = service.getAuthorizationUrl();

  var html = HtmlService.createHtmlOutput('<html><body><a href="' + authorizationUrl + '" target="_blank">認証ページを開く</a></body></html>')
    .setWidth(400)
    .setHeight(60);
  SpreadsheetApp.getUi().SelectModal(html, 'リンクを開いて認証を行ってください');
}

// 認証用URLのコピー用ポップアップ
// ------------------------------------------------------------
function showCallbackUrl() {
  var scriptId = ScriptApp.getScriptId();
  var callbackUrl = 'https://script.google.com/macros/d/' + scriptId + '/usercallback';
  var htmlContent = '<input id="url" value="' + callbackUrl + '" readonly style="width:100%">' +
    '<button onclick="copyToClipboard()">コピーする</button>' +
    '<script>' +
    'function copyToClipboard() {' +
    '  var copyText = document.getElementById("url");' +
    '  copyText.select();' +
    '  document.execCommand("copy");' +
    '  alert("コピーしました: " + copyText.value);' +
    '}' +
    '</script>';

  var html = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(400)
    .setHeight(150);
  SpreadsheetApp.getUi().SelectModal(html, 'コールバックURL');
}



// freeeAPIのサービスを取得
// ------------------------------------------------------------
function getService() {
  var userProperties = PropertiesService.getUserProperties();
  var clientId = userProperties.getProperty('freeeClientId');
  var clientSecret = userProperties.getProperty('freeeClientSecret');

  return OAuth2.createService('freee')
    .setAuthorizationBaseUrl('https://accounts.secure.freee.co.jp/public_api/authorize')
    .setTokenUrl('https://accounts.secure.freee.co.jp/public_api/token')
    .setClientId(clientId)
    .setClientSecret(clientSecret)
    .setCallbackFunction('authCallback')
    .setPropertyStore(PropertiesService.getUserProperties());
}

//　認証用のコールバック関数(アクセストークンの取得)
// ------------------------------------------------------------
function authCallback(request) {
  var service = getService();
  var isAuthorized = service.handleCallback(request);
  if (isAuthorized) {
    return HtmlService.createHtmlOutput('認証に成功しました。タブを閉じてください。');
    //　これ自動で閉じるようにならんかな
  } else {
    return HtmlService.createHtmlOutput('認証に失敗しました。');
  }
}

/******************************************************************
function name |clearService
summary       |認証解除
******************************************************************/
function clearService() {
  OAuth2.createService( "freee" )
      .setPropertyStore( PropertiesService.getUserProperties() )
      .reset();
  
}