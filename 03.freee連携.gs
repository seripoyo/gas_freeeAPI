/*
***********************************************************************************
参照ライブラリ
title               |OAuth2
project_key  |1B7FSrk5Zi6L1rSxxTDgDEUsPzlukDsi4KGuTMorsTQHhGBzBkMun4iDF
***********************************************************************************
*/

// 各自で作成したアプリのアプリ詳細画面からClient IDとClient Secretをカスタムメニューへコピペ
// ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
// カスタムダイアログを表示する関数
function inputClientInfo() {
  var html = HtmlService.createHtmlOutputFromFile('ClientInfoForm')
    .setWidth(400)
    .setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(html, 'freee API設定');
}

// プロパティにクライアント情報を保存する関数
function saveClientInfo(clientId, clientSecret) {
  PropertiesService.getUserProperties()
    .setProperty('freeeClientId', clientId)
    .setProperty('freeeClientSecret', clientSecret);
}

// inputClientInfo()実行後に認証URL（認証のエンドポイント）を出力
// ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
function alertAuth() {
  var service = getService();
  var authorizationUrl = service.getAuthorizationUrl();

  var html = HtmlService.createHtmlOutput('<html><body><a href="' + authorizationUrl + '" target="_blank">認証ページを開く</a></body></html>')
    .setWidth(400)
    .setHeight(60);
  SpreadsheetApp.getUi().showModalDialog(html, 'リンクを開いて認証を行ってください');
}

// 認証用コールバックURLのコピペできるように出力
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
    .setHeight(100);
  SpreadsheetApp.getUi().showModalDialog(html, 'コールバックURL'); // ここを変更しました
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
    var html = HtmlService.createHtmlOutput('<html><body>' +
      '認証に成功しました。このタブはすぐに閉じます。<br>' +
      '<script>setTimeout(function() { window.top.close(); }, 5000);</script>' +
      '</body></html>');
    return html;
  } else {
    return HtmlService.createHtmlOutput('認証に失敗しました。');
  }
}


/******************************************************************
function name |clearService
summary       |認証解除
******************************************************************/
function clearService() {
  OAuth2.createService("freee")
    .setPropertyStore(PropertiesService.getUserProperties())
    .reset();

}