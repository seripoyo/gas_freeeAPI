
/*********************************************************

 * カスタムメニューfreeeMenuの内容となる関数を作成
 * freeeMenuに追加される項目はfreee_menu.gsに記載


******************************************************** */



/******************************************************************
関数：alertAuth
概要：inputClientInfo()実行後に認証URL（認証のエンドポイント）を出力
******************************************************************/
function alertAuth() {
  var service = getService();
  var authorizationUrl = service.getAuthorizationUrl();

  var html = HtmlService.createHtmlOutput('<html><body>' +
    '<a href="' + authorizationUrl + '" target="_blank" onclick="google.script.host.close();">認証ページを開く</a>' +
    '</body></html>')
    .setWidth(400)
    .setHeight(60);
  SpreadsheetApp.getUi().showModalDialog(html, 'リンクを開いて認証を行ってください');
}

/******************************************************************
関数：GetMyCompaniesID
概要：事業所の一覧をFreee APIから取得し、選択可能なポップアップとして表示する関数
******************************************************************/
function GetMyCompaniesID() {
  try {
    var accessToken = getService().getAccessToken();
    var requestUrl = "https://api.freee.co.jp/api/1/companies";
    var params = {
      method: "get",
      headers: { Authorization: "Bearer " + accessToken },
    };

    var response = UrlFetchApp.fetch(requestUrl, params);
    var responseData = JSON.parse(response.getContentText());
    var formattedData = responseData.companies.map((company) => {
      return {
        id: company.id,
        name: company.name || company.display_name || "名称未設定",
      };
    });

    SelectModal(formattedData);
  } catch (e) {
    /** エラーが発生した場合のアラート表示
    **********************************************************/
    SpreadsheetApp.getUi().alert("事業所の一覧を取得できませんでした。エラー: " + e.message);
  }
}
/** -----------------------------------------------------------------------------------------
関数：SelectModal
概要：業所一覧を取得し、このスプシで取引を送信する事業所を選択させるポップアップを出力
---------------------------------------------------------------------------------------- **/

function SelectModal(companies) {
  var html = '<style>' +

    /** 特定性を高めるためにIDセレクタを使用*************************************************/
    '#companyList li:before { position: absolute; content: ""; right: 0px; bottom: 0px; border-width: 0px 0px 15px 15px; border-style: solid; border-color: white white white #124fbd;}' +
    '#companyList li:hover, #companyList li.selected { border-left:10px solid #E91E63 !important; background-color: #fbeff7 !important; font-weight:bold; }' +
    '#companyList li:hover:before { border-color: white white white #E91E63 !important;}' +
    '.title { font-size: 18px; color: #333; padding: 10px; font-family: "Noto Sans JP"; }' +
    '</style>';

  /** 事業所一覧を<li>要素としてループで出力**/
  html += '<ul id="companyList" style="list-style-type: none; padding: 0;">'; /** IDを追加 **/
  companies.forEach(function (company, index) {
    html += '<li id="company_' + index + '" style="position:relative;cursor: pointer; margin-bottom: 1rem; padding: 0.7rem; border-left: 10px solid #4349c5; border-radius: 3px; background-color: #eff3ff; color: #333; line-height: 1.5; font-family: \'Noto Sans JP\', sans-serif;" ' +
      'onclick="selectCompany(' + company.id + ', ' + index + ')">' + company.name + '</li>';
  });
  html += '</ul>';

  html +=
    '<script>' +
    'function selectCompany(companyId, index) { ' +
    'var allLis = document.querySelectorAll("#companyList li");' +
    'allLis.forEach(function(li) { li.classList.remove("selected"); });' +
    'document.getElementById("company_" + index).classList.add("selected");' +
    'google.script.run.withSuccessHandler(function() { ' +
    'google.script.host.close(); }).setSelectedCompanyId(companyId); ' +
    '}' +
    '</script>';


  var ui = HtmlService.createHtmlOutput(html).setWidth(500).setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(ui, "事業所を選択してください！");
}
// 

/******************************************************************
関数：setSelectedCompanyId
概要：選択されたcompanyIdをselectedCompanyIdに保存
*******************************************************************/
function setSelectedCompanyId(companyId) {
  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty("selectedCompanyId", companyId);

  // 
  /** 事業所選択が完了したら他の関数を実行
  ********************************************************/
  onCompanySelected();
}
function getSelectedCompanyId() {
  var userProperties = PropertiesService.getUserProperties();
  var companyId = userProperties.getProperty("selectedCompanyId");
  return companyId ? parseInt(companyId, 10) : null; // 数値として返す
}

/******************************************************************
関数：onCompanySelected
概要：SelectModal事業所がで選択され、そのIDを取得したら連動して実行
******************************************************************/

function onCompanySelected() {
  // 他のAPI呼び出し関数を実行
  manage_Walletables(); //口座
  get_Taxes(); //税区分
  get_AccountItems();//勘定科目
  manage_Partners();//取引先
  get_Items_Register();//品目
}

/******************************************************************
関数：submit_freee
概要：SelectModal事業所がで選択され、そのIDを取得したら連動して実行
******************************************************************/

function submit_freee() {

  dealsTranscription(); //取引データを作成して
  postDeals(); //送信！

}


/******************************************************************
関数：inputClientInfo
概要：アプリ詳細画面からClient IDとClient Secretをカスタムメニューへコピペ
******************************************************************/
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

/******************************************************************
関数：showCallbackUrl
概要：認証用コールバックURLのコピペできるように出力
******************************************************************/

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
  SpreadsheetApp.getUi().showModalDialog(html, 'コールバックURL'); 
}

/******************************************************************
関数：getService
概要：freeeAPIのサービスを取得
******************************************************************/

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
/******************************************************************
関数：authCallback
概要：認証用のコールバック関数(アクセストークンの取得)
******************************************************************/

function authCallback(request) {
  var service = getService();
  var isAuthorized = service.handleCallback(request);
  if (isAuthorized) {
    return HtmlService.createHtmlOutput('認証に成功しました。このウィンドウを閉じてください。');
  } else {
    return HtmlService.createHtmlOutput('認証に失敗しました。');
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
関数：clearService
概要：認証解除
******************************************************************/
function clearService() {
  OAuth2.createService("freee")
    .setPropertyStore(PropertiesService.getUserProperties())
    .reset();

}