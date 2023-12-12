
/*********************************************************

 * カスタムメニューfreeeMenuの内容となる関数を作成
 * freeeMenuに追加される項目はfreee_menu.gsに記載

******************************************************** */

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
関数：alertAuth_First
概要：最初に諸々の権限を認証するよ
******************************************************************/
function alertAuth_First() {
  var service = getService();
  var authorizationUrl = service.getAuthorizationUrl();

  var html = HtmlService.createHtmlOutput('<html><body>' +
    '<a href="' + authorizationUrl + '" target="_blank" onclick="handleAuthClick();">認証ページを開く</a>' +
    '</body></html>')
    .setWidth(500)
    .setHeight(60);
  var ui = SpreadsheetApp.getUi();
  ui.showModalDialog(html, 'リンクを開いて認証を行ってください');
}

/******************************************************************
show_CallbackUrl_and_Applink
概要：認証用コールバックURLのコピペ→アプリ作成画面URLを自動で開く
******************************************************************/

function show_CallbackUrl_and_Applink() {
  var scriptId = ScriptApp.getScriptId();
  var callbackUrl = 'https://script.google.com/macros/d/' + scriptId + '/usercallback';
  var htmlContent = '<input id="url" value="' + callbackUrl + '" readonly style="width:100%; padding: 10px;">' +
    '<a href="#" class="btn" onclick="copyToClipboard()">' +
    '  <div class="btn-icon">' +
    '    <svg xmlns="http://www.w3.org/2000/svg" height="1em" viewBox="0 0 448 512">' +
    '      <path d="M48 96V416c0 8.8 7.2 16 16 16H384c8.8 0 16-7.2 16-16V170.5c0-4.2-1.7-8.3-4.7-11.3l33.9-33.9c12 12 18.7 28.3 18.7 45.3V416c0 35.3-28.7 64-64 64H64c-35.3 0-64-28.7-64-64V96C0 60.7 28.7 32 64 32H309.5c17 0 33.3 6.7 45.3 18.7l74.5 74.5-33.9 33.9L320.8 84.7c-.3-.3-.5-.5-.8-.8V184c0 13.3-10.7 24-24 24H104c-13.3 0-24-10.7-24-24V80H64c-8.8 0-16 7.2-16 16zm80-16v80H272V80H128zm32 240a64 64 0 1 1 128 0 64 64 0 1 1 -128 0z"></path>' + // SVGのpathをここに挿入
    '    </svg>' +
    '  </div>' +
    '  <span>コピーする</span>' +
    '</a>' +
'<script>' +
'  function copyToClipboard() {' +
'    var copyText = document.getElementById("url");' +
'    copyText.select();' +
'    document.execCommand("copy");' +
'    alert("コピーしました: " + copyText.value);' +
'    google.script.run.openLink();' + // ここで openLink を実行
'    google.script.host.close();' + // ダイアログを閉じる
'  }' +
'</script>' +
    '<style>' +
    '  body {' +
    '        overflow: hidden; ' +
    '  }' +
    '  .btn {' +
    '    display: flex; align-items: center; box-sizing: border-box; border-radius: 0 3px 3px 0; height: 45px; background-color: #eff3ff; text-align: center; text-decoration: none; margin-top:34px; max-width: 190px; ' +
    '  }' +
    '  .btn-icon {' +
    '    position: relative; border-radius: 3px 0 0 3px; background-color:#4349c5; width: 40px; height: 100%; color: white; transition: 0.3s;' +
    '  }' +
    '  .btn svg {' +
    '    position: absolute; inset: 0; margin: auto; fill: #fff; transition: .5s all;' +
    '  }' +
    '  .btn span {' +
    '    display: inline-block; width: 100px; color: #353535; text-align: center; padding-left:10px;  font-family: "Noto Sans JP"; ' +
    '  }' +
    '  .btn:hover svg {' +
    '    animation: iconAnime 1s linear;' +
    '  }' +
    '  .btn:hover span {' +
    '    font-weight:bold; ' +
    '  }' +
    '  @keyframes iconAnime {' +
    '    25% { transform: scale(1.2); }' +
    '    50% { transform: scale(1.1); }' +
    '    75% { transform: scale(1.2); }' +
    '  }' +
    '</style>';


  var html = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(500)
    .setHeight(150);
  SpreadsheetApp.getUi().showModalDialog(html, 'コールバックURLをコピーしてください');
}
/** --------------------------------------------------------------------
関数：openLink
概要：連携アプリ作成リンクを別タブで開く
---------------------------------------------------------------------- **/
function openLink() {
  var url = 'https://app.secure.freee.co.jp/developers/companies';
  var html = HtmlService.createHtmlOutput(
    '<html><script>window.open("' + url + '");' +
    'function onUserAction() {' +
    '  google.script.run.input_ClientInfo();' +
    '  google.script.host.close();' +
    '}</script>' +
    '<button onclick="onUserAction()">次へ進む</button></html>')
    .setWidth(400)
    .setHeight(60);
  SpreadsheetApp.getUi().showModalDialog(html, '連携アプリ作成画面を開いています');
}

/******************************************************************
関数：input_ClientInfo
概要：アプリ詳細画面からClient IDとClient Secretをカスタムメニューへコピペ
******************************************************************/
// カスタムダイアログを表示する関数
function input_ClientInfo() {
  var htmlContent = `
    <!DOCTYPE html>
    <html>
    <head>
        <base target="_top">
        <style>
            body {
                overflow: hidden;
            }
            .input-group {
                position: relative;
                margin: 2em 0 1em 0;
            }

            .input-group input {
                padding: 0.8em;
                width: 95%;
                outline: none;
                border: 2px solid #bcbcbc;
                border-radius: 0.5rem;
                background-color: transparent;
                font-size: 100%;
            }

            .input-group label {
                position: absolute;
                top: 50%;
                left: 0;
                transform: translateY(-50%);
                margin-left: 1em;
                color: #bcbcbc;
                pointer-events: none;
                transition: all 0.3s ease;
            }

            .input-group :is(input:focus, input:valid)~label {
                transform: translateY(-130%) scale(.8);
                margin: 0em;
                margin-left: 0.5em;
                padding: 0.4em;
                color: #37bcf8;
                background-color: #fff;
            }

            .input-group :is(input:focus, input:valid) {
                border-color: #37bcf8;
            }

            .btn {
                display: flex; align-items: center; box-sizing: border-box; border-radius:3px; height: 45px; background-color: #eff3ff; border:1px solid #4349c5; text-align: center; text-decoration: none; margin-top:34px; max-width: 190px;
            }

            .btn-icon {
                position: relative; background-color:#4349c5; width: 40px; height: 100%; color: white; transition: 0.3s;
            }

            .btn svg {
                position: absolute; inset: 0; margin: auto; fill: #fff; transition: .5s all;
            }

            .btn span {
                display: inline-block; width: 100px; color: #353535; text-align: center; padding-left:10px;  font-family: "Noto Sans JP";
            }

            .btn:hover {
                border:2px solid #4349c5;
            }

            .btn:hover svg {
                animation: iconAnime 1s linear;
            }

            .btn:hover span {
                font-weight:bold;
            }

            @keyframes iconAnime {
                25% { transform: scale(1.2); }
                50% { transform: scale(1.1); }
                75% { transform: scale(1.2); }
            }
        </style>
    </head> 
    <body>
        <form id="clientInfoForm">
            <div class="input-group">
                <input id="clientId" type="text" name="clientId">
                <label for="clientId">Client ID:</label>
            </div>
            <div class="input-group">
                <input id="clientSecret" type="text" name="clientSecret">
                <label for="clientSecret">Client Secret:</label>
            </div>
    </form>
    <a href="#" class="btn" onclick="save_ClientInfo()">
    <div class="btn-icon">
      <svg xmlns="http://www.w3.org/2000/svg" height="1em" viewBox="0 0 448 512">
        <path d="M48 96V416c0 8.8 7.2 16 16 16H384c8.8 0 16-7.2 16-16V170.5c0-4.2-1.7-8.3-4.7-11.3l33.9-33.9c12 12 18.7 28.3 18.7 45.3V416c0 35.3-28.7 64-64 64H64c-35.3 0-64-28.7-64-64V96C0 60.7 28.7 32 64 32H309.5c17 0 33.3 6.7 45.3 18.7l74.5 74.5-33.9 33.9L320.8 84.7c-.3-.3-.5-.5-.8-.8V184c0 13.3-10.7 24-24 24H104c-13.3 0-24-10.7-24-24V80H64c-8.8 0-16 7.2-16 16zm80-16v80H272V80H128zm32 240a64 64 0 1 1 128 0 64 64 0 1 1 -128 0z"></path> // SVGのpathをここに挿入
      </svg>
    </div>
    <span>保存する</span>
    </a>    

        <script>
            function save_ClientInfo() {
                var clientId = document.getElementById('clientId').value;
                var clientSecret = document.getElementById('clientSecret').value;
                google.script.run.save_ClientInfo(clientId, clientSecret);
                google.script.host.close();
            }
        </script>
    </body>
    </html>
  `;

  var html = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(600)
    .setHeight(250);
  SpreadsheetApp.getUi().showModalDialog(html, 'アプリページの情報をそのままコピペしてください！');
}


/** --------------------------------------------------------------------
関数：save_ClientInfo
概要：プロパティにクライアント情報を保存する関数
---------------------------------------------------------------------- **/
function save_ClientInfo(clientId, clientSecret) {
  PropertiesService.getUserProperties()
    .setProperty('freeeClientId', clientId)
    .setProperty('freeeClientSecret', clientSecret);
  // クライアント情報の保存後に認証を行う
  alertAuth();
}
/** --------------------------------------------------------------------
関数：alertAuth
概要：save_ClientInfo実行後にプロパティ認証リンクをダイアログで出力
---------------------------------------------------------------------- **/
function alertAuth() {
  var service = getService();
  var authorizationUrl = service.getAuthorizationUrl();

  var html = HtmlService.createHtmlOutput('<html><body>' +
    '<a href="' + authorizationUrl + '" target="_blank" onclick="google.script.host.close();">認証ページを開く</a>' +
    '</body></html>')
    .setWidth(500)
    .setHeight(60);
  var ui = SpreadsheetApp.getUi();
  ui.showModalDialog(html, 'リンクを開いて認証を行ってください');
}

/** --------------------------------------------------------------------
関数：authCallback
概要：認証用のコールバック関数→実行後にgetMyCompaniesIDを実行
---------------------------------------------------------------------- **/

function authCallback(request) {
  var service = getService();
  var isAuthorized = service.handleCallback(request);
  if (isAuthorized) {
    // 認証成功時のHTML出力にJavaScriptを追加
    return HtmlService.createHtmlOutput(
      '<html><body>認証に成功しました。このウィンドウを閉じてください。' +
      '</body></html>');
  } else {
    return HtmlService.createHtmlOutput('認証に失敗しました。');
  }
}


/******************************************************************
関数：getMyCompaniesID
概要：事業所の一覧をFreee APIから取得し、選択可能なポップアップとして表示する関数
******************************************************************/
function getMyCompaniesID() {
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
   ---------------------------------------------------------------------- **/
    SpreadsheetApp.getUi().alert("事業所の一覧を取得できませんでした。エラー: " + e.message);
  }
}



/** -----------------------------------------------------------------------------------------
関数：SelectModal
概要：業所一覧を取得し、このスプシで取引を送信する事業所を選択させるポップアップを出力
---------------------------------------------------------------------------------------- **/
function SelectModal(companies) {
  var htmlContent = '<style>' +
    // New CSS styles
    '.switch-checkbox { margin-bottom: 1rem; align-items: center; display:flex;}' +
    '.switch-checkbox input[type=checkbox] { position: relative; cursor: pointer; width: 3.5rem; height: 1.9rem; margin-top: -0.2rem; border-radius: 60px; background-color: #ddd; -webkit-appearance: none; -moz-appearance: none; appearance: none; vertical-align: middle; transition: .5s; }' +
    '.switch-checkbox input[type=checkbox]::before { position: absolute; top: .2rem; left: .2rem; width: 1.5rem; height: 1.5rem; box-sizing: border-box; border-radius: 50%; background: white; color: black; content: \'\'; transition: .3s ease; }' +
    '.switch-checkbox input[type=checkbox]:checked { background: #1ff210; }' +
    '.switch-checkbox input[type=checkbox]:checked::before { left: 1.8rem; }' +
    '.switch-checkbox label { margin-left: 1rem;     font-size: 18px; font-family: "Noto Sans JP"; margin-bottom:10px;}' +
    '</style>';

  htmlContent += '<ul id="companyList">';
  companies.forEach(function (company, index) {
    htmlContent += '<li class="switch-checkbox" id="company_' + index + '">' + 
      '<input type="checkbox" id="switch-check-' + index + '" onclick="selectCompany(' + company.id + ')">' +
      '<label for="switch-check-' + index + '">' + company.name + '</label>' +
      '</li>';
  });
  htmlContent += '</ul>';

  htmlContent +=
    '<script>' +
    'function selectCompany(companyId) { ' +
    '  google.script.run.withSuccessHandler(closeDialog).set_Selected_CompanyId(companyId); ' +
    '}' +
    'function closeDialog() { ' +
    '   google.script.run.withSuccessHandler(function() { ' +
    '   google.script.host.close(); ' +
    '  }).showAlert(); ' +
    '}' +
    '</script>';

  var ui = HtmlService.createHtmlOutput(htmlContent).setWidth(500).setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(ui, "事業所を選択してください！");
}

/** --------------------------------------------------------------------
関数：showAlert
概要：サーバーサイドでアラートを表示
---------------------------------------------------------------------- **/
function showAlert() {
  SpreadsheetApp.getUi().alert('freeeと連携完了しました！');
}

/** --------------------------------------------------------------------
関数：closeDialog
概要：ダイアログを閉じる前にサーバーサイドのアラート表示関数を呼び出す
---------------------------------------------------------------------- **/


/** --------------------------------------------------------------------
関数：set_Selected_CompanyId
概要：選択されたcompanyIdをselectedCompanyIdに保存
---------------------------------------------------------------------- **/
function set_Selected_CompanyId(companyId) {
  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty("selectedCompanyId", companyId);
}
/** --------------------------------------------------------------------
関数：getSelectedCompanyId
概要：selectedCompanyIdに保存されたcompanyIdを取得
---------------------------------------------------------------------- **/

function getSelectedCompanyId() {
  var userProperties = PropertiesService.getUserProperties();
  var companyId = userProperties.getProperty("selectedCompanyId");

  if (companyId) {
    var fixedCompanyId = parseFloat(companyId).toFixed(0);
    return parseInt(fixedCompanyId, 10);
  }


  return null; // companyIdがnullまたはundefinedの場合

}

function logSelectedCompanyId() {
  var companyId = getSelectedCompanyId(); // 修正されたcompanyIdを取得
  if (companyId !== null) {
    Logger.log("選択された事業所ID: " + companyId); // コンソールに出力
  } else {
    Logger.log("事業所IDが選択されていません"); // companyIdが取得できなかった場合
  }
}



