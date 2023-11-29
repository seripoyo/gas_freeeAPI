// ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
// 実行内容：GET
// 処理：事業所の一覧をFreee APIから取得し、選択可能なポップアップとして表示する関数
// 使用API：https://api.freee.co.jp/api/1/companies
// ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
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
    // エラーが発生した場合のアラート表示
    SpreadsheetApp.getUi().alert("事業所の一覧を取得できませんでした。エラー: " + e.message);
  }
}

// ------------------------------------------------------------------------------------------
// 実行内容：GET
// 処理：事業所一覧を取得し、このスプシで取引を送信する事業所を選択させるポップアップを出力
/**  @param {Array} companies - 表示する事業所の一覧*/
// ------------------------------------------------------------------------------------------
function SelectModal(companies) {
  var html = '<style>' +
             // 特定性を高めるためにIDセレクタを使用
             '#companyList li:before { position: absolute; content: ""; right: 0px; bottom: 0px; border-width: 0px 0px 15px 15px; border-style: solid; border-color: white white white #124fbd;}' +
             '#companyList li:hover, #companyList li.selected { border-left:10px solid #E91E63 !important; background-color: #fbeff7 !important; font-weight:bold; }' +
             '#companyList li:hover:before { border-color: white white white #E91E63 !important;}' +
             '.title { font-size: 18px; color: #333; padding: 10px; font-family: "Noto Sans JP"; }' +
             '</style>';

  // タイトル部分のスタイルを適用
  html += '<ul id="companyList" style="list-style-type: none; padding: 0;">'; // IDを追加
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
