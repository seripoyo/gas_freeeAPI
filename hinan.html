/******************************************************************
 * 
 * 年度毎のスプシ一覧出力→抽出対象スプシを選択→指定セルの数値が
 * 「過去の取引一覧」シートに出力され送信する時に統一される
 * 
 * 別フォーマットのスプシ請求書を使っていた人でも取引一覧を作れるよ！
 * 取引一覧を作るために、まず年度指定でスプシ一覧を出力するよ。
 * 過去のスプシから取引一覧に入れたい請求書シート名を太字にしてね。
 * 全部の請求書雛形が場合はF~L列のデフォルトセルを入力してね。
 * もし一部の請求書雛形の出力セルが異なる場合は各行で入力が必要。
 * シート名のある行で特に記入がなければデフォルトセルが出力されるよ。
 * 
******************************************************************/



/******************************************************************
 * 関数：listSpreadsheetsByYear
 * 概要：過去の年度を選択してスプレッドシートの一覧を出力する関数
 ******************************************************************/
function listSpreadsheetsByYear() {
  // 現在の年と前年を取得
  var currentYear = new Date().getFullYear();
  var lastYear = currentYear - 1;

  // SVGアイコンの定義
  var svgIcon = '<svg xmlns="http://www.w3.org/2000/svg" x="0px" y="0px" width="20" height="20" viewBox="0 0 40 40" style="fill:#000000;">' +
    '<path fill="#FFF" d="M3.385 31.269L3.385 36.808 36.615 36.808 36.615 31.269 38 31.269 38 38.192 2 38.192 2 31.269z"></path>' +
    '<path fill="#FFF" d="M21 12L21 1 19 1 19 12 14 12 20 18 26 12zM32.196 24.521L35.246 13.953 33.324 13.398 30.274 23.967 25.47 22.581 29.572 30.009 37 25.908zM7.804 24.521L4.754 13.953 6.676 13.398 9.726 23.967 14.53 22.581 10.428 30.009 3 25.908z"></path>' +
    '</svg>';

  // HTMLコンテンツの生成
  var htmlContent = '<html><head><style>' +
    '.btn { cursor : pointer; display: flex; align-items: center; box-sizing: border-box; border-radius: 0 3px 3px 0; height: 50px; background-color: #b2d5de; text-align: center; text-decoration: none ; max-width: 300px; }' +
    '.btn-icon { position: relative; border-radius: 3px 0 0 3px; background-color: #697d82; width: 50px; height: 100%; color: white; transition: 0.3s; }' +
    '.btn svg { position: absolute; inset: 0; margin: auto; fill: #fff; transition: .5s all; }' +
    '.btn span { display: inline-block; color: #353535; text-align: center;padding-left: 24px;  font-family: \'Noto Sans JP\'; }' +
    '.btn:hover  {transform:0.3s; background-color: #b2dec3; }' +
    '.btn:hover span { transform:0.3s; font-weight: bold; }' +
    '.btn:hover svg { animation: iconAnime 1s linear; ransform:0.3s;}' +
    '.btn:hover .btn-icon {background-color: #496e26; transform:0.3s;}' +
    '@keyframes iconAnime { 25% { opacity: 0.8; } 50% { opacity: 1; } 75% { opacity: 0.6; } }' +
    '</style></head><body>' +
    '<p style=" font-family: \'Noto Sans JP\'; ">確定申告したい年度を選択してください:</p>' +
    '<div class="btn" onclick="selectYear(' + lastYear + ')">' +
    '  <div class="btn-icon">' + svgIcon + '</div>' +
    '  <span>' + lastYear + '年度の一覧を出力する</span>' +
    '</div><br>' +
    '<div class="btn" onclick="selectYear(' + currentYear + ')">' +
    '  <div class="btn-icon">' + svgIcon + '</div>' +
    '  <span>' + currentYear + '年度の一覧を出力する</span>' +
    '</div>' +
    '<script>' +
    'function selectYear(year) {' +
    '  google.script.run.withSuccessHandler(function() {}).processYearSelection(year);' +
    '  setTimeout(function() { google.script.host.close(); }, 2000);' +
    '}' +
    '</script>' +
    '</body></html>';

  // HTMLダイアログを表示
  var html = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(400)
    .setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(html, '過去のスプシから請求書を抽出');
}



/******************************************************************
 * 関数：processYearSelection
 * 概要：選択された年度の間に変更されたスプレッドシートをリストアップする関数
 * 
 * @param {number} selectedYear - 選択された年度
 ******************************************************************/
function processYearSelection(selectedYear) {
  // 選択された年度の開始日と終了日を設定
  var startDate = new Date(selectedYear, 0, 1);
  var endDate = new Date(selectedYear, 11, 31);

  // 処理中を通知するアラートを表示
  var ui = SpreadsheetApp.getUi();
  ui.alert('処理中です、このまま20秒ほどお待ちください…');

  // 選択された年度内に変更されたスプレッドシートを検索
  var files = DriveApp.searchFiles(
    'mimeType = "application/vnd.google-apps.spreadsheet" and modifiedDate >= "' + formatDate(startDate) + '" and modifiedDate <= "' + formatDate(endDate) + '"'
  );

  // 移行用スプシ一覧シートを取得または作成
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('移行用スプシ一覧');
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('移行用スプシ一覧');
  }

  // 既存のデータを削除
  if (sheet.getLastRow() > 2) {
    sheet.deleteRows(3, sheet.getLastRow() - 2);
  }

  var dataToWrite = [];
  while (files.hasNext()) {
    var file = files.next();
    var fileId = file.getId();
    var lastUpdated = file.getLastUpdated();
    var url = file.getUrl();
    var name = file.getName();

    // スプレッドシートを開いて、シート名を取得
    var spreadsheet = SpreadsheetApp.openById(fileId);
    var sheetNames = spreadsheet.getSheets().map(function (sheet) {
      return sheet.getName();
    });

    if (sheetNames.length > 0) {
      sheetNames.forEach(function (sheetName, index) {
        // スプレッドシートとシートへのリンクを作成し、データを配列に追加
        var sheetUrl = url + '#gid=' + spreadsheet.getSheetByName(sheetName).getSheetId();
        var row = [
          index === 0 ? formatDate(lastUpdated) : '',
          index === 0 ? '=HYPERLINK("' + url + '","' + name + '")' : '',
          '=HYPERLINK("' + sheetUrl + '","' + sheetName + '")'
        ];
        dataToWrite.push(row);
      });
    } else {
      // 子シートがない場合
      dataToWrite.push([formatDate(lastUpdated), '=HYPERLINK("' + url + '","' + name + '")', '']);
    }
  }

  // データがあればシートに書き込む
  if (dataToWrite.length > 0) {
    sheet.getRange(3, 3, dataToWrite.length, dataToWrite[0].length).setValues(dataToWrite);
  }

  // 処理が完了したことを通知
  ui.alert('処理が完了しました。このアラートを閉じると出力されます！');
  alert.close(); // alert.close() の部分は alert がどこから来たのか不明です。必要に応じて修正してください。
}



/******************************************************************
 * 関数：transferDataToTransactionList
 * 概要：移行用スプレッドシートから取引データを抽出し、過去の取引一覧シートに転記する関数
 ******************************************************************/
function transferDataToTransactionList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dstSheet = ss.getSheetByName("過去の取引一覧");
  var listSheet = ss.getSheetByName("移行用スプシ一覧");

  // "過去の取引一覧"シートが存在しない場合、エラーをスロー
  if (!dstSheet) {
    throw new Error('"過去の取引一覧"という名前のシートが見つかりません。');
  }

  var nextRow = 2; // 開始行

  for (var i = 3; i <= listSheet.getLastRow(); i++) {
    var fontBold = listSheet.getRange(i, 5).getFontWeight();

    // フォントが太字の行を処理
    if (fontBold === 'bold') {
      var formula = listSheet.getRange(i, 5).getFormula();
      var childSheetUrl = extractUrlFromFormula(formula);
      var childSheetName = extractNameFromFormula(formula);

      // シートURLと名前が取得できた場合
      if (childSheetUrl && childSheetName) {
        try {
          var srcSpreadsheet = SpreadsheetApp.openByUrl(childSheetUrl);
          var srcSheet = srcSpreadsheet.getSheetByName(childSheetName);
          var cellReferences = [];
          var isDefaultNeeded = true;

          // 移行用スプシ一覧で指定されたセルから値を取得してセット
          for (var col = 6; col <= 12; col++) {
            var cellRef = listSheet.getRange(i, col).getValue();
            if (cellRef) {
              cellReferences.push(cellRef);
              isDefaultNeeded = false; // 少なくとも一つの列に入力がある
            } else {
              cellReferences.push(null);
            }
          }

          // すべての列が空白の場合、デフォルト値を使用
          if (isDefaultNeeded) {
            for (var col = 6; col <= 12; col++) {
              cellReferences[col - 6] = listSheet.getRange(2, col).getValue();
            }
          }

          var initialRow = nextRow; // 現在の行の開始位置を保存

          // 収支区分を設定
          dstSheet.getRange("A" + nextRow).setValue("収入");

          // 移行用スプシ一覧で指定されたセルから値を取得してセット
          dstSheet.getRange("E" + nextRow).setValue(removeSuffix(getValueFromCell(srcSheet, cellReferences[0]))); // 取引先
          dstSheet.getRange("L" + nextRow).setValue(getValueFromCell(srcSheet, cellReferences[1])); // 品目
          dstSheet.getRange("C" + nextRow).setValue(formatDate(getValueFromCell(srcSheet, cellReferences[2]))); // 発生日
          dstSheet.getRange("O" + nextRow).setValue(formatDate(getValueFromCell(srcSheet, cellReferences[3]))); // 決済日

          // 金額の整数変換とセット
          dstSheet.getRange("H" + nextRow).setValue(convertToInt(getValueFromCell(srcSheet, cellReferences[4]))); // トータル金額
          dstSheet.getRange("J" + nextRow).setValue(convertToInt(getValueFromCell(srcSheet, cellReferences[5]))); // 税額

          dstSheet.getRange("I" + nextRow).setValue("税込");
          dstSheet.getRange("G" + nextRow).setValue("課税売上10%");
          dstSheet.getRange("F" + nextRow).setValue("売上高");
          dstSheet.getRange("P" + nextRow).setValue("現金");

          nextRow += 1;

                  // 以下はマイナスの金額として出力得
        // ------------------------------------------------------------------------------------------

          // 源泉徴収額の処理（L列が空白でない場合のみ）
          if (cellReferences[6] && getValueFromCell(srcSheet, cellReferences[6])) {
            var withholdingTax = getValueFromCell(srcSheet, cellReferences[6]);
            dstSheet.getRange("H" + nextRow).setValue(withholdingTax * -1);
                    dstSheet.getRange("F" + nextRow).setValue("事業主貸");
            dstSheet.getRange("G" + nextRow).setValue("対象外");
            dstSheet.getRange("L" + nextRow).setValue("源泉所得税");
            nextRow += 1;
          }

          var total = 0;
          for (var j = initialRow; j < nextRow; j++) {
            total += convertToInt(dstSheet.getRange("H" + j).getValue());
          }
          dstSheet.getRange("Q" + initialRow).setValue(total); // 合計を整数として初めての行のQ列に設定

        } catch (e) {
          Logger.log('エラー: ' + e.message);
        }
      }
    }
  }
}
function extractUrlFromFormula(formula) {
  var urlMatch = formula.match(/=HYPERLINK\("(.*?)"/);
  return urlMatch ? urlMatch[1] : null;
}

function extractNameFromFormula(formula) {
  var nameMatch = formula.match(/,"(.*?)"\)/);
  return nameMatch ? nameMatch[1] : null;
}


