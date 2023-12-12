function listSpreadsheetsByYear() {
  var currentYear = new Date().getFullYear();
  var lastYear = currentYear - 1;
  var htmlContent = '<html><body>' +
    '<p>年度を選択してください:</p>' +
    '<button onclick="selectYear(' + lastYear + ')">' + lastYear + '年度の一覧を出力する</button><br>' +
    '<button onclick="selectYear(' + currentYear + ')">' + currentYear + '年度の一覧を出力する</button>' +
    '<script>' +
    'function selectYear(year) {' +
    '  google.script.run.withSuccessHandler(function() { google.script.host.close(); }).processYearSelection(year);' +
    '}' +
    '</script>' +
    '</body></html>';

  var html = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(300)
    .setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(html, '年度を選択');
}

function processYearSelection(selectedYear) {
  var startDate = new Date(selectedYear, 0, 1);
  var endDate = new Date(selectedYear, 11, 31);

  // スプレッドシートの情報を取得
  var files = DriveApp.searchFiles(
    'mimeType = "application/vnd.google-apps.spreadsheet" and modifiedDate >= "' + formatDate(startDate) + '" and modifiedDate <= "' + formatDate(endDate) + '"'
  );

  // 「移行用スプシ一覧」シートを取得または作成
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('移行用スプシ一覧');
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('移行用スプシ一覧');
    // 1-2行目を設定するコードがここに必要になる場合があります
  }

  // 既存のデータをクリアするが、最初の2行は保持する
  if (sheet.getLastRow() > 2) {
    sheet.deleteRows(3, sheet.getLastRow() - 2);
  }

  var currentRow = 3; // 3行目から出力開始
  while (files.hasNext()) {
    var file = files.next();
    var lastUpdated = file.getLastUpdated();
    var url = file.getUrl();
    var fileId = file.getId();
    var name = file.getName();

    var spreadsheet = SpreadsheetApp.openById(fileId);
    var sheetNames = spreadsheet.getSheets().map(function(sheet) { 
      var sheetUrl = spreadsheet.getUrl() + '#gid=' + sheet.getSheetId();
      return ['=HYPERLINK("' + sheetUrl + '","' + sheet.getName() + '")']; 
    });

    // 親シート情報を出力
    sheet.getRange(currentRow, 3).setValue(formatDate(lastUpdated));
    sheet.getRange(currentRow, 4).setFormula('=HYPERLINK("' + url + '","' + name + '")');

    // 子シート名にリンクを付与して出力
    if (sheetNames.length > 0) {
      sheet.getRange(currentRow, 5, sheetNames.length, 1).setFormulas(sheetNames);
      currentRow += sheetNames.length;
      currentRow++; // 子シートがない場合、行を1つ進める
    }
  }
}

function copyBoldSheetsToNewSpreadsheet() {
  // 新しいスプレッドシート「抽出した請求書一覧」を作成
  var newSpreadsheet = SpreadsheetApp.create('抽出した請求書一覧');
  var newSpreadsheetUrl = newSpreadsheet.getUrl();
  Logger.log('新しいスプレッドシートが作成されました: ' + newSpreadsheetUrl);

  // '移行用スプシ一覧'シートを取得
  var sourceSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = sourceSpreadsheet.getSheetByName('移行用スプシ一覧');
  if (!sourceSheet) {
    throw new Error("'移行用スプシ一覧'シートが見つかりません。");
  }

  // E列（子シート名とリンク）のデータとフォントの太字情報を取得
  var range = sourceSheet.getRange(3, 5, sourceSheet.getLastRow() - 2, 1);
  var values = range.getValues();
  var fonts = range.getFontWeights();
  Logger.log('子シートのデータを取得しました。');

  // 3行目以降で太字になっている子シートをコピー
  for (var i = 0; i < values.length; i++) {
    if (fonts[i][0] === 'bold' && values[i][0]) {
      var formula = sourceSheet.getRange(i + 3, 5).getFormula(); // 数式を取得
      Logger.log('行 ' + (i + 3) + ' の数式: ' + formula); // 数式をログに記録

      var childSheetUrl = extractUrlFromFormula(formula); // 数式からURLを抽出
      var childSheetName = extractNameFromFormula(formula); // 数式からシート名を抽出

      if (childSheetUrl && childSheetName) {
        try {
          // コピー元のシートを取得し、新しいスプレッドシートにコピー
          var sheetToCopy = SpreadsheetApp.openByUrl(childSheetUrl).getSheetByName(childSheetName);
          sheetToCopy.copyTo(newSpreadsheet).setName(childSheetName);
          Logger.log('シートをコピーしました: ' + childSheetName);
        } catch (e) {
          Logger.log('エラー: ' + e.message);
        }
      } else {
        Logger.log('有効なURLまたはシート名が見つかりません: 行 ' + (i + 3));
      }
    }
  }

  Logger.log('処理完了');
}

function extractUrlFromFormula(formula) {
  var urlMatch = formula.match(/=HYPERLINK\("(.*?)"/);
  return urlMatch ? urlMatch[1] : null;
}

function extractNameFromFormula(formula) {
  var nameMatch = formula.match(/,"(.*?)"\)/);
  return nameMatch ? nameMatch[1] : null;
}
