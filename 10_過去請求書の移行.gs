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
  // サブフォルダIDを取得
  // var userProperties = PropertiesService.getUserProperties();
  // var subFolderId = userProperties.getProperty('subFolderId');
  // var subFolder = DriveApp.getFolderById(subFolderId);

  var userProperties = PropertiesService.getUserProperties();
  var subFolderId = userProperties.getProperty('subFolderId');
    var subFolder = DriveApp.getFolderById(subFolderId);
  if (!subFolderId) {
    throw new Error("サブフォルダのIDが見つかりません。");
  }

  // '移行前'フォルダに新しいスプレッドシートを作成
  var subFolder = DriveApp.getFolderById(subFolderId);
  var newSpreadsheet = SpreadsheetApp.create('抽出した請求書一覧');
  var newSpreadsheetFile = DriveApp.getFileById(newSpreadsheet.getId());
  var newSpreadsheetUrl = newSpreadsheet.getUrl();
  subFolder.addFile(newSpreadsheetFile);
    Logger.log('新しいスプレッドシートがサブフォルダ内に作成されました: ' + newSpreadsheetUrl);

  // '移行用スプシ一覧'シートを取得
  var sourceSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = sourceSpreadsheet.getSheetByName('移行用スプシ一覧');
  if (!sourceSheet) {
    throw new Error("'移行用スプシ一覧'シートが見つかりません。");
  }

  // 3行目以降で太字になっている子シートをコピー
  for (var i = 3; i <= sourceSheet.getLastRow(); i++) {
    var fontBold = sourceSheet.getRange(i, 5).getFontWeight();
    if (fontBold === 'bold') {
      var formula = sourceSheet.getRange(i, 5).getFormula();
      var childSheetUrl = extractUrlFromFormula(formula);
      var childSheetName = extractNameFromFormula(formula);

      if (childSheetUrl && childSheetName) {
        var parentSpreadsheet = SpreadsheetApp.openByUrl(childSheetUrl);
        var parentSheetName = parentSpreadsheet.getName(); // 親シート名を取得

        var newName = parentSheetName + '_' + childSheetName; // 新しいシート名
        try {
          var sheetToCopy = parentSpreadsheet.getSheetByName(childSheetName);
          sheetToCopy.copyTo(newSpreadsheet).setName(newName);
          Logger.log('シートをコピーしました: ' + newName);
        } catch (e) {
          Logger.log('エラー: ' + e.message);
        }
      } else {
        Logger.log('有効なURLまたはシート名が見つかりません: 行 ' + i);
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
