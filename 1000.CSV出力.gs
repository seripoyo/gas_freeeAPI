/**
 * スプレッドシートの内容をCSVファイルへ出力
 */
function output_csv_File() {
  // 出力するフォルダのIDをプロパティより取得
  var userProperties = PropertiesService.getUserProperties();
  var folderId = userProperties.getProperty('recentFolderId');

  // 出力するファイル名
  const fileName = '都道府県.csv';

  // アクティブなシートを取得
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // テキストに出力する範囲のデータを取得（A列とB列に入力されているデータを取得）
    let values = sheet.getRange(1, 1, sheet.getLastRow(), 17).getValues();

  // 改行で連結（各行の値をカンマで区切り）
  let contents = values.map(row => row.join(',')).join('\n');

  // CSVファイル書き出し
  createCsvFile_(folderId, fileName, contents);
}

/**
 * CSVファイル書き出し
 * @param {string} folderId フォルダID
 * @param {string} fileName ファイル名
 * @param {string} contents ファイルの内容
 */
function createCsvFile_(folderId, fileName, contents) {  
  // コンテンツタイプ
  const contentType = 'text/csv';
  
  // 文字コード
  const charset = 'UTF-8';

  // 出力するフォルダ
  const folder = DriveApp.getFolderById(folderId);

  // Blob を作成する
  const blob = Utilities.newBlob('', contentType, fileName).setDataFromString(contents, charset);

  // ファイルに保存
  folder.createFile(blob);
}
