

// 指定したシートと列番号に基づいて情報が入力されている最終行を取得する関数
function getLastRowInColumn(sheetName, columnNumber) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var columnData = sheet.getRange(1, columnNumber, sheet.getMaxRows(), 1).getValues();
  var lastRow = columnData.filter(String).length;
  return lastRow;
}

