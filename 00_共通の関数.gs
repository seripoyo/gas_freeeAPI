/** --------------------------------------------------------------------
関数：getValueFromCell
概要：指定したセルの値を取得する
引数：
* sheet：セルが存在するシート
* cellReference：セルの参照
返り値： セルの値
---------------------------------------------------------------------- **/
function getValueFromCell(sheet, cellReference) {
  if (cellReference) {
    return sheet.getRange(cellReference).getValue();
  }
  return '';
}

/** --------------------------------------------------------------------
関数：formatDate
概要：日付をYYYY-MM-DD形式で返す
引数：
* dateValue：日付
返り値：
* YYYY-MM-DD形式の日付
---------------------------------------------------------------------- **/
function formatDate(dateValue) {
  if (dateValue instanceof Date) {
    var year = dateValue.getFullYear();
    var month = ('0' + (dateValue.getMonth() + 1)).slice(-2);
    var day = ('0' + dateValue.getDate()).slice(-2);
    return year + '-' + month + '-' + day;
  } else {
    return dateValue; // 日付でない場合はそのまま返す
  }
}

/** --------------------------------------------------------------------
関数：removeSuffix
概要：文字列の末尾から指定した文字列を削除する
引数：
* name：文字列
* suffix：削除する文字列
返り値：
* suffixを削除した文字列
---------------------------------------------------------------------- **/
function removeSuffix(name) {
  if (typeof name === 'string') {
    return name.replace(/(御中|様|殿)$/, '');
  }
  return name;
}

/** --------------------------------------------------------------------
関数：convertToInt
概要：数値から「円」や「￥」などの単位を削除し、整数に変換
---------------------------------------------------------------------- **/
function convertToInt(value) {
  // 数値から「円」や「￥」などの単位を削除し、整数に変換
  if (typeof value === 'string') {
    value = value.replace(/[円￥,]/g, ''); // 単位とカンマを削除
  }
  return Math.round(Number(value)); // 数値に変換し、四捨五入
}