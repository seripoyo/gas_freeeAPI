
/******************************************************************
function name |dealsTranscription
summary       |入力された取引をPOST用の取引シートに転写する
******************************************************************/

function dealsTranscription() {
  var sourceSheetName = "売上履歴";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName(sourceSheetName);
  var targetSheet = ss.getSheetByName("取引");

  // getLastRowNumber関数を使って各列の最終行を確認
  var lastRowCol1 = getLastRowNumber(1, sourceSheetName);
  var lastRowCol6 = getLastRowNumber(6, sourceSheetName);
  var lastRowCol15 = getLastRowNumber(15, sourceSheetName);

  // 実際にデータが存在する最終行を取得
  var sourceSheetLastRow = Math.max(lastRowCol1, lastRowCol6, lastRowCol15);

  if (sourceSheetLastRow < 2) {
    Logger.log("データがありません。");
    return;
  }
  var numRows = sourceSheetLastRow - 1; // 1行目はヘッダー行として除外

  // ソースシートからデータを取得
  var sourceSheetValues = sourceSheet.getRange(2, 1, numRows, 17).getValues();
  var companyId = getSelectedCompanyId();


  // var partnersSheet = ss.getSheetByName("取引先一覧");
  // var partnersLsatRow = getLastRowNumber(3, "取引先一覧");
  // if (partnersLsatRow != 0) {
  //   var partnersValues = partnersSheet.getRange(1, 1, partnersLsatRow, 4).getValues();
  // };
  // var accountItemsSheet = ss.getSheetByName("勘定科目一覧");
  // var accountItemsLastRow = getLastRowNumber(2, "勘定科目一覧");
  // var accountItemsValues = accountItemsSheet.getRange(1, 1, accountItemsLastRow, 7).getValues();

  // var taxesCodesSheet = ss.getSheetByName("税区分コード一覧");
  // var taxesCodesLastRow = getLastRowNumber(1, "税区分コード一覧");
  // var taxesCodesValues = taxesCodesSheet.getRange(1, 1, taxesCodesLastRow, 3).getValues();

  // var itemsSheet = ss.getSheetByName("品目一覧");
  // var itemsLastRow = getLastRowNumber(3, "品目一覧");
  // var itemsValues = itemsSheet.getRange(1, 1, itemsLastRow, 5).getValues();

  // var walletablesSheet = ss.getSheetByName("口座一覧");
  // var walletablesLastRow = getLastRowNumber(2, "口座一覧");
  // var walletable_values = walletablesSheet.getRange(1, 1, walletablesLastRow, 6).getValues();

 // データの初期化
  targetSheet.getRange(2, 1, targetSheet.getLastRow(), targetSheet.getLastColumn()).clear();

  //メモタグ列の書式を設定
  targetSheet.getRange(2, 11).setNumberFormat('@');

  //転記
 // 以後の転記処理...
  for (var i = 0; i < sourceSheetValues.length; i++) {
    var row = sourceSheetValues[i];

    // companyId
    targetSheet.getRange(i + 2, 1).setValue(companyId);

    // typeを設定（収入/支出に基づいてincome/expenseを設定）
    var typeValue = row[0] == "収入" ? "income" : "expense";
    targetSheet.getRange(i + 2, 4).setValue(typeValue);

    //発生日
    targetSheet.getRange(i + 2, 2).setValue(sourceSheetValues[i][2]);

    //決済期日
    targetSheet.getRange(i + 2, 3).setValue(sourceSheetValues[i][3]);

    /*******************************************************
      //取引先
      var partnerName = sourceSheetValues[ i ][ 4 ];
      for ( var e = 0 ; e < partnersLsatRow ; e++ ) {
        if ( partnersValues[ e ][ 2 ] == partnerName ) {
          targetSheet.getRange( i + 2 , 5 ).setValue( partnersValues[ e ][ 0 ] ); 
        };
      };
  
      //ref_number
      targetSheet.getRange( i + 2 , 6 ).setValue( sourceSheetValues[ i ][ 1 ] ); 
      
      //勘定科目
      var accountItemName = sourceSheetValues[ i ][ 5 ];
      for ( var e = 0 ; e < accountItemsLastRow ; e++ ) {
        if ( accountItemsValues[ e ][ 1 ] == accountItemName ) {
          targetSheet.getRange( i + 2 , 7 ).setValue( accountItemsValues[ e ][ 0 ] ); 
          targetSheet.getRange( i + 2 , 8 ).setValue( accountItemsValues[ e ][ 5 ] ); 
          break;
        };
      };
      
      //税計算区分
      var taxCodeName = sourceSheetValues[ i ][ 6 ];
      if ( taxCodeName != "" ) {
        for ( var e = 0 ; e < taxesCodesLastRow ; e++ ) {
          if ( taxesCodesValues[ e ][ 2 ] == taxCodeName ) {
            targetSheet.getRange( i + 2 , 8 ).setValue( taxesCodesValues[ e ][ 0 ] ); 
            break;
          };
        };
      };
      
      //品目
      var itemName = sourceSheetValues[ i ][ 11 ];
      for ( var e = 0 ; e < itemsLastRow ; e++ ) {
        if ( itemsValues[ e ][ 2 ] == itemName ) {
          targetSheet.getRange( i + 2 , 9 ).setValue( itemsValues[ e ][ 0 ] ); 
          break;
        };
      };
      
      //section_id
      var sectionName = sourceSheetValues[ i ][ 12 ];
      for ( var e = 0 ; e < sectionsLastRow ; e++ ) {
        if ( sectionsValues[ e ][ 2 ] == sectionName ) {
          targetSheet.getRange( i + 2 , 10 ).setValue( sectionsValues[ e ][ 0 ] ); 
          break;
        };
      };
      
      //tag_id
      var tagsName = sourceSheetValues[ i ][ 13 ].split( "," );
      var tagId = [];
      for ( var t = 0 ; t < tagsName.length ; t++ ) {
        var tagName = tagsName[ t ];
        
        for ( var e = 0 ; e < tagsLastRow ; e++ ) {
          if ( tagsValues[ e ][ 2 ] == tagName ) {
            tagId.push( tagsValues[ e ][ 0 ] ); 
            break;
          };
        };
        targetSheet.getRange( i + 2 , 11 ).setValue( String( tagId ) );
      };
      ******************************************************************/
    // 金額(amount)
      targetSheet.getRange( i + 2 , 12 ).setValue( sourceSheetValues[ i ][ 7 ] ); 

    // 税額(vat)
      targetSheet.getRange( i + 2 , 13 ).setValue( sourceSheetValues[ i ][ 9 ] ); 

    // 決済金額(amount)
      targetSheet.getRange( i + 2 , 18 ).setValue( sourceSheetValues[ i ][ 16 ] ); 

    var issueDate = formatDate(row[2]);
    targetSheet.getRange(i + 2, 2).setValue(issueDate);

    // targetSheet.getRange(i + 2, 12).setValue(sourceSheetValues[i][7]);

    //備考
    // targetSheet.getRange(i + 2, 13).setValue(sourceSheetValues[i][10]);

    //決済日
    // targetSheet.getRange(i + 2, 15).setValue(sourceSheetValues[i][14]);
     var settlementDate = formatDate(row[14]);
    targetSheet.getRange(i + 2, 15).setValue(settlementDate);

    /******************************************************************
    //from_walletable_type,from_walletable_id
    var walletableName = sourceSheetValues[ i ][ 15 ];
    for ( var e = 0 ; e < walletablesLastRow ; e++ ) {
      if ( walletable_values[ e ][ 1 ] == walletableName ) {
        targetSheet.getRange( i + 2 , 15 ).setValue( walletable_values[ e ][ 3 ] ); 
        targetSheet.getRange( i + 2 , 16 ).setValue( walletable_values[ e ][ 0 ] );
        break;
      };
    };
    
    //amount_payments
    targetSheet.getRange( i + 2 , 17 ).setValue( sourceSheetValues[ i ][ 16 ] ); 
  } ******************************************************************/


  }

}

/******************************************************************
function name |formatDate()
summary       |日付をyyyy-mm-dd形式にフォーマット
******************************************************************/
function formatDate(date) {
  if (date) {
    return Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return '';
}

/******************************************************************
function name |getLastRowNumber()
summary       |値の入っている最終行を返す（なぜかgetLastRowだとが関係ない最終行が返る）
******************************************************************/

function getLastRowNumber(column, sheetname) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetname);
  var lastRow = sheet.getLastRow();
  var values = sheet.getRange(1, column, lastRow, 1).getValues();

  for (var i = values.length - 1; i >= 0; i--) {
    if (values[i][0] != "") {
      return i + 1;
    }
  }
  return 0;
}