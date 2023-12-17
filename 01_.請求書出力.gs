/******************************************************************
 * 
 * Googleドライブにフォルダを作成＆中に指定したシートを複製
 * フォルダ内のシートを取引一覧シートに出力するよ
 * それ以外は保存してあるやつから情報を引っ張るよ！というやつ
 * それぞれ全部引っ張ったら取引シートに入力出来るよ。
 * 
******************************************************************/



/******************************************************************
関数：create_Folder_And_Update_Menu
概要：Googleドライブに新しくフォルダを作成
******************************************************************/

function create_Folder_And_Update_Menu() {
  var ui = SpreadsheetApp.getUi();

  // 既にフォルダが作成されているか確認
  var userProperties = PropertiesService.getUserProperties();
  var folderCreated = userProperties.getProperty('folderCreated');

  if (folderCreated) {
    ui.alert('もうフォルダを作成済みです。');
    return;
  }

  // 処理中のモードレスダイアログを表示
  var htmlOutput = HtmlService.createHtmlOutput('<p>処理中です、少しお待ちください🙏<br>このポップアップは閉じて大丈夫です。</p>')
    .setWidth(400)
    .setHeight(50);
  ui.showModelessDialog(htmlOutput, '処理中');

  var response = ui.prompt('作成する請求書フォルダ名を入力してください');

  if (response.getSelectedButton() != ui.Button.CANCEL) {
    var folderName = response.getResponseText();
    var folder = DriveApp.createFolder(folderName);
    userProperties.setProperty('recentFolderId', folder.getId());
    userProperties.setProperty('folderUrl', folder.getUrl());

    // 複製元スプレッドシートのID
    var templateId1 = '1frGGW4Awz4aIeiWfKSAar5k1qY9NyGu7fbLZ4pjCx8o';
    var templateId2 = '1B4WyWlv7HKH1eQk4pAZW_YBL2k6PBb15oQ1jcnOpqnY';

    DriveApp.getFileById(templateId1).makeCopy('サンプル', folder);
    DriveApp.getFileById(templateId2).makeCopy('複製用テンプレ', folder);

    userProperties.setProperty('folderCreated', 'true');

    ui.alert('「' + folderName + '」フォルダが作成されました。このダイアログは閉じてください。');
  } else {
    ui.alert('フォルダ作成がキャンセルされました。');
  }

  menu(); // メニューを更新
}


/******************************************************************
関数：get_SpreadsheetIds_From_Folder
概要：Googleドライブに存在する請求書を取引一覧シートに出力
******************************************************************/
function get_SpreadsheetIds_From_Folder(folderId) {
  // IDをログに出力して確認
  // Logger.log('Fetching spreadsheets from folder ID: ' + folderId);

  var folder = null;

  try {
    folder = DriveApp.getFolderById(folderId);
  } catch (e) {
    // エラー情報をログに出力
    Logger.log('Error fetching folder: ' + e.toString());
    // エラーを再投げして、呼び出し元に通知
    throw new Error('Error fetching folder with ID: ' + folderId);
  }

  var files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  //　複数のシートを配列として扱う
  var spreadsheetIds = [];

  while (files.hasNext()) {
    var file = files.next();
    spreadsheetIds.push(file.getId());
  }

  // スプシが重複した場合は削除
  return Array.from(new Set(spreadsheetIds));
}

/******************************************************************
関数：copy_Data_From_MultipleSheets
概要：特定のフォルダ内の複数のスプレッドシートからデータを集めて、
      "取引一覧"シートにコピーする。
******************************************************************/

function copy_Data_From_MultipleSheets() {
  // ユーザー設定のプロパティからフォルダIDを取得
  var userProperties = PropertiesService.getUserProperties();
  var folderId = userProperties.getProperty('recentFolderId');
  // Logger.log(folderId);

  // フォルダIDが設定されていない場合、エラーを投げる
  if (!folderId) {
    throw new Error('まだフォルダが作成されていないため、先にフォルダを作成してください。');
  }

  // 指定されたフォルダ内のスプレッドシートのIDを取得
  var sourceSpreadsheetIds = get_SpreadsheetIds_From_Folder(folderId);

  // 現在アクティブなスプレッドシートを取得し、"取引一覧"シートを探す
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dstSheet = ss.getSheetByName("取引一覧");
  if (!dstSheet) {
    throw new Error('"取引一覧"という名前のシートが見つかりません。');
  }

  // コピー開始位置（2行目）を設定
  var nextRow = 2;

  // 各スプレッドシートIDに対して処理を行う
  sourceSpreadsheetIds.forEach(function (spreadsheetId) {
    var srcSpreadsheet = SpreadsheetApp.openById(spreadsheetId);
    var srcSheets = srcSpreadsheet.getSheets();

    srcSheets.forEach(function (srcSheet) {
      var sheetName = srcSheet.getName();
      // 特定のシート名以外は処理しない
      if (sheetName !== "テンプレ" && sheetName !== "インボイス対応テンプレ" && sheetName !== "プロフィール") {
        // Logger.log('Processing sheet: ' + sheetName);
        
        var initialRow = nextRow;

        // 請求書フォーマットから取引一覧フォーマットへ値を入力
        dstSheet.getRange("A" + nextRow).setValue("収入"); //収支区分

        // 取引内容が確定した日（=発生日）を yyyy-mm-dd 形式で取得し設定
        var issueDate = formatDate(srcSheet.getRange("P4").getValue());

        dstSheet.getRange("C" + nextRow).setValue(issueDate);

        var T14Value = srcSheet.getRange("T14").getValue();
        var M15Value = srcSheet.getRange("M15").getValue();

        // 入金日が支払期限と一致しているか否かを確認し、適切な値を設定
        if (T14Value !== "" && T14Value !== M15Value) {
          // T14が入力されており、M15と異なる場合はT14の値を yyyy-mm-dd 形式で使用
          dstSheet.getRange("O" + nextRow).setValue(formatDate(T14Value));
        } else {
          // T14が空、またはM15と同じ場合はM15の値を yyyy-mm-dd 形式で使用
          dstSheet.getRange("O" + nextRow).setValue(formatDate(M15Value));
        }

        // 項目の概要
        // ------------------------------------------------------------------------------------------
        dstSheet.getRange("P" + nextRow).setValue("現金"); //決済した口座
        dstSheet.getRange("L" + nextRow).setValue(srcSheet.getRange("C6").getValue()); //売上の概要（件名）

        dstSheet.getRange("E" + nextRow).setValue(srcSheet.getRange("A3").getValue()); //請求先（クライアント名）
        dstSheet.getRange("F" + nextRow).setValue(srcSheet.getRange("T3").getValue()); //勘定科目。デフォルトは売上高
        dstSheet.getRange("H" + nextRow).setValue(srcSheet.getRange("D15").getValue()); //合計の請求額


        // ------------------------------------------------------------------------------------------
        // V19とAA19のセルの値をチェックして処理
        var tax10Percent = srcSheet.getRange("V19").getValue(); // 10%税額
        var tax8Percent = srcSheet.getRange("AA19").getValue(); // 8%税額

        // 10%税率の処理
        if (tax10Percent !== "" && parseFloat(tax10Percent) !== 0) {
          dstSheet.getRange("L" + nextRow).setValue(srcSheet.getRange("C6").getValue()); // 売上の概要
          dstSheet.getRange("F" + nextRow).setValue(srcSheet.getRange("T3").getValue()); // 勘定科目
          dstSheet.getRange("J" + nextRow).setValue(tax10Percent); // 10%税額
          var total10Percent = parseFloat(tax10Percent) + parseFloat(srcSheet.getRange("V20").getValue());
          dstSheet.getRange("H" + nextRow).setValue(total10Percent); // 10%対象額の合計
          dstSheet.getRange("I" + nextRow).setValue("税込"); // 税計算区分
          dstSheet.getRange("G" + nextRow).setValue("課税売上10%");
          nextRow += 1;
        }

        // 8%税率の処理
        if (tax8Percent !== "" && parseFloat(tax8Percent) !== 0) {
          var salesSummary = srcSheet.getRange("C6").getValue() + " - 軽減税率対象";
          dstSheet.getRange("L" + nextRow).setValue(salesSummary); // 売上の概要
          dstSheet.getRange("F" + nextRow).setValue(srcSheet.getRange("T3").getValue()); // 勘定科目
          dstSheet.getRange("J" + nextRow).setValue(tax8Percent); // 8%税額
          var total8Percent = parseFloat(tax8Percent) + parseFloat(srcSheet.getRange("AA20").getValue());
          dstSheet.getRange("H" + nextRow).setValue(total8Percent); // 8%対象額の合計
          dstSheet.getRange("I" + nextRow).setValue("税込"); // 税計算区分
          dstSheet.getRange("G" + nextRow).setValue("課税売上8%（軽）");
          nextRow += 1;
        }

        // 以下はマイナスの金額として出力得
        // ------------------------------------------------------------------------------------------

        dstSheet.getRange("H" + nextRow).setValue(srcSheet.getRange("L32").getValue() * -1);
        dstSheet.getRange("F" + nextRow).setValue("事業主貸");
        dstSheet.getRange("G" + nextRow).setValue("対象外");
        dstSheet.getRange("L" + nextRow).setValue("源泉所得税"); //勘定科目に源泉徴収税を追加

        if (srcSheet.getRange("C35").getValue() != "") {
          nextRow += 1;

          dstSheet.getRange("H" + nextRow).setValue(srcSheet.getRange("C35").getValue() * -1);
          dstSheet.getRange("F" + nextRow).setValue("事業主貸");
          dstSheet.getRange("G" + nextRow).setValue("対象外");
          dstSheet.getRange("L" + nextRow).setValue("サービス使用手数料");
        }

        var total = 0;
        for (var j = initialRow; j <= nextRow; j++) {
          total += dstSheet.getRange("H" + j).getValue();
        }
        dstSheet.getRange("Q" + initialRow).setValue(total);
        // Logger.log('Data copied to row ' + nextRow);

        nextRow++;
      }
    });
  });
}
