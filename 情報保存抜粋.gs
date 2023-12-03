
// /******************************************************************
//  * 
//  * 勘定科目データの取得
//  * 
//  * 出力例：保存した勘定科目: [{"id":343670330,"name":"売上高"},{"id":343670305,"name":"事業主貸"},{"id":343670315,"name":"前受金"},{"id":343670305,"name":"事業主貸"},{"id":343670330,"name":"売上高"},{"id":343670305,"name":"事業主貸"},{"id":343670330,"name":"売上高"},{"id":343670305,"name":"事業主貸"},{"id":343670330,"name":"売上高"}
//  * 
//  * 取引シートに転記する情報：品目列に記載されている項目と合致するnameのidをitem_idとして転記
//  * 
// ******************************************************************/
// // 保存した品目データを取得する関数
// function getSavedItemsData() {
//   var userProperties = PropertiesService.getUserProperties();
//   var itemsDataString = userProperties.getProperty("itemsData");
//   return itemsDataString ? JSON.parse(itemsDataString) : [];
// }
// /******************************************************************
//  * 
//  * 取引先データの呼び出し関数 
//  * 
//  * 出力例：保存した取引先一覧: [{"partner_id":23416140,"name":"株式会社EPARKリラク&エステ"},{"partner_id":23416387,"name":"株式会社GEKIRIN"},{"partner_id":23416491,"name":"株式会社ブランディングテクノロジー"},{"partner_id":23416616,"name":"株式会社IKUSA"},{"partner_id":32135586,"name":"forword design"}
//  * 
//  * 取引シートに転記する情報：取引先列に記載されている項目とnameと合致する際のpartner_id
//  * 
//  ******************************************************************/

// function savePartnersData(partners) {

//   // 配列を作成し、要素を格納（IDを整数に変換）
//   var partnersData = partners.map(function (partner) {
//     return {
//       partner_id: parseInt(partner.id, 10), // IDを整数に変換
//       name: partner.name
//     };
//   });

//   // 取引先データをユーザープロパティに保存
//   var userProperties = PropertiesService.getUserProperties();
//   userProperties.setProperty("partnersData", JSON.stringify(partnersData));
//   Logger.log("保存した取引先一覧: " + JSON.stringify(partnersData));
// }

// /******************************************************************
//  * 取引先データの呼び出し関数 
//  ******************************************************************/

// function saved_PartnersData() {
//   var userProperties = PropertiesService.getUserProperties();
//   var partnersDataString = userProperties.getProperty("partnersData");

//   if (partnersDataString) {
//     var partnersData = JSON.parse(partnersDataString);
//     Logger.log("取得した取引先データ: ");
//     Logger.log(partnersData);
//     return partnersData;
//   } else {
//     Logger.log("保存された取引先データはありません。");
//     return []; // データがない場合は空の配列を返す
//   }
// }

// /******************************************************************
//  * 
//  * 概要：口座一覧取得
//  * 
//  * 出力例：保存した口座一覧: [{"from_walletable_id":"2428169","name":"現金","from_walletable_type":"wallet","bank_id":null,"walletable_balance":-1354039,"last_balance":0},{"from_walletable_id":"5629615","name":"P列に記載されている項目","from_walletable_type":"wallet","bank_id":null,"walletable_balance":0,"last_balance":10000}
//  * 
//  * 取引シートに転記する情報：決済口座列に記載されている項目と合致するnameのfrom_walletable_idとfrom_walletable_type

// ******************************************************************/
// function manage_Walletables() {
//   var freeeApp = getService();
//   var accessToken = freeeApp.getAccessToken();
//   var companyId = getSelectedCompanyId();
//   var requestUrl = "https://api.freee.co.jp/api/1/walletables?company_id=" + companyId + "&with_balance=true";
//   var headers = { "Authorization": "Bearer " + accessToken };

//   var options = {
//     "method": "get",
//     "headers": headers
//   };

//   // APIからデータ取得
//   var res = UrlFetchApp.fetch(requestUrl, options).getContentText();
//   var walletablesResponse = JSON.parse(res);
//   var walletables = walletablesResponse.walletables;

//   // データ処理
//   var processedWalletables = walletables.map(function (walletable) {
//     return {
//       from_walletable_id: parseInt(walletable.id, 10).toString(), // IDを整数に変換して文字列化
//       name: walletable.name,
//       from_walletable_type: walletable.type,
//       bank_id: walletable.bank_id ? parseInt(walletable.bank_id, 10).toString() : null, // bank_idがnullでなければ整数に変換
//       walletable_balance: parseInt(walletable.walletable_balance, 10), // walletable_balanceを整数に変換
//       last_balance: parseInt(walletable.last_balance, 10) // last_balanceを整数に変換
//     };
//   });

//   // データをユーザープロパティに保存
//   var userProperties = PropertiesService.getUserProperties();
//   userProperties.setProperty("walletablesData", JSON.stringify(processedWalletables));
//   // Logger.log("保存した口座一覧: " + JSON.stringify(processedWalletables));
// }


// /******************************************************************
// 概要：税区分一覧取得&保存

// 出力例：保存した税区分: [{"tax_code":"129","name_ja":"課税売上10%","default_Tax_Ccode":34},{"tax_code":"156","name_ja":"課税売上8%（軽）","default_Tax_Ccode":34},{"tax_code":"101","name_ja":"課税売上8%","default_Tax_Ccode":34},{"tax_code":"21","name_ja":"課税売上","default_Tax_Ccode":34}

// 取引シートに転記する情報：税区分列に記載されている項目と合致するname_jaのtax_code
// ******************************************************************/
// function get_Taxes() {
//   var freeeApp = getService();
//   var accessToken = freeeApp.getAccessToken();
//   var companyId = getSelectedCompanyId();
//   var requestUrl = "https://api.freee.co.jp/api/1/taxes/companies/" + companyId;
//   var headers = { "Authorization": "Bearer " + accessToken };
//   var options = { "method": "get", "headers": headers };

//   var response = UrlFetchApp.fetch(requestUrl, options).getContentText();
//   var taxesResponse = JSON.parse(response);
//   var taxes = taxesResponse.taxes;

//   // 税区分一覧のIDを整数に変換して配列に格納
//   var taxesData = taxes.map(function (tax) {
//     return {
//       tax_code: parseInt(tax.code, 10).toString(), // IDを整数に変換して文字列化
//       name_ja: tax.name_ja,
//       // default_Tax_Ccode: 34,
//     };
//   });

//   // 税区分データをユーザープロパティに保存
//   var userProperties = PropertiesService.getUserProperties();
//   userProperties.setProperty("taxesData", JSON.stringify(taxesData));
//   Logger.log("保存した税区分: " + JSON.stringify(taxesData));

// }
// /******************************************************************
//  * 
//  * 概要：勘定科目一覧の取得
//  * 
//  * 出力例：保存した取引先一覧: [出力例：保存した取引先一覧: [{"partner_id":23416140,"name":"株式会社EPARKリラク&エステ"},{"partner_id":23416387,"name":"株式会社GEKIRIN"},{"partner_id":23416491,"name":"株式会社ブランディングテクノロジー"},{"partner_id":23416616,"name":"株式会社IKUSA"},{"partner_id":32135586,"name":"forword design"}
//  * 
//  * 取引シートに転記する情報：勘定項目列に記載されている項目と合致するnameのaccount_item_id
//  * 
//  * ただしget_Taxesにてtax_codeが存在しないときはdefaultTaxIdも転記
//  * 
// ******************************************************************/
// function get_AccountItems() {
//   var freeeApp = getService();
//   var accessToken = freeeApp.getAccessToken();
//   var companyId = getSelectedCompanyId();
//   var requestUrl = "https://api.freee.co.jp/api/1/account_items?company_id=" + companyId;
//   var headers = { "Authorization": "Bearer " + accessToken };
//   var options = { "method": "get", "headers": headers };

//   var response = UrlFetchApp.fetch(requestUrl, options).getContentText();
//   var accountItemsResponse = JSON.parse(response);
//   var accountItems = accountItemsResponse.account_items;

//   // スプレッドシートのF列のデータを取得
//   var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("売上履歴");
//   var lastRowInFColumn = getLastRowInColumn("売上履歴", 6);
//   var fColumnData = sheet.getRange(2, 6, lastRowInFColumn - 1).getValues();

//   // 合致する勘定科目のnameとidを保存する配列
//   var matchingAccountItems = [];

//   fColumnData.forEach(function (fValue) {
//     var matchingAccountItem = accountItems.find(accountItem => accountItem.name === fValue[0]);
//     if (matchingAccountItem) {
//       matchingAccountItems.push({
//         account_item_id: matchingAccountItem.id,
//         name: matchingAccountItem.name,
//         defaultTaxId: matchingAccountItem.default_tax_id
//       });
//     }
//   });

//   // 合致する勘定科目をプロパティサービスに保存
//   var userProperties = PropertiesService.getUserProperties();
//   userProperties.setProperty("matchingAccountItems", JSON.stringify(matchingAccountItems));

//   // 確認のためにログに出力
//   Logger.log("保存した勘定科目: " + JSON.stringify(matchingAccountItems));
// }