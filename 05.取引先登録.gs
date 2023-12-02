

// 新しい取引先をログに出力する関数
function logNewPartners(registeredPartners) {
  if (registeredPartners.length > 0) {
    Logger.log(registeredPartners.length + "件を登録しました: " + registeredPartners.join(", "));
  } else {
    Logger.log("登録していない取引先はありません");
  }
}
