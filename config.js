// Google Apps Script のウェブアプリURLを設定する(セットアップは README.md を参照)。
// 管理者の合言葉は、ここではなく、Apps Script側のスクリプトプロパティ ADMIN_PIN に保存する。
//
// 手元での確認(node scripts/serve.cjs)のときは、自動的に疑似サーバー "/gas" につながる。
window.APP_CONFIG = {
  webAppUrl: ["localhost", "127.0.0.1"].includes(location.hostname)
    ? "/gas"
    : "https://script.google.com/macros/s/xxxx/exec",
}
