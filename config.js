// Google Apps Script のウェブアプリURLを設定する(セットアップは README.md を参照)。
// 管理者の合言葉は、ここではなく、Apps Script側のスクリプトプロパティ ADMIN_PIN に保存する。
//
// 手元での確認(node scripts/serve.cjs)のときは、自動的に疑似サーバー "/gas" につながる。
//
// デモ環境: URLの末尾に ?demo=1 を付けると、デモ用のスプレッドシートにつながる(本番のデータには触れない)。
// デモでは、画面上の「今日」を demoToday として動かす(デモ用のApps Scriptにも TEST_TODAY を同じ日付で入れておく)。
;(function () {
  const isLocal = ["localhost", "127.0.0.1"].includes(location.hostname)
  const isDemo = new URLSearchParams(location.search).get("demo") === "1"
  const PRODUCTION_URL = "https://script.google.com/macros/s/AKfycbxU3fFSNOTUzJiws6Gjbq8qd9cTuRBeMNF_cEIf9OH5agRI_GL_F4b-8e2i7jpKgLdCzA/exec"
  const DEMO_URL = "https://script.google.com/macros/s/AKfycbz2lPxAeW8nCDu_8jNQMJ5MGm7qLo1ZrwJpZnLul9x3qmBX6lIHH_Nu7XcyLnEgRdYB/exec"
  window.APP_CONFIG = {
    demo: isDemo,
    demoToday: "2026-11-14",
    webAppUrl: isLocal ? "/gas" : isDemo ? DEMO_URL : PRODUCTION_URL,
  }
})()
