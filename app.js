// 共通ロジック(日付・通信・リアクションの種類など)
// データはGoogle Apps Script経由でGoogleスプレッドシートに保存される。

// Code.gs の PROGRAM_START_DATE / PROGRAM_LENGTH と同じ値にしておくこと
const PROGRAM_START_DATE = "2026-11-01"
const PROGRAM_LENGTH = 30
const MAX_TEXT_LENGTH = 300
const STUDENT_ID_KEY = "jmc_student_id"

// リアクションの種類(Code.gs の REACTION_TYPES と合わせる)
const REACTIONS = [
  { type: "like", emoji: "💗", label: "いいね" },
  { type: "understand", emoji: "🤝", label: "わかる" },
  { type: "energy", emoji: "✨", label: "元気もらった" },
  { type: "try", emoji: "🌱", label: "やってみたい" },
]

// ---- 日付ユーティリティ(タイムゾーンによるズレを避けるため文字列ベースで計算) ----

function isLocalhost() {
  return ["localhost", "127.0.0.1"].includes(location.hostname)
}

// 動作確認用: URLの末尾に ?today=2026-11-05 を付けると、画面上の「今日」をその日にして見られる。
// (画面の見え方が変わるだけ。書き込みの日付の判定はサーバー側で行うので、これで不正はできない)
function devTodayOverride() {
  const v = new URLSearchParams(location.search).get("today")
  return v && /^\d{4}-\d{2}-\d{2}$/.test(v) ? v : null
}

function todayStr() {
  const override = devTodayOverride()
  if (override) return override
  const jst = new Date(Date.now() + 9 * 60 * 60 * 1000)
  return jst.toISOString().slice(0, 10)
}

function parseYmd(s) {
  const [y, m, d] = s.split("-").map(Number)
  return [y, m - 1, d]
}

function dayOffsetFromStart(startDate, targetDate = todayStr()) {
  const [sy, sm, sd] = parseYmd(startDate)
  const [ty, tm, td] = parseYmd(targetDate)
  return Math.round((Date.UTC(ty, tm, td) - Date.UTC(sy, sm, sd)) / (1000 * 60 * 60 * 24))
}

function addDays(s, delta) {
  const [y, m, d] = parseYmd(s)
  const date = new Date(Date.UTC(y, m, d))
  date.setUTCDate(date.getUTCDate() + delta)
  return date.toISOString().slice(0, 10)
}

const WEEKDAYS = ["日", "月", "火", "水", "木", "金", "土"]

// "2026-11-03" → "11/3(火)"
function formatMdWeekday(s) {
  const [y, m, d] = parseYmd(s)
  const wd = WEEKDAYS[new Date(Date.UTC(y, m, d)).getUTCDay()]
  return `${m + 1}/${d}(${wd})`
}

// ---- 敬称 ----

function withChan(name) {
  return name.endsWith("ちゃん") || name.endsWith("さん") || name.endsWith("くん") ? name : `${name}ちゃん`
}

// ---- 吹き出しで話すキャラクター(3種類からランダム) ----

const CHARACTER_IMAGES = ["images/character-1.png", "images/character-2.png", "images/character-3.png"]

function randomCharacterImage() {
  return CHARACTER_IMAGES[Math.floor(Math.random() * CHARACTER_IMAGES.length)]
}

// ---- 受講生IDの保存(この端末が「自分は誰か」を覚えておくためだけのもの。
//      本体データはスプレッドシート側にある) ----

function getStudentIdFromLocalStorage() {
  try {
    return localStorage.getItem(STUDENT_ID_KEY)
  } catch {
    return null
  }
}

function setStudentIdToLocalStorage(id) {
  try {
    localStorage.setItem(STUDENT_ID_KEY, id)
  } catch {
    // 無視
  }
}

function clearStudentIdFromLocalStorage() {
  try {
    localStorage.removeItem(STUDENT_ID_KEY)
  } catch {
    // 無視
  }
}

// ---- 名前の引き継ぎ(チャレンジシート ⇔ 自分満たし) ----
// 2つのツールは同じ fasting-test.github.io に置かれているので、端末の記憶(localStorage)を共有できる。
// 片方でログインした名前を覚えておき、もう片方の登録画面の名前欄に最初から入れる。
// (入れるだけ。登録は、もう片方の画面で「はじめる」を押したときに行われる)
const SHARED_NAME_KEY = "skipf_shared_display_name"

function getSharedDisplayName() {
  try {
    return localStorage.getItem(SHARED_NAME_KEY)
  } catch {
    return null
  }
}

function setSharedDisplayName(name) {
  try {
    localStorage.setItem(SHARED_NAME_KEY, name)
  } catch {
    // 無視
  }
}

function clearSharedDisplayName() {
  try {
    localStorage.removeItem(SHARED_NAME_KEY)
  } catch {
    // 無視
  }
}

// 書きかけの文章の自動保存(日付ごと)
function draftKey(date) {
  return `jmc_draft_${date}`
}

function loadDraft(date) {
  try {
    return localStorage.getItem(draftKey(date))
  } catch {
    return null
  }
}

function saveDraft(date, text) {
  try {
    if (text) localStorage.setItem(draftKey(date), text)
    else localStorage.removeItem(draftKey(date))
  } catch {
    // 無視
  }
}

// ---- Google Apps Script ウェブアプリとの通信 ----

function getWebAppUrl() {
  const url = window.APP_CONFIG && window.APP_CONFIG.webAppUrl
  if (!url || url.indexOf("xxxx") !== -1) {
    throw new Error("config.jsにGoogle Apps ScriptのウェブアプリURLを設定してください")
  }
  return url
}

async function apiGet(params) {
  const url = new URL(getWebAppUrl(), location.href)
  Object.entries(params).forEach(([k, v]) => url.searchParams.set(k, v))
  const res = await fetch(url.toString())
  const data = await res.json()
  if (data && data.error) throw new Error(apiErrorMessage(data.error))
  return data
}

async function apiPost(body) {
  // Content-Typeを明示しない(text/plainのままにする)ことで、
  // Google Apps Script側でのCORSプリフライトの問題を回避する。
  const res = await fetch(new URL(getWebAppUrl(), location.href).toString(), {
    method: "POST",
    body: JSON.stringify(body),
  })
  const data = await res.json()
  if (data && data.error) throw new Error(apiErrorMessage(data.error))
  return data
}

function apiErrorMessage(error) {
  const messages = {
    not_found: "データが見つかりませんでした",
    student_not_found: "データが見つかりませんでした",
    post_not_found: "この投稿は見つかりませんでした",
    unknown_action: "通信エラーが発生しました",
    future_date: "まだ先の日付には書けません",
    out_of_range: "この日はチャレンジの期間外です",
    invalid_date: "日付が正しくありません",
    empty_text: "内容を入力してください",
    text_too_long: `${MAX_TEXT_LENGTH}文字以内で入力してください`,
    own_post: "自分の投稿にはリアクションできません",
    wrong_pin: "合言葉が違います",
    admin_pin_not_set: "管理者の合言葉がまだ設定されていません(Apps Scriptのスクリプトプロパティ ADMIN_PIN を設定してください)",
  }
  return messages[error] || error
}

// ---- 受講生・投稿・リアクション ----

async function registerOrResumeStudent(displayName) {
  const name = displayName.trim().slice(0, 40)
  if (!name) throw new Error("名前を入力してください")
  return apiPost({ action: "registerOrResume", display_name: name })
}

async function getStudentById(id) {
  if (!id) return null
  const data = await apiGet({ action: "get", id })
  return data && data.id ? data : null
}

async function savePost(studentId, date, text) {
  return apiPost({ action: "savePost", student_id: studentId, date, text })
}

async function deletePost(studentId, date) {
  return apiPost({ action: "deletePost", student_id: studentId, date })
}

async function fetchFeed(studentId) {
  const data = await apiGet({ action: "feed", id: studentId })
  return data.posts || []
}

async function reactToPost(studentId, postId, type) {
  return apiPost({ action: "react", student_id: studentId, post_id: postId, type })
}

async function adminList(pin) {
  return apiPost({ action: "adminList", pin })
}

async function adminHidePost(pin, postId, hidden) {
  return apiPost({ action: "hidePost", pin, post_id: postId, hidden })
}

function fireConfetti() {
  const colors = ["#ff8a5c", "#ffc93c", "#ff6f91", "#6bcb77", "#4d96ff"]
  confetti({ particleCount: 70, spread: 75, startVelocity: 38, origin: { y: 0.65 }, colors })
  confetti({ particleCount: 30, spread: 100, startVelocity: 25, origin: { y: 0.55 }, colors, scalar: 0.7 })
}
