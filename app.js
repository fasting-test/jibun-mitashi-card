// 共通ロジック(日付・通信・リアクションの種類など)
// データはGoogle Apps Script経由でGoogleスプレッドシートに保存される。

// Code.gs の PROGRAM_START_DATE / PROGRAM_LENGTH と同じ値にしておくこと
const PROGRAM_START_DATE = "2026-11-01"
const PROGRAM_LENGTH = 30
const MAX_TEXT_LENGTH = 300
const MAX_FEELING_LENGTH = 300
// デモ環境では、端末に覚える内容(ログイン中の受講生など)を本番と分ける
const IS_DEMO = !!(window.APP_CONFIG && window.APP_CONFIG.demo)
const STORAGE_PREFIX = IS_DEMO ? "jmc_demo_" : "jmc_"
const STUDENT_ID_KEY = STORAGE_PREFIX + "student_id"
const RECORD_DRAFT_KEY = STORAGE_PREFIX + "record_draft"

// 五感と気持ちの記録は、11/12のプレ講座のあとから使えるようにする
const SENSES_START_DATE = "2026-11-12"

// 五感(Code.gs の SENSE_TYPES と合わせる)
const SENSES = [
  { type: "smell", emoji: "🌿", label: "嗅覚", hint: "香り" },
  { type: "touch", emoji: "🧶", label: "触覚", hint: "肌触り・温かさ" },
  { type: "taste", emoji: "🍵", label: "味覚", hint: "味" },
  { type: "hearing", emoji: "🎵", label: "聴覚", hint: "音" },
  { type: "sight", emoji: "🌷", label: "視覚", hint: "景色・色" },
]

function senseOf(type) {
  return SENSES.find((s) => s.type === type) || null
}

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
// デモ環境では、何も付けなくても demoToday の日付で動く。
function devTodayOverride() {
  const v = new URLSearchParams(location.search).get("today")
  if (v && /^\d{4}-\d{2}-\d{2}$/.test(v)) return v
  return IS_DEMO ? window.APP_CONFIG.demoToday : null
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
  if (IS_DEMO) return null
  try {
    return localStorage.getItem(SHARED_NAME_KEY)
  } catch {
    return null
  }
}

function setSharedDisplayName(name) {
  if (IS_DEMO) return
  try {
    localStorage.setItem(SHARED_NAME_KEY, name)
  } catch {
    // 無視
  }
}

function clearSharedDisplayName() {
  if (IS_DEMO) return
  try {
    localStorage.removeItem(SHARED_NAME_KEY)
  } catch {
    // 無視
  }
}

// 書きかけの記録の自動保存(画面を閉じても消えない)
function loadRecordDraft() {
  try {
    return JSON.parse(localStorage.getItem(RECORD_DRAFT_KEY) || "null")
  } catch {
    return null
  }
}

function saveRecordDraft(draft) {
  try {
    if (draft && (draft.text || draft.feeling || (draft.senses && draft.senses.length))) {
      localStorage.setItem(RECORD_DRAFT_KEY, JSON.stringify(draft))
    } else {
      localStorage.removeItem(RECORD_DRAFT_KEY)
    }
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
    post_not_found: "この記録は見つかりませんでした",
    unknown_action: "通信エラーが発生しました",
    future_date: "まだ先の日付には書けません",
    out_of_range: "この日は期間外です",
    invalid_date: "日付が正しくありません",
    empty_text: "内容を入力してください",
    text_too_long: `${MAX_TEXT_LENGTH}文字以内で入力してください`,
    own_post: "自分の記録にはリアクションできません",
    no_group: "まだ期(グループ)の設定がありません。運営にお知らせください",
    not_open: `${formatMdWeekday(PROGRAM_START_DATE)}のオープン後に投稿できます`,
    invalid_kind: "投稿の種類を選んでください",
    invalid_photo: "写真はJPEG・PNG・WebPの画像を選んでください",
    photo_too_large: "写真が大きすぎます。別の写真でお試しください",
    comment_not_found: "このコメントは見つかりませんでした",
    not_yours: "自分の投稿・コメントだけ削除できます",
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

// recordId を渡すと書き直し、空なら新しい記録。
// photo(compressPhoto の結果)を渡すと写真を添える/差し替える。removePhoto が true なら写真を外す。
async function saveRecord(studentId, recordId, date, text, senses, feeling, photo, removePhoto) {
  const body = { action: "saveRecord", student_id: studentId, record_id: recordId || "", date, text, senses, feeling }
  if (photo) {
    body.photo_base64 = photo.base64
    body.photo_mime = photo.mime
  } else if (removePhoto) {
    body.remove_photo = true
  }
  return apiPost(body)
}

async function deleteRecord(studentId, recordId) {
  return apiPost({ action: "deleteRecord", student_id: studentId, record_id: recordId })
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

// ================= 交流の場(期ごとのグループ) =================

// 投稿の種類(Code.gs の FEED_KINDS と合わせる)
const FEED_KINDS = [
  { kind: "好き", emoji: "💛", hint: "例: 最近ハマっているもの、好きな場所や時間" },
  { kind: "気づき", emoji: "💡", hint: "例: 今日気づいたこと、やってみて感じたこと" },
]
const MAX_POST_LENGTH = 500
const MAX_COMMENT_LENGTH = 200
const COMMUNITY_DRAFT_KEY = STORAGE_PREFIX + "community_draft"

// 写真は、送る前にスマホの中で小さくする(長辺1280px・JPEG)。
// サーバー側の上限(1.5MB)に収まるよう、画質を少しずつ下げて調整する。
const PHOTO_MAX_EDGE = 1280
const PHOTO_MAX_BYTES = 1200000

// "2026-11-12T05:03:00.000Z" → "11/12(木) 14:03"(日本時間)
function formatDateTime(iso) {
  const d = new Date(iso)
  if (isNaN(d)) return ""
  const jst = new Date(d.getTime() + 9 * 60 * 60 * 1000)
  const ymd = jst.toISOString().slice(0, 10)
  return `${formatMdWeekday(ymd)} ${String(jst.getUTCHours()).padStart(2, "0")}:${String(jst.getUTCMinutes()).padStart(2, "0")}`
}

function loadCommunityDraft() {
  try {
    return JSON.parse(localStorage.getItem(COMMUNITY_DRAFT_KEY) || "null")
  } catch {
    return null
  }
}

function saveCommunityDraft(draft) {
  try {
    if (draft && draft.text) localStorage.setItem(COMMUNITY_DRAFT_KEY, JSON.stringify(draft))
    else localStorage.removeItem(COMMUNITY_DRAFT_KEY)
  } catch {
    // 無視
  }
}

function loadImageFromFile(file) {
  return new Promise((resolve, reject) => {
    const url = URL.createObjectURL(file)
    const img = new Image()
    img.onload = () => {
      URL.revokeObjectURL(url)
      resolve(img)
    }
    img.onerror = () => {
      URL.revokeObjectURL(url)
      reject(new Error("写真を読み込めませんでした。別の写真でお試しください"))
    }
    img.src = url
  })
}

// 選んだ写真を縮小して、送信用の文字列(base64)と、画面に出す用のURLを返す
async function compressPhoto(file) {
  if (!file.type.startsWith("image/")) throw new Error(apiErrorMessage("invalid_photo"))
  const img = await loadImageFromFile(file)
  const scale = Math.min(1, PHOTO_MAX_EDGE / Math.max(img.naturalWidth, img.naturalHeight))
  const canvas = document.createElement("canvas")
  canvas.width = Math.max(1, Math.round(img.naturalWidth * scale))
  canvas.height = Math.max(1, Math.round(img.naturalHeight * scale))
  const ctx = canvas.getContext("2d")
  ctx.fillStyle = "#fff" // 透明な部分は白にする(JPEGには透明がないため)
  ctx.fillRect(0, 0, canvas.width, canvas.height)
  ctx.drawImage(img, 0, 0, canvas.width, canvas.height)

  let quality = 0.85
  let dataUrl = canvas.toDataURL("image/jpeg", quality)
  while (dataUrl.length * 0.75 > PHOTO_MAX_BYTES && quality > 0.4) {
    quality -= 0.1
    dataUrl = canvas.toDataURL("image/jpeg", quality)
  }
  if (dataUrl.length * 0.75 > PHOTO_MAX_BYTES) throw new Error(apiErrorMessage("photo_too_large"))
  return { base64: dataUrl.split(",")[1], mime: "image/jpeg", previewUrl: dataUrl }
}

// 投稿・記録に添えられた写真(タップすると大きく開く)
function buildPostPhoto(url) {
  const link = document.createElement("a")
  link.href = url
  link.target = "_blank"
  link.rel = "noopener noreferrer"
  link.className = "post-photo-link"
  const img = document.createElement("img")
  img.className = "post-photo"
  img.src = url
  img.alt = "添付された写真"
  img.loading = "lazy"
  link.appendChild(img)
  return link
}

async function fetchGroupFeed(studentId) {
  return apiGet({ action: "groupFeed", id: studentId })
}

async function createFeedPost(studentId, kind, text, photo) {
  const body = { action: "createFeedPost", student_id: studentId, kind, text }
  if (photo) {
    body.photo_base64 = photo.base64
    body.photo_mime = photo.mime
  }
  return apiPost(body)
}

async function deleteFeedPost(studentId, postId) {
  return apiPost({ action: "deleteFeedPost", student_id: studentId, post_id: postId })
}

async function addComment(studentId, postId, parentId, text) {
  return apiPost({ action: "addComment", student_id: studentId, post_id: postId, parent_id: parentId || "", text })
}

async function deleteComment(studentId, commentId) {
  return apiPost({ action: "deleteComment", student_id: studentId, comment_id: commentId })
}

async function reactToFeedPost(studentId, postId, type) {
  return apiPost({ action: "reactFeed", student_id: studentId, post_id: postId, type })
}

async function adminFeed(pin) {
  return apiPost({ action: "adminFeed", pin })
}

async function adminHideFeedPost(pin, postId, hidden) {
  return apiPost({ action: "hideFeedPost", pin, post_id: postId, hidden })
}

async function adminHideComment(pin, commentId, hidden) {
  return apiPost({ action: "hideComment", pin, comment_id: commentId, hidden })
}

async function adminSetGroup(pin, displayName, group) {
  return apiPost({ action: "setGroup", pin, display_name: displayName, group })
}

async function adminImportRoster(pin, text) {
  return apiPost({ action: "importRoster", pin, text })
}
