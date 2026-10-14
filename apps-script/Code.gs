// 自分満たしツール — Googleスプレッドシートをデータベース代わりに使うAPI
//
// 使い方:
// 1. 新しいGoogleスプレッドシートを作成する(鬼トレとは別のファイル)
// 2. 「拡張機能」→「Apps Script」を開き、このファイルの中身を全部貼り付けて保存
// 3. 左の歯車「プロジェクトの設定」→「スクリプト プロパティ」→「スクリプト プロパティを追加」
//    - プロパティ: ADMIN_PIN
//    - 値: 管理者タブに入るための合言葉(自分で決める。この値はGitHubには載りません)
// 4. 右上の「デプロイ」→「新しいデプロイ」→種類は「ウェブアプリ」
//    - 実行するユーザー: 自分
//    - アクセスできるユーザー: 全員
// 5. デプロイ後に表示されるURL(.../exec で終わるもの)を config.js に貼る

const STUDENTS_SHEET = "students"
const POSTS_SHEET = "posts"
const STUDENT_HEADERS = ["id", "display_name", "start_date", "stamp_dates", "achieved_count", "created_at", "updated_at"]
const POST_HEADERS = ["id", "student_id", "display_name", "date", "text", "reactions", "hidden", "created_at", "updated_at"]

// チャレンジ全体の開始日(固定)と日数。いつ登録しても、この日が1日目になる。
const PROGRAM_START_DATE = "2026-11-01"
const PROGRAM_LENGTH = 30

const MAX_TEXT_LENGTH = 300

// 交流の場(期ごとのグループ)
const FEED_SHEET = "feed_posts"
const COMMENTS_SHEET = "comments"
const ROSTER_SHEET = "roster"
const FEED_HEADERS = ["id", "student_id", "display_name", "group", "kind", "text", "photo_id", "reactions", "hidden", "created_at"]
const COMMENT_HEADERS = ["id", "post_id", "parent_id", "student_id", "display_name", "group", "text", "hidden", "created_at"]
const ROSTER_HEADERS = ["display_name", "group"]
const FEED_KINDS = ["好き", "気づき"]
const MAX_POST_LENGTH = 500
const MAX_COMMENT_LENGTH = 200
const MAX_PHOTO_BYTES = 1500000
const PHOTO_FOLDER_NAME = "自分満たし写真"
const GROUP_FEED_LIMIT = 100
// リアクションの種類(画面側の app.js の REACTIONS と合わせる)
const REACTION_TYPES = ["like", "understand", "energy", "try"]
const FEED_LIMIT = 200

function doGet(e) {
  try {
    const action = e.parameter.action || "get"
    if (action === "get") return respond(getStudentById(e.parameter.id) || { error: "not_found" })
    if (action === "feed") return respond(getFeed(e.parameter.id))
    if (action === "groupFeed") return respond(getGroupFeed(e.parameter.id))
    return respond({ error: "unknown_action" })
  } catch (err) {
    return respond({ error: String(err) })
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents)
    const lock = LockService.getScriptLock()
    lock.waitLock(10000)
    try {
      switch (body.action) {
        case "registerOrResume":
          return respond(registerOrResume(body.display_name))
        case "savePost":
          return respond(savePost(body.student_id, body.date, body.text))
        case "deletePost":
          return respond(deletePost(body.student_id, body.date))
        case "react":
          return respond(reactToPost(body.student_id, body.post_id, body.type))
        case "createFeedPost":
          return respond(createFeedPost(body.student_id, body.kind, body.text, body.photo_base64, body.photo_mime))
        case "deleteFeedPost":
          return respond(deleteFeedPost(body.student_id, body.post_id))
        case "addComment":
          return respond(addComment(body.student_id, body.post_id, body.parent_id, body.text))
        case "deleteComment":
          return respond(deleteComment(body.student_id, body.comment_id))
        case "reactFeed":
          return respond(reactToFeedPost(body.student_id, body.post_id, body.type))
        case "adminFeed":
          return respond(adminFeed(body.pin))
        case "hideFeedPost":
          return respond(hideFeedPost(body.pin, body.post_id, body.hidden))
        case "hideComment":
          return respond(hideComment(body.pin, body.comment_id, body.hidden))
        case "setGroup":
          return respond(setGroup(body.pin, body.display_name, body.group))
        case "importRoster":
          return respond(importRoster(body.pin, body.text))
        case "adminList":
          return respond(adminList(body.pin))
        case "hidePost":
          return respond(hidePost(body.pin, body.post_id, body.hidden))
        default:
          return respond({ error: "unknown_action" })
      }
    } finally {
      lock.releaseLock()
    }
  } catch (err) {
    return respond({ error: String(err) })
  }
}

function respond(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON)
}

// ---------------- シート操作 ----------------

function getSheet(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  let sheet = ss.getSheetByName(name)
  if (!sheet) {
    sheet = ss.insertSheet(name)
    // 日付が勝手に日付型へ変換されたり、「=」で始まる投稿が計算式として実行されたりしないよう、
    // すべてのセルを文字列として扱う。
    sheet.getRange(1, 1, sheet.getMaxRows(), headers.length).setNumberFormat("@")
    sheet.appendRow(headers)
  }
  return sheet
}

function studentsSheet() {
  return getSheet(STUDENTS_SHEET, STUDENT_HEADERS)
}

function postsSheet() {
  return getSheet(POSTS_SHEET, POST_HEADERS)
}

// 先頭が「=」の文字列は、スプレッドシートに計算式として実行されてしまい、「#ERROR!」に化ける
// (セルを「文字列」の書式にしても防げない)。そこで、先頭が「= + - @」の文字列には、
// 見えない文字(ゼロ幅スペース)を先頭に付けてただの文字として保存し、読むときに外す。
const SAFE_PREFIX = "\u200B"

function encodeCell(value) {
  if (typeof value === "string" && (/^[=+\-@]/.test(value) || value.charAt(0) === SAFE_PREFIX)) {
    return SAFE_PREFIX + value
  }
  return value
}

function decodeCell(value) {
  if (typeof value === "string" && value.charAt(0) === SAFE_PREFIX) return value.slice(1)
  return value
}

function readTable(sheet, headers) {
  const data = sheet.getDataRange().getValues()
  const rows = []
  for (let i = 1; i < data.length; i++) {
    const obj = {}
    headers.forEach((h, c) => (obj[h] = decodeCell(data[i][c])))
    obj._row = i + 1 // 1-indexedの行番号
    rows.push(obj)
  }
  return rows
}

// 末尾に1行追加する。追加する行も先に「文字列」の書式にしてから値を入れる。
function appendTextRow(sheet, values) {
  const row = sheet.getLastRow() + 1
  if (row > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), 100)
  const range = sheet.getRange(row, 1, 1, values.length)
  range.setNumberFormat("@")
  range.setValues([values.map(encodeCell)])
}

function setCell(sheet, headers, row, header, value) {
  const range = sheet.getRange(row, headers.indexOf(header) + 1)
  range.setNumberFormat("@")
  range.setValue(encodeCell(value))
}

// ---------------- 日付 ----------------

// 公開前の動作確認用: スクリプトプロパティ TEST_TODAY(例: 2026-11-05)を入れると、
// その日を「今日」として扱う。確認が終わったら、必ずこのプロパティを削除すること。
function todayStr() {
  const test = PropertiesService.getScriptProperties().getProperty("TEST_TODAY")
  if (test && /^\d{4}-\d{2}-\d{2}$/.test(test)) return test
  return Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy-MM-dd")
}

function addDaysStr(dateStr, delta) {
  const [y, m, d] = dateStr.split("-").map(Number)
  return Utilities.formatDate(new Date(Date.UTC(y, m - 1, d + delta)), "UTC", "yyyy-MM-dd")
}

function toDateStr(value) {
  if (value instanceof Date) return Utilities.formatDate(value, "Asia/Tokyo", "yyyy-MM-dd")
  return String(value || "")
}

function toIsoStr(value) {
  if (value instanceof Date) return value.toISOString()
  return String(value || "")
}

function parseDatesCell(value) {
  if (!value) return []
  if (value instanceof Date) return [toDateStr(value)]
  return String(value).split(",").filter(Boolean)
}

// 押す日付が「今日以前で、チャレンジ期間内」かを確認する。問題なければ null を返す。
function checkDateInRange(startDate, target) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(target)) return "invalid_date"
  if (target > todayStr()) return "future_date"
  const lastDay = addDaysStr(startDate, PROGRAM_LENGTH - 1)
  if (target < startDate || target > lastDay) return "out_of_range"
  return null
}

// ---------------- 受講生 ----------------

function findStudentById(id) {
  if (!id) return null
  const rows = readTable(studentsSheet(), STUDENT_HEADERS)
  return rows.find((r) => String(r.id) === String(id)) || null
}

function findStudentByName(name) {
  const rows = readTable(studentsSheet(), STUDENT_HEADERS)
  return rows.find((r) => String(r.display_name).toLowerCase() === name.toLowerCase()) || null
}

// 本人の投稿(編集用に本文つき): { "2026-11-03": { id, text } }
function myPostsOf(studentId) {
  const result = {}
  readTable(postsSheet(), POST_HEADERS).forEach((p) => {
    if (String(p.student_id) === String(studentId)) {
      result[toDateStr(p.date)] = { id: p.id, text: String(p.text) }
    }
  })
  return result
}

function formatStudent(obj, withPosts) {
  const student = {
    id: obj.id,
    display_name: obj.display_name,
    start_date: toDateStr(obj.start_date),
    stamp_dates: parseDatesCell(obj.stamp_dates),
    achieved_count: Number(obj.achieved_count || 0),
    created_at: toIsoStr(obj.created_at),
    updated_at: toIsoStr(obj.updated_at),
  }
  if (withPosts) student.my_posts = myPostsOf(obj.id)
  return student
}

function getStudentById(id) {
  const student = findStudentById(id)
  return student ? formatStudent(student, true) : null
}

function registerOrResume(displayName) {
  const name = String(displayName || "").trim().slice(0, 40)
  if (!name) return { error: "名前を入力してください" }

  const existing = findStudentByName(name)
  if (existing) return formatStudent(existing, true)

  const id = Utilities.getUuid()
  const now = new Date().toISOString()
  appendTextRow(studentsSheet(), [id, name, PROGRAM_START_DATE, "", 0, now, now])
  return formatStudent(findStudentById(id), true)
}

function saveStampDates(student, dates) {
  const sheet = studentsSheet()
  dates.sort()
  setCell(sheet, STUDENT_HEADERS, student._row, "stamp_dates", dates.join(","))
  setCell(sheet, STUDENT_HEADERS, student._row, "achieved_count", dates.length)
  setCell(sheet, STUDENT_HEADERS, student._row, "updated_at", new Date().toISOString())
}

// ---------------- 投稿(1日1つ。書く=スタンプを押す) ----------------

function findPost(studentId, date) {
  const rows = readTable(postsSheet(), POST_HEADERS)
  return rows.find((p) => String(p.student_id) === String(studentId) && toDateStr(p.date) === date) || null
}

function savePost(studentId, dateStr, rawText) {
  const student = findStudentById(studentId)
  if (!student) return { error: "student_not_found" }

  const target = dateStr ? String(dateStr) : todayStr()
  const dateError = checkDateInRange(toDateStr(student.start_date), target)
  if (dateError) return { error: dateError }

  const text = String(rawText || "").trim()
  if (!text) return { error: "empty_text" }
  if (text.length > MAX_TEXT_LENGTH) return { error: "text_too_long" }

  const now = new Date().toISOString()
  const sheet = postsSheet()
  const existing = findPost(studentId, target)
  if (existing) {
    // 書き直し。リアクションはそのまま残す。
    setCell(sheet, POST_HEADERS, existing._row, "text", text)
    setCell(sheet, POST_HEADERS, existing._row, "updated_at", now)
  } else {
    appendTextRow(sheet, [Utilities.getUuid(), student.id, student.display_name, target, text, "{}", "", now, now])
  }

  const dates = parseDatesCell(student.stamp_dates)
  if (dates.indexOf(target) === -1) dates.push(target)
  saveStampDates(student, dates)

  return formatStudent(findStudentById(studentId), true)
}

function deletePost(studentId, dateStr) {
  const student = findStudentById(studentId)
  if (!student) return { error: "student_not_found" }

  const target = String(dateStr || "")
  const dateError = checkDateInRange(toDateStr(student.start_date), target)
  if (dateError) return { error: dateError }

  const existing = findPost(studentId, target)
  if (existing) postsSheet().deleteRow(existing._row)

  const dates = parseDatesCell(student.stamp_dates).filter((d) => d !== target)
  saveStampDates(student, dates)

  return formatStudent(findStudentById(studentId), true)
}

// ---------------- みんなの投稿(匿名)とリアクション ----------------

function parseReactions(cell) {
  try {
    const obj = JSON.parse(String(cell || "{}"))
    return obj && typeof obj === "object" ? obj : {}
  } catch (err) {
    return {}
  }
}

function countsOf(reactions) {
  const counts = {}
  REACTION_TYPES.forEach((t) => (counts[t] = Array.isArray(reactions[t]) ? reactions[t].length : 0))
  return counts
}

function mineOf(reactions, studentId) {
  return REACTION_TYPES.filter((t) => Array.isArray(reactions[t]) && reactions[t].indexOf(studentId) !== -1)
}

// 投稿者の名前・IDは返さない(匿名)。自分の投稿かどうかだけ is_mine で伝える。
function getFeed(studentId) {
  const student = findStudentById(studentId)
  if (!student) return { error: "not_found" }

  const posts = readTable(postsSheet(), POST_HEADERS)
    .filter((p) => String(p.hidden) !== "1")
    .sort((a, b) => (toIsoStr(a.created_at) < toIsoStr(b.created_at) ? 1 : -1))
    .slice(0, FEED_LIMIT)
    .map((p) => {
      const reactions = parseReactions(p.reactions)
      return {
        id: p.id,
        date: toDateStr(p.date),
        text: String(p.text),
        counts: countsOf(reactions),
        mine: mineOf(reactions, String(student.id)),
        is_mine: String(p.student_id) === String(student.id),
      }
    })
  return { posts }
}

function reactToPost(studentId, postId, type) {
  const student = findStudentById(studentId)
  if (!student) return { error: "student_not_found" }
  if (REACTION_TYPES.indexOf(type) === -1) return { error: "invalid_reaction" }

  const sheet = postsSheet()
  const post = readTable(sheet, POST_HEADERS).find((p) => String(p.id) === String(postId))
  if (!post || String(post.hidden) === "1") return { error: "post_not_found" }
  if (String(post.student_id) === String(student.id)) return { error: "own_post" }

  const reactions = parseReactions(post.reactions)
  const list = Array.isArray(reactions[type]) ? reactions[type] : []
  const sid = String(student.id)
  const idx = list.indexOf(sid)
  if (idx >= 0) list.splice(idx, 1)
  else list.push(sid)
  reactions[type] = list
  setCell(sheet, POST_HEADERS, post._row, "reactions", JSON.stringify(reactions))

  return { post_id: post.id, counts: countsOf(reactions), mine: mineOf(reactions, sid) }
}

// ---------------- 管理者 ----------------

// 合言葉はスクリプトプロパティ ADMIN_PIN に保存する(コードには書かない)。
function checkAdmin(pin) {
  const expected = PropertiesService.getScriptProperties().getProperty("ADMIN_PIN")
  if (!expected) return "admin_pin_not_set"
  if (String(pin || "") !== String(expected)) return "wrong_pin"
  return null
}

function adminList(pin) {
  const authError = checkAdmin(pin)
  if (authError) return { error: authError }

  const groups = rosterMap()
  const students = readTable(studentsSheet(), STUDENT_HEADERS).map((s) => {
    const student = formatStudent(s, false)
    student.group = groups[String(s.display_name).trim().toLowerCase()] || ""
    return student
  })
  const posts = readTable(postsSheet(), POST_HEADERS)
    .sort((a, b) => (toIsoStr(a.created_at) < toIsoStr(b.created_at) ? 1 : -1))
    .map((p) => {
      const counts = countsOf(parseReactions(p.reactions))
      return {
        id: p.id,
        display_name: p.display_name,
        date: toDateStr(p.date),
        text: String(p.text),
        counts,
        hidden: String(p.hidden) === "1",
      }
    })
  return { students, posts }
}

function hidePost(pin, postId, hidden) {
  const authError = checkAdmin(pin)
  if (authError) return { error: authError }

  const sheet = postsSheet()
  const post = readTable(sheet, POST_HEADERS).find((p) => String(p.id) === String(postId))
  if (!post) return { error: "post_not_found" }
  setCell(sheet, POST_HEADERS, post._row, "hidden", hidden ? "1" : "")
  return { post_id: post.id, hidden: !!hidden }
}

// ================= 交流の場(期ごとのグループ) =================
// 投稿・コメント・返信はすべて匿名。見られるのは、自分と同じ期のメンバーの投稿だけ。
// 期は、スプレッドシートの roster シート(呼び名 → 期)から、受講生の名前で判定する。

function feedSheet() {
  return getSheet(FEED_SHEET, FEED_HEADERS)
}

function commentsSheet() {
  return getSheet(COMMENTS_SHEET, COMMENT_HEADERS)
}

function rosterSheet() {
  return getSheet(ROSTER_SHEET, ROSTER_HEADERS)
}

// { "呼び名(小文字)": "3期" }
function rosterMap() {
  const map = {}
  readTable(rosterSheet(), ROSTER_HEADERS).forEach((r) => {
    const name = String(r.display_name).trim().toLowerCase()
    if (name) map[name] = String(r.group).trim()
  })
  return map
}

function groupOf(displayName) {
  return rosterMap()[String(displayName || "").trim().toLowerCase()] || ""
}

function photoUrl(photoId) {
  if (!photoId) return ""
  const base = PropertiesService.getScriptProperties().getProperty("PHOTO_URL_BASE")
  return base ? base + photoId : "https://drive.google.com/thumbnail?id=" + photoId + "&sz=w1000"
}

// 写真の保存先フォルダ(最初の1回だけ作って、IDをスクリプトプロパティに覚える)
function photoFolder() {
  const props = PropertiesService.getScriptProperties()
  const saved = props.getProperty("PHOTO_FOLDER_ID")
  if (saved) {
    try {
      return DriveApp.getFolderById(saved)
    } catch (err) {
      // 消されていたら作り直す
    }
  }
  const folder = DriveApp.createFolder(PHOTO_FOLDER_NAME)
  props.setProperty("PHOTO_FOLDER_ID", folder.getId())
  return folder
}

function savePhoto(base64, mime) {
  if (["image/jpeg", "image/png", "image/webp"].indexOf(mime) === -1) return { error: "invalid_photo" }
  const bytes = Utilities.base64Decode(String(base64))
  if (bytes.length > MAX_PHOTO_BYTES) return { error: "photo_too_large" }
  const ext = mime === "image/png" ? "png" : mime === "image/webp" ? "webp" : "jpg"
  const file = photoFolder().createFile(Utilities.newBlob(bytes, mime, Utilities.getUuid() + "." + ext))
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW)
  return { id: file.getId() }
}

function trashPhoto(photoId) {
  if (!photoId) return
  try {
    DriveApp.getFileById(photoId).setTrashed(true)
  } catch (err) {
    // すでに無ければ何もしない
  }
}

// 写真を保存する権限の確認用。Apps Script のエディタでこの関数を1回「実行」して、権限を許可する。
function authorizeDrive() {
  photoFolder()
}

// 同じ期のメンバーだけが見られる、という条件を確認する。問題なければ { student, group } を返す。
function requireGroupMember(studentId) {
  const student = findStudentById(studentId)
  if (!student) return { error: "student_not_found" }
  const group = groupOf(student.display_name)
  if (!group) return { error: "no_group" }
  return { student, group }
}

function createFeedPost(studentId, kind, rawText, photoBase64, photoMime) {
  const member = requireGroupMember(studentId)
  if (member.error) return member
  if (todayStr() < PROGRAM_START_DATE) return { error: "not_open" }
  if (FEED_KINDS.indexOf(kind) === -1) return { error: "invalid_kind" }

  const text = String(rawText || "").trim()
  if (!text) return { error: "empty_text" }
  if (text.length > MAX_POST_LENGTH) return { error: "text_too_long" }

  let photoId = ""
  if (photoBase64) {
    const saved = savePhoto(photoBase64, photoMime)
    if (saved.error) return saved
    photoId = saved.id
  }

  const id = Utilities.getUuid()
  const now = new Date().toISOString()
  appendTextRow(feedSheet(), [id, member.student.id, member.student.display_name, member.group, kind, text, photoId, "{}", "", now])
  return { post_id: id }
}

function deleteFeedPost(studentId, postId) {
  const student = findStudentById(studentId)
  if (!student) return { error: "student_not_found" }
  const sheet = feedSheet()
  const post = readTable(sheet, FEED_HEADERS).find((p) => String(p.id) === String(postId))
  if (!post) return { error: "post_not_found" }
  if (String(post.student_id) !== String(student.id)) return { error: "not_yours" }

  // この投稿へのコメントと返信も一緒に消す(下の行から消して、行番号がずれないようにする)
  const csheet = commentsSheet()
  readTable(csheet, COMMENT_HEADERS)
    .filter((c) => String(c.post_id) === String(postId))
    .sort((a, b) => b._row - a._row)
    .forEach((c) => csheet.deleteRow(c._row))
  sheet.deleteRow(post._row)
  trashPhoto(post.photo_id)
  return { post_id: postId }
}

function addComment(studentId, postId, parentId, rawText) {
  const member = requireGroupMember(studentId)
  if (member.error) return member

  const text = String(rawText || "").trim()
  if (!text) return { error: "empty_text" }
  if (text.length > MAX_COMMENT_LENGTH) return { error: "text_too_long" }

  const post = readTable(feedSheet(), FEED_HEADERS).find((p) => String(p.id) === String(postId))
  if (!post || String(post.hidden) === "1") return { error: "post_not_found" }
  if (String(post.group) !== member.group) return { error: "no_group" }

  const comments = readTable(commentsSheet(), COMMENT_HEADERS)
  if (parentId) {
    // 返信は1段だけ(コメントへの返信)。返信への返信は、元のコメントにつなげる。
    const parent = comments.find((c) => String(c.id) === String(parentId) && String(c.post_id) === String(postId))
    if (!parent || String(parent.hidden) === "1") return { error: "comment_not_found" }
    if (parent.parent_id) parentId = parent.parent_id
  }

  const id = Utilities.getUuid()
  const now = new Date().toISOString()
  appendTextRow(commentsSheet(), [id, postId, parentId || "", member.student.id, member.student.display_name, member.group, text, "", now])
  return { comment_id: id }
}

function deleteComment(studentId, commentId) {
  const student = findStudentById(studentId)
  if (!student) return { error: "student_not_found" }
  const sheet = commentsSheet()
  const all = readTable(sheet, COMMENT_HEADERS)
  const comment = all.find((c) => String(c.id) === String(commentId))
  if (!comment) return { error: "comment_not_found" }
  if (String(comment.student_id) !== String(student.id)) return { error: "not_yours" }

  // 自分のコメントと、それにつながる返信を消す
  all
    .filter((c) => String(c.id) === String(commentId) || String(c.parent_id) === String(commentId))
    .sort((a, b) => b._row - a._row)
    .forEach((c) => sheet.deleteRow(c._row))
  return { comment_id: commentId }
}

function reactToFeedPost(studentId, postId, type) {
  const member = requireGroupMember(studentId)
  if (member.error) return member
  if (REACTION_TYPES.indexOf(type) === -1) return { error: "invalid_reaction" }

  const sheet = feedSheet()
  const post = readTable(sheet, FEED_HEADERS).find((p) => String(p.id) === String(postId))
  if (!post || String(post.hidden) === "1") return { error: "post_not_found" }
  if (String(post.group) !== member.group) return { error: "no_group" }
  if (String(post.student_id) === String(member.student.id)) return { error: "own_post" }

  const reactions = parseReactions(post.reactions)
  const list = Array.isArray(reactions[type]) ? reactions[type] : []
  const sid = String(member.student.id)
  const idx = list.indexOf(sid)
  if (idx >= 0) list.splice(idx, 1)
  else list.push(sid)
  reactions[type] = list
  setCell(sheet, FEED_HEADERS, post._row, "reactions", JSON.stringify(reactions))
  return { post_id: post.id, counts: countsOf(reactions), mine: mineOf(reactions, sid) }
}

// 自分の期の投稿とコメントを返す。投稿者の名前・IDは返さない(匿名)。
function getGroupFeed(studentId) {
  const student = findStudentById(studentId)
  if (!student) return { error: "not_found" }
  const group = groupOf(student.display_name)
  if (!group) return { group: "", posts: [], comments: {} }

  const sid = String(student.id)
  const posts = readTable(feedSheet(), FEED_HEADERS)
    .filter((p) => String(p.group) === group && String(p.hidden) !== "1")
    .sort((a, b) => (toIsoStr(a.created_at) < toIsoStr(b.created_at) ? 1 : -1))
    .slice(0, GROUP_FEED_LIMIT)
  const authors = {}
  posts.forEach((p) => (authors[String(p.id)] = String(p.student_id)))

  const comments = {}
  readTable(commentsSheet(), COMMENT_HEADERS)
    .filter((c) => authors[String(c.post_id)] !== undefined && String(c.hidden) !== "1")
    .sort((a, b) => (toIsoStr(a.created_at) < toIsoStr(b.created_at) ? -1 : 1))
    .forEach((c) => {
      const key = String(c.post_id)
      if (!comments[key]) comments[key] = []
      comments[key].push({
        id: c.id,
        parent_id: c.parent_id ? String(c.parent_id) : "",
        text: String(c.text),
        created_at: toIsoStr(c.created_at),
        is_mine: String(c.student_id) === sid,
        is_author: String(c.student_id) === authors[key],
      })
    })

  return {
    group,
    posts: posts.map((p) => {
      const reactions = parseReactions(p.reactions)
      return {
        id: p.id,
        kind: p.kind,
        text: String(p.text),
        photo_url: photoUrl(p.photo_id),
        created_at: toIsoStr(p.created_at),
        counts: countsOf(reactions),
        mine: mineOf(reactions, sid),
        is_mine: String(p.student_id) === sid,
      }
    }),
    comments,
  }
}

// ---------------- 交流の場: 管理者 ----------------

function adminFeed(pin) {
  const authError = checkAdmin(pin)
  if (authError) return { error: authError }

  const commentsByPost = {}
  readTable(commentsSheet(), COMMENT_HEADERS)
    .sort((a, b) => (toIsoStr(a.created_at) < toIsoStr(b.created_at) ? -1 : 1))
    .forEach((c) => {
      const key = String(c.post_id)
      if (!commentsByPost[key]) commentsByPost[key] = []
      commentsByPost[key].push({
        id: c.id,
        parent_id: c.parent_id ? String(c.parent_id) : "",
        display_name: c.display_name,
        text: String(c.text),
        hidden: String(c.hidden) === "1",
      })
    })

  const posts = readTable(feedSheet(), FEED_HEADERS)
    .sort((a, b) => (toIsoStr(a.created_at) < toIsoStr(b.created_at) ? 1 : -1))
    .map((p) => ({
      id: p.id,
      group: p.group,
      display_name: p.display_name,
      kind: p.kind,
      text: String(p.text),
      photo_url: photoUrl(p.photo_id),
      created_at: toIsoStr(p.created_at),
      counts: countsOf(parseReactions(p.reactions)),
      hidden: String(p.hidden) === "1",
      comments: commentsByPost[String(p.id)] || [],
    }))

  const roster = readTable(rosterSheet(), ROSTER_HEADERS).map((r) => ({ display_name: r.display_name, group: String(r.group) }))
  return { posts, roster }
}

function hideFeedPost(pin, postId, hidden) {
  const authError = checkAdmin(pin)
  if (authError) return { error: authError }
  const sheet = feedSheet()
  const post = readTable(sheet, FEED_HEADERS).find((p) => String(p.id) === String(postId))
  if (!post) return { error: "post_not_found" }
  setCell(sheet, FEED_HEADERS, post._row, "hidden", hidden ? "1" : "")
  return { post_id: post.id, hidden: !!hidden }
}

function hideComment(pin, commentId, hidden) {
  const authError = checkAdmin(pin)
  if (authError) return { error: authError }
  const sheet = commentsSheet()
  const comment = readTable(sheet, COMMENT_HEADERS).find((c) => String(c.id) === String(commentId))
  if (!comment) return { error: "comment_not_found" }
  setCell(sheet, COMMENT_HEADERS, comment._row, "hidden", hidden ? "1" : "")
  return { comment_id: comment.id, hidden: !!hidden }
}

// 名簿(呼び名 → 期)に1人ぶん登録する。すでにあれば上書き。期を空にすると、グループなしになる。
function upsertRoster(name, group) {
  const sheet = rosterSheet()
  const key = String(name).trim().toLowerCase()
  const existing = readTable(sheet, ROSTER_HEADERS).find((r) => String(r.display_name).trim().toLowerCase() === key)
  if (existing) setCell(sheet, ROSTER_HEADERS, existing._row, "group", group)
  else appendTextRow(sheet, [String(name).trim(), group])
}

function setGroup(pin, displayName, group) {
  const authError = checkAdmin(pin)
  if (authError) return { error: authError }
  const name = String(displayName || "").trim()
  if (!name) return { error: "名前を入力してください" }
  upsertRoster(name, String(group || "").trim().slice(0, 20))
  return { display_name: name, group: String(group || "").trim() }
}

// 「呼び名,期」を1行ずつ貼り付けて、まとめて登録する(区切りは、カンマ・タブ・読点)。
function importRoster(pin, text) {
  const authError = checkAdmin(pin)
  if (authError) return { error: authError }
  let count = 0
  String(text || "")
    .split(/\r?\n/)
    .forEach((line) => {
      const parts = line.split(/[,\t、，]/).map((s) => s.trim())
      if (parts.length < 2 || !parts[0] || !parts[1]) return
      // 見出しの行(「名前,期」など)は飛ばす
      if (["名前", "お名前", "呼び名", "display_name", "name"].indexOf(parts[0].toLowerCase()) !== -1) return
      upsertRoster(parts[0], parts[1].slice(0, 20))
      count++
    })
  return { imported: count }
}
