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
// リアクションの種類(画面側の app.js の REACTIONS と合わせる)
const REACTION_TYPES = ["like", "understand", "energy", "try"]
const FEED_LIMIT = 200

function doGet(e) {
  try {
    const action = e.parameter.action || "get"
    if (action === "get") return respond(getStudentById(e.parameter.id) || { error: "not_found" })
    if (action === "feed") return respond(getFeed(e.parameter.id))
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

  const students = readTable(studentsSheet(), STUDENT_HEADERS).map((s) => formatStudent(s, false))
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
