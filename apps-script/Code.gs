// わたし推し活(旧: 自分満たしカード) — Googleスプレッドシートをデータベース代わりに使うAPI
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
// 毎日の記録(1日に何件でも)。スタンプ方式だった頃の posts シートは使わない
const RECORDS_SHEET = "records"
const STUDENT_HEADERS = ["id", "display_name", "start_date", "stamp_dates", "achieved_count", "created_at", "updated_at"]
const RECORD_HEADERS = ["id", "student_id", "display_name", "date", "text", "senses", "feeling", "reactions", "hidden", "created_at", "updated_at"]

// チャレンジ全体の開始日(固定)と日数。いつ登録しても、この日が1日目になる。
const PROGRAM_START_DATE = "2026-11-01"
const PROGRAM_LENGTH = 30

const MAX_TEXT_LENGTH = 300
const MAX_FEELING_LENGTH = 300
// 五感(画面側の app.js の SENSES と合わせる)。記録には「smell,taste」のように保存する
const SENSE_TYPES = ["smell", "touch", "taste", "hearing", "sight"]

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
        case "saveRecord":
          return respond(saveRecord(body.student_id, body.record_id, body.date, body.text, body.senses, body.feeling))
        case "deleteRecord":
          return respond(deleteRecord(body.student_id, body.record_id))
        case "react":
          return respond(reactToRecord(body.student_id, body.post_id, body.type))
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
          return respond(hideRecord(body.pin, body.post_id, body.hidden))
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

function recordsSheet() {
  return getSheet(RECORDS_SHEET, RECORD_HEADERS)
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

function parseSenses(value) {
  const list = Array.isArray(value) ? value : String(value || "").split(",")
  const result = []
  list.forEach((s) => {
    const key = String(s).trim()
    if (SENSE_TYPES.indexOf(key) !== -1 && result.indexOf(key) === -1) result.push(key)
  })
  return result
}

// 記録を新しい順に並べる(記録した日 → 書いた時刻)
function byNewest(a, b) {
  const da = toDateStr(a.date)
  const db = toDateStr(b.date)
  if (da !== db) return da < db ? 1 : -1
  return toIsoStr(a.created_at) < toIsoStr(b.created_at) ? 1 : -1
}

// 本人の記録の一覧(振り返り用)
function myRecordsOf(studentId) {
  return readTable(recordsSheet(), RECORD_HEADERS)
    .filter((r) => String(r.student_id) === String(studentId))
    .sort(byNewest)
    .map((r) => ({
      id: r.id,
      date: toDateStr(r.date),
      text: String(r.text),
      senses: parseSenses(r.senses),
      feeling: String(r.feeling || ""),
      created_at: toIsoStr(r.created_at),
    }))
}

function formatStudent(obj, withRecords) {
  const student = {
    id: obj.id,
    display_name: obj.display_name,
    start_date: toDateStr(obj.start_date),
    created_at: toIsoStr(obj.created_at),
    updated_at: toIsoStr(obj.updated_at),
  }
  if (withRecords) student.my_records = myRecordsOf(obj.id)
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

function touchStudent(student) {
  setCell(studentsSheet(), STUDENT_HEADERS, student._row, "updated_at", new Date().toISOString())
}

// ---------------- 記録(1日に何件でも) ----------------

// recordId があれば、自分の記録の書き直し。なければ新しい記録。
function saveRecord(studentId, recordId, dateStr, rawText, rawSenses, rawFeeling) {
  const student = findStudentById(studentId)
  if (!student) return { error: "student_not_found" }

  const target = dateStr ? String(dateStr) : todayStr()
  const dateError = checkDateInRange(toDateStr(student.start_date), target)
  if (dateError) return { error: dateError }

  const text = String(rawText || "").trim()
  if (!text) return { error: "empty_text" }
  if (text.length > MAX_TEXT_LENGTH) return { error: "text_too_long" }
  const feeling = String(rawFeeling || "").trim()
  if (feeling.length > MAX_FEELING_LENGTH) return { error: "text_too_long" }
  const senses = parseSenses(rawSenses).join(",")

  const now = new Date().toISOString()
  const sheet = recordsSheet()
  if (recordId) {
    const record = readTable(sheet, RECORD_HEADERS).find((r) => String(r.id) === String(recordId))
    if (!record) return { error: "post_not_found" }
    if (String(record.student_id) !== String(student.id)) return { error: "not_yours" }
    setCell(sheet, RECORD_HEADERS, record._row, "date", target)
    setCell(sheet, RECORD_HEADERS, record._row, "text", text)
    setCell(sheet, RECORD_HEADERS, record._row, "senses", senses)
    setCell(sheet, RECORD_HEADERS, record._row, "feeling", feeling)
    setCell(sheet, RECORD_HEADERS, record._row, "updated_at", now)
  } else {
    appendTextRow(sheet, [Utilities.getUuid(), student.id, student.display_name, target, text, senses, feeling, "{}", "", now, now])
  }
  touchStudent(student)
  return formatStudent(findStudentById(studentId), true)
}

function deleteRecord(studentId, recordId) {
  const student = findStudentById(studentId)
  if (!student) return { error: "student_not_found" }
  const sheet = recordsSheet()
  const record = readTable(sheet, RECORD_HEADERS).find((r) => String(r.id) === String(recordId))
  if (!record) return { error: "post_not_found" }
  if (String(record.student_id) !== String(student.id)) return { error: "not_yours" }
  sheet.deleteRow(record._row)
  touchStudent(student)
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

  const posts = readTable(recordsSheet(), RECORD_HEADERS)
    .filter((p) => String(p.hidden) !== "1")
    .sort((a, b) => (toIsoStr(a.created_at) < toIsoStr(b.created_at) ? 1 : -1))
    .slice(0, FEED_LIMIT)
    .map((p) => {
      const reactions = parseReactions(p.reactions)
      return {
        id: p.id,
        date: toDateStr(p.date),
        text: String(p.text),
        senses: parseSenses(p.senses),
        feeling: String(p.feeling || ""),
        counts: countsOf(reactions),
        mine: mineOf(reactions, String(student.id)),
        is_mine: String(p.student_id) === String(student.id),
      }
    })
  return { posts }
}

function reactToRecord(studentId, postId, type) {
  const student = findStudentById(studentId)
  if (!student) return { error: "student_not_found" }
  if (REACTION_TYPES.indexOf(type) === -1) return { error: "invalid_reaction" }

  const sheet = recordsSheet()
  const post = readTable(sheet, RECORD_HEADERS).find((p) => String(p.id) === String(postId))
  if (!post || String(post.hidden) === "1") return { error: "post_not_found" }
  if (String(post.student_id) === String(student.id)) return { error: "own_post" }

  const reactions = parseReactions(post.reactions)
  const list = Array.isArray(reactions[type]) ? reactions[type] : []
  const sid = String(student.id)
  const idx = list.indexOf(sid)
  if (idx >= 0) list.splice(idx, 1)
  else list.push(sid)
  reactions[type] = list
  setCell(sheet, RECORD_HEADERS, post._row, "reactions", JSON.stringify(reactions))

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

  const roster = rosterEntries()
  const records = readTable(recordsSheet(), RECORD_HEADERS)
  const stats = {} // 受講生ID → { count, days }
  records.forEach((r) => {
    const key = String(r.student_id)
    if (!stats[key]) stats[key] = { count: 0, days: {} }
    stats[key].count++
    stats[key].days[toDateStr(r.date)] = true
  })
  const students = readTable(studentsSheet(), STUDENT_HEADERS).map((s) => {
    const student = formatStudent(s, false)
    const st = stats[String(s.id)] || { count: 0, days: {} }
    student.group = resolveGroup(s.display_name, roster)
    student.record_count = st.count
    student.record_days = Object.keys(st.days).length
    return student
  })
  const posts = records
    .sort((a, b) => (toIsoStr(a.created_at) < toIsoStr(b.created_at) ? 1 : -1))
    .map((p) => {
      const counts = countsOf(parseReactions(p.reactions))
      return {
        id: p.id,
        display_name: p.display_name,
        date: toDateStr(p.date),
        text: String(p.text),
        senses: parseSenses(p.senses),
        feeling: String(p.feeling || ""),
        counts,
        hidden: String(p.hidden) === "1",
      }
    })
  return { students, posts }
}

function hideRecord(pin, postId, hidden) {
  const authError = checkAdmin(pin)
  if (authError) return { error: authError }

  const sheet = recordsSheet()
  const post = readTable(sheet, RECORD_HEADERS).find((p) => String(p.id) === String(postId))
  if (!post) return { error: "post_not_found" }
  setCell(sheet, RECORD_HEADERS, post._row, "hidden", hidden ? "1" : "")
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

// 名前を「ほぼ一致」で比べるための整形。
// 全角/半角・大文字/小文字・空白・記号(〜 ♡ ★ など)・頭の「◯期」・末尾の「ちゃん/さん/くん/さま」を無視する。
function normalizeName(value) {
  let s = String(value || "")
  if (s.normalize) s = s.normalize("NFKC")
  s = s.toLowerCase().replace(/[\s　]+/g, "")
  s = s.replace(/^[0-9０-９]+期/, "")
  s = s.replace(/[〜~ー\-_・･♡♥☆★!！?？.。、,，()（）「」『』\[\]【】*＊+＋]/g, "")
  const stripped = s.replace(/(ちゃん|さん|くん|さま|様|氏)$/, "")
  if (stripped.length >= 2) s = stripped
  return s
}

// 名簿の行の一覧 [{ name: 整形した呼び名, group: "3期" }]
function rosterEntries() {
  return readTable(rosterSheet(), ROSTER_HEADERS)
    .map((r) => ({ name: normalizeName(r.display_name), group: String(r.group).trim() }))
    .filter((r) => r.name)
}

// 整形した名前が同じなら一致。違っても、どちらかがもう一方の先頭と一致していれば(2文字以上)同じ人とみなす。
// 候補が複数あるときは、いちばん長く一致したものを採る。
function resolveGroup(displayName, entries) {
  const name = normalizeName(displayName)
  if (!name) return ""
  const exact = entries.find((e) => e.name === name)
  if (exact) return exact.group
  let best = null
  entries.forEach((e) => {
    const short = e.name.length <= name.length ? e.name : name
    const long = e.name.length <= name.length ? name : e.name
    if (short.length >= 2 && long.indexOf(short) === 0) {
      if (!best || short.length > best.len) best = { len: short.length, group: e.group }
    }
  })
  return best ? best.group : ""
}

function groupOf(displayName) {
  return resolveGroup(displayName, rosterEntries())
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

// まとめて登録する。次の2つの書き方に対応する。
//   A. 1行に「呼び名,期」(区切りは、カンマ・タブ・読点)
//   B. 「3期」のような行を書いてから、その下に呼び名を1行ずつ並べる(次の「◯期」の行まで、その期になる)
function importRoster(pin, text) {
  const authError = checkAdmin(pin)
  if (authError) return { error: authError }
  let count = 0
  let currentGroup = ""
  String(text || "")
    .split(/\r?\n/)
    .forEach((rawLine) => {
      const line = rawLine.trim()
      if (!line) return
      if (/^[0-9０-９]+期$/.test(line)) {
        currentGroup = line.normalize ? line.normalize("NFKC") : line
        return
      }
      const parts = line.split(/[,\t、，]/).map((s) => s.trim())
      // 見出しの行(「名前,期」など)は飛ばす
      if (["名前", "お名前", "呼び名", "display_name", "name"].indexOf(parts[0].toLowerCase()) !== -1) return
      const name = parts[0]
      const group = parts.length >= 2 && parts[1] ? parts[1] : currentGroup
      if (!name || !group) return
      upsertRoster(name, group.slice(0, 20))
      count++
    })
  return { imported: count }
}
