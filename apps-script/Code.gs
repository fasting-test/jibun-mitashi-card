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
// photo_id は後から足した列(既存のシートには、見出しを自動で書き足す)
const RECORD_HEADERS = ["id", "student_id", "display_name", "date", "text", "senses", "feeling", "reactions", "hidden", "created_at", "updated_at", "photo_id"]

// チャレンジ全体の開始日(固定)と日数。いつ登録しても、この日が1日目になる。
const PROGRAM_START_DATE = "2026-11-01"
const PROGRAM_LENGTH = 30

const MAX_TEXT_LENGTH = 300
const MAX_FEELING_LENGTH = 300
// 五感(画面側の app.js の SENSES と合わせる)。記録には「smell,taste」のように保存する
const SENSE_TYPES = ["smell", "touch", "taste", "hearing", "sight"]

// 記録へのコメントと、期(グループ)の名簿。みんなの記録は、同じ期のメンバーだけで見せ合う。
// (交流タブの投稿に使っていた feed_posts シートは、もう使わない。comments シートの post_id は記録のID)
const COMMENTS_SHEET = "comments"
const ROSTER_SHEET = "roster"
const COMMENT_HEADERS = ["id", "post_id", "parent_id", "student_id", "display_name", "group", "text", "hidden", "created_at"]
const ROSTER_HEADERS = ["display_name", "group"]
const MAX_COMMENT_LENGTH = 200
const MAX_PHOTO_BYTES = 1500000
const PHOTO_FOLDER_NAME = "自分満たし写真"
// リアクションの種類(画面側の app.js の REACTIONS と合わせる)
const REACTION_TYPES = ["like", "understand", "energy", "try"]
const FEED_LIMIT = 300
// 名簿で期を「運営」にした人は、すべての期の記録を見て、いいね・コメントできる。運営の記録は、すべての期に表示される。
// なりすまし防止のため、運営の名前でログインするときだけ、管理者の合言葉(ADMIN_PIN)が必要。
const STAFF_GROUP = "運営"
// 0期・1期・2期は、まとめて1つのグループ「0〜2期」にする(画面側の app.js の GROUPS と合わせる)
const ALUMNI_GROUP = "0〜2期"

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
          return respond(registerOrResume(body.display_name, body.pin))
        case "saveRecord":
          return respond(
            saveRecord(body.student_id, body.record_id, body.date, body.text, body.senses, body.feeling, body.photo_base64, body.photo_mime, body.remove_photo)
          )
        case "deleteRecord":
          return respond(deleteRecord(body.student_id, body.record_id))
        case "react":
          return respond(reactToRecord(body.student_id, body.post_id, body.type))
        case "addComment":
          return respond(addComment(body.student_id, body.post_id, body.parent_id, body.text))
        case "deleteComment":
          return respond(deleteComment(body.student_id, body.comment_id))
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
  } else if (sheet.getLastColumn() < headers.length) {
    // 後から列を足したときは、足りない見出しを右に書き足す
    const from = sheet.getLastColumn() + 1
    const missing = headers.slice(from - 1)
    sheet.getRange(1, from, sheet.getMaxRows(), missing.length).setNumberFormat("@")
    sheet.getRange(1, from, 1, missing.length).setValues([missing])
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

// 本人の記録の一覧(振り返り用)。もらったリアクションの数と、コメントも一緒に返す
function myRecordsOf(studentId) {
  const sid = String(studentId)
  const records = readTable(recordsSheet(), RECORD_HEADERS).filter((r) => String(r.student_id) === sid)
  const authors = {}
  records.forEach((r) => (authors[String(r.id)] = sid))
  const comments = visibleCommentsFor(authors, sid, groupLookup(rosterEntries()))
  return records.sort(byNewest).map((r) => ({
    id: r.id,
    date: toDateStr(r.date),
    text: String(r.text),
    senses: parseSenses(r.senses),
    feeling: String(r.feeling || ""),
    photo_url: photoUrl(r.photo_id),
    counts: countsOf(parseReactions(r.reactions)),
    comments: comments[String(r.id)] || [],
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
  if (withRecords) {
    student.group = groupOf(obj.display_name)
    student.my_records = myRecordsOf(obj.id)
  }
  return student
}

function getStudentById(id) {
  const student = findStudentById(id)
  return student ? formatStudent(student, true) : null
}

function registerOrResume(displayName, pin) {
  const name = String(displayName || "").trim().slice(0, 40)
  if (!name) return { error: "名前を入力してください" }
  if (groupOf(name) === STAFF_GROUP) {
    if (!pin) return { error: "staff_pin_required" }
    const authError = checkAdmin(pin)
    if (authError) return { error: authError }
  }

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
// 写真: photoBase64 があれば添える(書き直しなら差し替え)。removePhoto なら外す。どちらもなければ、今の写真のまま。
function saveRecord(studentId, recordId, dateStr, rawText, rawSenses, rawFeeling, photoBase64, photoMime, removePhoto) {
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
  let record = null
  if (recordId) {
    record = readTable(sheet, RECORD_HEADERS).find((r) => String(r.id) === String(recordId))
    if (!record) return { error: "post_not_found" }
    if (String(record.student_id) !== String(student.id)) return { error: "not_yours" }
  }

  let photoId = ""
  if (photoBase64) {
    const saved = savePhoto(photoBase64, photoMime)
    if (saved.error) return saved
    photoId = saved.id
  }

  if (record) {
    const oldPhotoId = String(record.photo_id || "")
    if (photoId || removePhoto) {
      setCell(sheet, RECORD_HEADERS, record._row, "photo_id", photoId)
      trashPhoto(oldPhotoId)
    }
    setCell(sheet, RECORD_HEADERS, record._row, "date", target)
    setCell(sheet, RECORD_HEADERS, record._row, "text", text)
    setCell(sheet, RECORD_HEADERS, record._row, "senses", senses)
    setCell(sheet, RECORD_HEADERS, record._row, "feeling", feeling)
    setCell(sheet, RECORD_HEADERS, record._row, "updated_at", now)
  } else {
    appendTextRow(sheet, [Utilities.getUuid(), student.id, student.display_name, target, text, senses, feeling, "{}", "", now, now, photoId])
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
  deleteCommentsOf(recordId)
  sheet.deleteRow(record._row)
  trashPhoto(record.photo_id)
  touchStudent(student)
  return formatStudent(findStudentById(studentId), true)
}

// ---------------- みんなの記録とリアクション ----------------

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

// 記録を見られるか。同じ期どうしと、運営(すべての期を見られ、運営の記録はすべての期に出る)。
// 期が決まっていない人の記録は、誰にも出さない。
function canSee(viewerGroup, authorGroup) {
  if (!viewerGroup || !authorGroup) return false
  return viewerGroup === STAFF_GROUP || authorGroup === STAFF_GROUP || viewerGroup === authorGroup
}

// 見られる記録と、そのコメントを返す。記録した人の名前と期も付ける(受講生のIDは返さない)。
// 自分の記録かどうかは is_mine で伝える。期が決まっていない人には、何も返さない。
function getFeed(studentId) {
  const student = findStudentById(studentId)
  if (!student) return { error: "not_found" }
  const roster = rosterEntries()
  const info = resolveGroup(student.display_name, roster)
  const group = info.group
  // 確認待ち(名簿に似た名前はあるが確定していない)かどうかも伝える
  if (!group) return { group: "", pending: !!info.candidate, posts: [], comments: {} }

  const sid = String(student.id)
  const groupByName = groupLookup(roster)
  const posts = readTable(recordsSheet(), RECORD_HEADERS)
    .filter((p) => String(p.hidden) !== "1" && canSee(group, groupByName(p.display_name)))
    .sort((a, b) => (toIsoStr(a.created_at) < toIsoStr(b.created_at) ? 1 : -1))
    .slice(0, FEED_LIMIT)
  const authors = {}
  posts.forEach((p) => (authors[String(p.id)] = String(p.student_id)))

  return {
    group,
    posts: posts.map((p) => {
      const reactions = parseReactions(p.reactions)
      return {
        id: p.id,
        display_name: String(p.display_name),
        group: groupByName(p.display_name),
        date: toDateStr(p.date),
        text: String(p.text),
        senses: parseSenses(p.senses),
        feeling: String(p.feeling || ""),
        photo_url: photoUrl(p.photo_id),
        created_at: toIsoStr(p.created_at),
        counts: countsOf(reactions),
        mine: mineOf(reactions, sid),
        is_mine: String(p.student_id) === sid,
      }
    }),
    comments: visibleCommentsFor(authors, sid, groupByName),
  }
}

// 記録を見られる(リアクション・コメントできる)のは、canSee で見られる人だけ。
// 問題なければ { member, record } を返す。
function requireSameGroupRecord(studentId, recordId) {
  const member = requireGroupMember(studentId)
  if (member.error) return member
  const record = readTable(recordsSheet(), RECORD_HEADERS).find((r) => String(r.id) === String(recordId))
  if (!record || String(record.hidden) === "1") return { error: "post_not_found" }
  if (!canSee(member.group, groupOf(record.display_name))) return { error: "post_not_found" }
  return { member, record }
}

function reactToRecord(studentId, postId, type) {
  if (REACTION_TYPES.indexOf(type) === -1) return { error: "invalid_reaction" }
  const found = requireSameGroupRecord(studentId, postId)
  if (found.error) return found
  const student = found.member.student
  const post = found.record
  if (String(post.student_id) === String(student.id)) return { error: "own_post" }

  const sheet = recordsSheet()
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
    const info = resolveGroup(s.display_name, roster)
    student.group = info.group
    student.group_candidate = info.candidate
    student.group_candidate_name = info.candidate_name
    student.record_count = st.count
    student.record_days = Object.keys(st.days).length
    return student
  })
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
  const groupByName = groupLookup(roster)
  const posts = records
    .sort((a, b) => (toIsoStr(a.created_at) < toIsoStr(b.created_at) ? 1 : -1))
    .map((p) => {
      const counts = countsOf(parseReactions(p.reactions))
      return {
        id: p.id,
        display_name: p.display_name,
        group: groupByName(p.display_name),
        comments: commentsByPost[String(p.id)] || [],
        date: toDateStr(p.date),
        text: String(p.text),
        senses: parseSenses(p.senses),
        feeling: String(p.feeling || ""),
        photo_url: photoUrl(p.photo_id),
        counts,
        hidden: String(p.hidden) === "1",
      }
    })
  const rosterRows = readTable(rosterSheet(), ROSTER_HEADERS).map((r) => ({ display_name: r.display_name, group: normalizeGroup(r.group) }))
  return { students, posts, roster: rosterRows }
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

// ================= 期(グループ)とコメント =================
// みんなの記録・コメントには、名前が出る。見られるのは、自分と同じ期のメンバーと、運営の記録だけ。
// 期は、スプレッドシートの roster シート(呼び名 → 期)から、受講生の名前で判定する。

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
  s = s.replace(/^[0-9]+([~〜\-ー][0-9]+)?期/, "")
  s = s.replace(/[〜~ー\-_・･♡♥☆★!！?？.。、,，()（）「」『』\[\]【】*＊+＋]/g, "")
  const stripped = s.replace(/(ちゃん|さん|くん|さま|様|氏)$/, "")
  if (stripped.length >= 2) s = stripped
  return s
}

// 期の書き方をそろえる。全角/半角をそろえ、「0~2期」「0-2期」「1期」なども「0〜2期」にする
function normalizeGroup(value) {
  let g = String(value || "").trim()
  if (g.normalize) g = g.normalize("NFKC")
  g = g.replace(/\s+/g, "").replace(/[~〜\-ー–—]/g, "〜")
  if (/^0〜2期$/.test(g) || /^[012]期$/.test(g)) return ALUMNI_GROUP
  return g
}

// 名簿の行の一覧 [{ name: 整形した呼び名, group: "3期" }]
function rosterEntries() {
  return readTable(rosterSheet(), ROSTER_HEADERS)
    .map((r) => ({ name: normalizeName(r.display_name), raw: String(r.display_name).trim(), group: normalizeGroup(r.group) }))
    .filter((r) => r.name)
}

// 期の判定。整形した名前が名簿と同じなら、その期で確定(group)。
// 似ているだけ(どちらかがもう一方の先頭と一致、2文字以上)のときは、自動では決めず、
// 候補(candidate)として返す。管理者が確認して確定する。
function resolveGroup(displayName, entries) {
  const name = normalizeName(displayName)
  if (!name) return { group: "", candidate: "", candidate_name: "" }
  const exact = entries.find((e) => e.name === name)
  if (exact) return { group: exact.group, candidate: "", candidate_name: "" }
  let best = null
  entries.forEach((e) => {
    const short = e.name.length <= name.length ? e.name : name
    const long = e.name.length <= name.length ? name : e.name
    if (short.length >= 2 && long.indexOf(short) === 0) {
      if (!best || short.length > best.len) best = { len: short.length, group: e.group, name: e.raw }
    }
  })
  return { group: "", candidate: best ? best.group : "", candidate_name: best ? best.name : "" }
}

// 確定した期だけを返す(確認待ちは "")
function groupOf(displayName) {
  return resolveGroup(displayName, rosterEntries()).group
}

// 名前 → 確定した期。たくさんの記録の期を調べるとき用(同じ名前は1回だけ判定する)
function groupLookup(roster) {
  const cache = {}
  return (name) => {
    const key = String(name)
    if (!(key in cache)) cache[key] = resolveGroup(key, roster).group
    return cache[key]
  }
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

// 記録のコメント(非表示のものを除く)を、記録IDごとにまとめる。書いた人の名前と期も付ける。
// authors は { 記録ID: 記録した人のID }。見る人(viewerId)から見た is_mine / is_author を付ける。
function visibleCommentsFor(authors, viewerId, groupByName) {
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
        display_name: String(c.display_name),
        group: groupByName(c.display_name),
        text: String(c.text),
        created_at: toIsoStr(c.created_at),
        is_mine: String(c.student_id) === viewerId,
        is_author: String(c.student_id) === authors[key],
      })
    })
  return comments
}

// 記録を消すときに、その記録へのコメントと返信も消す(下の行から消して、行番号がずれないようにする)
function deleteCommentsOf(recordId) {
  const sheet = commentsSheet()
  readTable(sheet, COMMENT_HEADERS)
    .filter((c) => String(c.post_id) === String(recordId))
    .sort((x, y) => y._row - x._row)
    .forEach((c) => sheet.deleteRow(c._row))
}

function addComment(studentId, postId, parentId, rawText) {
  const text = String(rawText || "").trim()
  if (!text) return { error: "empty_text" }
  if (text.length > MAX_COMMENT_LENGTH) return { error: "comment_too_long" }

  const found = requireSameGroupRecord(studentId, postId)
  if (found.error) return found
  const member = found.member

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

// ---------------- 管理者: コメントと名簿 ----------------

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
function upsertRoster(name, rawGroup) {
  const group = normalizeGroup(rawGroup)
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
  return { display_name: name, group: normalizeGroup(group) }
}

// まとめて登録する。次の2つの書き方に対応する。
//   A. 1行に「呼び名,期」(区切りは、カンマ・タブ・読点)
//   B. 「3期」「0〜2期」のような行を書いてから、その下に呼び名を1行ずつ並べる(次の「◯期」の行まで、その期になる)
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
      if (/^[0-9０-９]+(\s*[~〜～\-ー－]\s*[0-9０-９]+)?期$/.test(line)) {
        currentGroup = normalizeGroup(line)
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
