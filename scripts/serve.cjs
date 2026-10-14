// 開発用サーバー(確認用)。
// 静的ファイルを配信しつつ、/gas で「Google Apps Scriptの代わり」を動かす。
// apps-script/Code.gs をそのまま読み込み、メモリ上の疑似スプレッドシートで実行するので、
// 本物のスプレッドシートに触れずに、サーバー側のプログラムの動作を確認できる。
//
//   node scripts/serve.cjs [--port=8091] [--today=2026-11-05] [--admin-pin=test1234]
//
// 試験用の操作:
//   /__set?today=2026-11-10   疑似的に「今日」の日付を変える(サーバー側)
//   /__reset                  データをすべて消して最初に戻す

const http = require("http")
const fs = require("fs")
const path = require("path")
const vm = require("vm")

const args = Object.fromEntries(
  process.argv.slice(2).map((a) => {
    const [k, v] = a.replace(/^--/, "").split("=")
    return [k, v === undefined ? true : v]
  })
)
const PORT = Number(args.port || 8091)
const ROOT = path.join(__dirname, "..")
const CODE = fs.readFileSync(path.join(ROOT, "apps-script", "Code.gs"), "utf8")

// 写真の置き場(createGas を作り直しても残す)
const photos = {}
function createFile(blob) {
  const id = "photo-" + require("crypto").randomUUID()
  photos[id] = { bytes: Buffer.from(blob.bytes), mime: blob.mime }
  return { getId: () => id, setSharing() {} }
}

function createGas() {
  const sheets = {}
  const props = { ADMIN_PIN: args["admin-pin"] || null, PHOTO_URL_BASE: "/__photo?id=" }

  function makeSheet() {
    const data = []
    let maxRows = 1000
    // 本物のスプレッドシートと同じく、「=」で始まる文字列は計算式になって「#ERROR!」になる
    const parse = (v) => (typeof v === "string" && v.startsWith("=") ? "#ERROR!" : v)
    const pad = () => {
      const width = Math.max(0, ...data.map((r) => r.length))
      data.forEach((r) => {
        while (r.length < width) r.push("")
      })
    }
    const ensure = (r, c) => {
      while (data.length < r) data.push([])
      while (data[r - 1].length < c) data[r - 1].push("")
    }
    return {
      getMaxRows: () => maxRows,
      insertRowsAfter: (_row, n) => {
        maxRows += n
      },
      getLastRow: () => data.length,
      getDataRange: () => ({
        getValues: () => {
          pad()
          return data.map((r) => r.slice())
        },
      }),
      appendRow: (row) => {
        data.push(row.map(parse))
      },
      deleteRow: (n) => {
        data.splice(n - 1, 1)
      },
      getRange: (r, c, nr = 1, nc = 1) => {
        const range = {
          setNumberFormat() {
            return range
          },
          setValue(v) {
            ensure(r, c)
            data[r - 1][c - 1] = parse(v)
            return range
          },
          setValues(vs) {
            vs.forEach((row, i) =>
              row.forEach((v, j) => {
                ensure(r + i, c + j)
                data[r + i - 1][c + j - 1] = parse(v)
              })
            )
            return range
          },
          getValue() {
            ensure(r, c)
            return data[r - 1][c - 1]
          },
        }
        return range
      },
    }
  }

  const ctx = vm.createContext({
    console,
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ({
        getSheetByName: (name) => sheets[name] || null,
        insertSheet: (name) => (sheets[name] = makeSheet()),
      }),
    },
    LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: (k) => (k in props ? props[k] : null),
        setProperty: (k, v) => {
          props[k] = v
        },
      }),
    },
    // 写真の保存(メモリ上)。/__photo?id=... で取り出せる
    DriveApp: {
      Access: { ANYONE_WITH_LINK: "link" },
      Permission: { VIEW: "view" },
      createFolder: (name) => ({ getId: () => "folder-" + name, createFile: (blob) => createFile(blob) }),
      getFolderById: (id) => {
        if (!id.startsWith("folder-")) throw new Error("no folder")
        return { getId: () => id, createFile: (blob) => createFile(blob) }
      },
      getFileById: (id) => {
        if (!photos[id]) throw new Error("no file")
        return { setTrashed: () => delete photos[id] }
      },
    },
    Utilities: {
      getUuid: () => require("crypto").randomUUID(),
      base64Decode: (b64) => Array.from(Buffer.from(b64, "base64")),
      newBlob: (bytes, mime, name) => ({ bytes, mime, name }),
      formatDate: (d, tz, fmt) => {
        if (fmt !== "yyyy-MM-dd") throw new Error("fake formatDate: unsupported format " + fmt)
        const offset = tz === "Asia/Tokyo" ? 9 * 3600 * 1000 : 0
        return new Date(d.getTime() + offset).toISOString().slice(0, 10)
      },
    },
    ContentService: {
      MimeType: { JSON: "json" },
      createTextOutput: (text) => ({ text, setMimeType() { return this } }),
    },
  })

  // 「今日」を固定できるDate(省略時は本当の現在時刻)
  ctx.__now = args.today ? Date.parse(args.today + "T03:00:00Z") : null
  vm.runInContext(
    `(function () {
       const Real = Date
       class FakeDate extends Real {
         constructor(...a) { if (a.length === 0 && globalThis.__now) super(globalThis.__now); else super(...a) }
         static now() { return globalThis.__now || Real.now() }
       }
       globalThis.Date = FakeDate
     })()`,
    ctx
  )
  vm.runInContext(CODE, ctx)
  return ctx
}

let gas = createGas()

const TYPES = {
  ".html": "text/html; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".js": "application/javascript; charset=utf-8",
  ".png": "image/png",
  ".json": "application/json",
}

http
  .createServer((req, res) => {
    const url = new URL(req.url, `http://localhost:${PORT}`)

    if (url.pathname === "/__reset") {
      Object.keys(photos).forEach((k) => delete photos[k])
      const now = gas.__now
      gas = createGas()
      gas.__now = now
      res.end("reset")
      return
    }
    if (url.pathname === "/__set") {
      const d = url.searchParams.get("today")
      gas.__now = d ? Date.parse(d + "T03:00:00Z") : null
      res.end("today=" + d)
      return
    }

    if (url.pathname === "/__photo") {
      const photo = photos[url.searchParams.get("id")]
      if (!photo) {
        res.writeHead(404)
        res.end("no photo")
        return
      }
      res.writeHead(200, { "Content-Type": photo.mime })
      res.end(photo.bytes)
      return
    }

    if (url.pathname === "/gas") {
      const send = (out) => {
        res.writeHead(200, { "Content-Type": "application/json; charset=utf-8" })
        res.end(out.text)
      }
      if (req.method === "GET") {
        return send(gas.doGet({ parameter: Object.fromEntries(url.searchParams) }))
      }
      let body = ""
      req.on("data", (c) => (body += c))
      req.on("end", () => send(gas.doPost({ postData: { contents: body } })))
      return
    }

    let file = decodeURIComponent(url.pathname)
    if (file.endsWith("/")) file += "index.html"
    const full = path.join(ROOT, file)
    if (!full.startsWith(ROOT) || !fs.existsSync(full) || fs.statSync(full).isDirectory()) {
      res.writeHead(404)
      res.end("not found")
      return
    }
    res.writeHead(200, { "Content-Type": TYPES[path.extname(full)] || "application/octet-stream", "Cache-Control": "no-store" })
    fs.createReadStream(full).pipe(res)
  })
  .listen(PORT, () => console.log(`http://localhost:${PORT}  (today=${args.today || "実際の日付"})`))
