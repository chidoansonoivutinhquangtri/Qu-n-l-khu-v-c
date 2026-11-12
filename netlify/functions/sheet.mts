// Cầu nối webapp ↔ Google Sheet. Có 3 chế độ, tự chọn theo cấu hình:
//
// 0) GOOGLE_SERVICE_ACCOUNT_JSON đã đặt (khuyên dùng): đọc + GHI THẲNG lên Google Sheet qua Google Sheets API,
//    không cần Apps Script, không trễ. Đăng nhập theo sheet "Tài khoản". Xem netlify/lib/gsheets.mts.
// 1) GOOGLE_SCRIPT_URL đã đặt (Apps Script "Ứng dụng web"): đăng nhập + đọc + GHI 2 chiều, cập nhật ~20 giây.
// 2) Chưa đặt: đọc trực tiếp từ link Google Sheet "Xuất bản lên web" (pubhtml) – đăng nhập theo sheet
//    "Tài khoản", xem dữ liệu, CHƯA ghi được từ web lên sheet. Google có thể trễ ~5 phút với bản xuất bản.
//
// Link xuất bản mặc định ở dưới; đổi bằng biến môi trường SHEET_PUB_URL nếu dùng file Google Sheet khác.

import { createHash, createHmac, timingSafeEqual } from 'node:crypto'
import { overlay, queue, flush, OpError } from '../lib/pending.mts'
import * as gs from '../lib/gsheets.mts'

const DEFAULT_PUB =
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vQvykG_Qkem_SRPL1M7PA9M_XSPFZpQazmBqyenRTcmq9716V7YiPB0CKI94xF_TZXuFjpeEJYpqx-W/pubhtml'
const SHEETS = { res: 'Người cư trú', own: 'Cơ sở', cfg: 'Tổng quan', acc: 'Tài khoản' }
const FIRST = 5
const RES: Record<string, number> = { fac: 2, name: 3, dob: 6, cccd: 7, addr: 8, job: 9, income: 10, moveIn: 12, status: 14, regDate: 15, expiry: 16, fine: 18, note: 19 }
const OWN: Record<string, number> = { fac: 2, name: 3, dob: 6, cccd: 7, type: 8, phone: 9, rooms: 10, note: 14 }
const ACC = { phone: 2, pass: 3, name: 4, role: 5, state: 6 }
const DATES = new Set(['dob', 'moveIn', 'regDate', 'expiry'])
const ROLES = ['Quản trị', 'Phó Quản trị', 'Nhập liệu', 'Chỉ xem', 'Khách']
// Toàn quyền (đổi số ngày nhắc, khôi phục sao lưu, xóa toàn bộ, xem mục Dữ liệu & bảo mật).
const ADMIN = new Set(['Quản trị', 'Phó Quản trị'])
// Mọi quyền KHÁC Quản trị / Phó Quản trị (Nhập liệu, Khách, Chỉ xem…) bị giới hạn:
// chỉ xem số liệu tổng quan, không bao giờ nhận thông tin cá nhân; Nhập liệu / Khách được nhập hồ sơ mới.
const limited = (role: string) => !ADMIN.has(role)
const SESSION_MS = 12 * 3600_000
const headers = { 'cache-control': 'no-store' }
const ACTIONS = new Set(['login', 'logout', 'read', 'write'])

const j = (data: unknown, status = 200) => Response.json(data, { status, headers })
class HttpError extends Error { code: number; constructor(code: number, msg: string) { super(msg); this.code = code } }

// ---------------- phân quyền ----------------
function guard(role: string, ops: any[]) {
  if (role === 'Chỉ xem') throw new HttpError(403, 'Tài khoản chỉ có quyền xem')
  for (const op of ops) {
    if ((op?.op === 'cfg' || op?.op === 'wipe') && !ADMIN.has(role))
      throw new HttpError(403, 'Chỉ tài khoản Quản trị hoặc Phó Quản trị được thực hiện thao tác này')
    if (limited(role) && !(op?.op === 'add' && (op.kind === 'res' || op.kind === 'own')))
      throw new HttpError(403, 'Tài khoản của bạn chỉ được nhập thêm hồ sơ mới')
  }
}
// Dữ liệu gửi cho Khách: chỉ tên cơ sở (để chọn khi nhập) và tình trạng tạm trú để đếm số liệu tổng quan.
// Không có họ tên, ngày sinh, CCCD, SĐT, ảnh, địa chỉ… và không có số dòng thật trên sheet.
function guestData(d: any) {
  return {
    cfg: { days: Number(d?.cfg?.days) || 20 },
    own: (d?.own || []).map((o: any, i: number) => ({ row: i + 1, fac: String(o.fac || '') })),
    res: (d?.res || []).map((r: any, i: number) => ({ row: i + 1, status: String(r.status || ''), expiry: String(r.expiry || '') })),
    pending: Number(d?.pending) || 0,
    at: d?.at,
  }
}
function scrub(out: any) {
  if (out && out.user && limited(out.user.role)) {
    if (out.data) out.data = guestData(out.data)
    delete out.results
  }
  return out
}

// ---------------- đọc bản xuất bản ----------------
function pubBase() {
  const u = (process.env.SHEET_PUB_URL || DEFAULT_PUB).trim()
  const m = u.match(/^(https:\/\/docs\.google\.com\/spreadsheets\/d\/e\/[\w-]+)\/pub/)
  if (!m) throw new HttpError(500, 'SHEET_PUB_URL không đúng dạng link "Xuất bản lên web" của Google Sheet')
  return m[1]
}

let gidCache: { at: number; base: string; map: Record<string, string> } | null = null
async function gids(base: string) {
  if (gidCache && gidCache.base === base && Date.now() - gidCache.at < 5 * 60_000) return gidCache.map
  const r = await fetch(`${base}/pubhtml?t=${Date.now()}`, { signal: AbortSignal.timeout(8000) })
  if (!r.ok) throw new HttpError(503, 'Google Sheet chưa được "Xuất bản lên web" hoặc link xuất bản không đúng')
  const html = await r.text(), map: Record<string, string> = {}
  for (const m of html.matchAll(/name: "((?:[^"\\]|\\.)*)", pageUrl: "[^"]*", gid: "(\d+)"/g))
    map[JSON.parse(`"${m[1]}"`).trim()] = m[2]
  if (!Object.keys(map).length) throw new HttpError(503, 'Không đọc được danh sách sheet từ bản xuất bản')
  gidCache = { at: Date.now(), base, map }
  return map
}

function parseCSV(t: string): string[][] {
  const rows: string[][] = []
  let row: string[] = [], cell = '', q = false
  for (let i = 0; i < t.length; i++) {
    const c = t[i]
    if (q) { if (c === '"') { if (t[i + 1] === '"') { cell += '"'; i++ } else q = false } else cell += c }
    else if (c === '"') q = true
    else if (c === ',') { row.push(cell); cell = '' }
    else if (c === '\n') { row.push(cell); rows.push(row); row = []; cell = '' }
    else if (c !== '\r') cell += c
  }
  if (cell || row.length) { row.push(cell); rows.push(row) }
  return rows
}

async function sheet(name: string) {
  const base = pubBase(), map = await gids(base), gid = map[name]
  if (!gid) throw new HttpError(503, `Bản xuất bản chưa có sheet "${name}". Hãy xuất bản "Toàn bộ tài liệu".`)
  const r = await fetch(`${base}/pub?gid=${gid}&single=true&output=csv&t=${Date.now()}`, { signal: AbortSignal.timeout(8000) })
  if (!r.ok) throw new HttpError(503, `Không đọc được sheet "${name}" từ bản xuất bản`)
  return parseCSV(await r.text())
}

const iso = (s: string) => {
  s = String(s || '').trim()
  const m = s.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})$/)
  if (m) return `${m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`
  return /^\d{4}-\d\d-\d\d$/.test(s) ? s : ''
}

function rows(data: string[][], map: Record<string, number>) {
  const out: Record<string, unknown>[] = []
  data.slice(FIRST - 1).forEach((d, i) => {
    if (!(d[map.fac - 1] || '').trim() && !(d[map.name - 1] || '').trim()) return
    const o: Record<string, unknown> = { row: FIRST + i, photo: '' }
    for (const k in map) { const v = (d[map[k] - 1] || '').trim(); o[k] = DATES.has(k) ? iso(v) : v }
    out.push(o)
  })
  return out
}

async function readAll() {
  const [res, own, cfg] = await Promise.all([sheet(SHEETS.res), sheet(SHEETS.own), sheet(SHEETS.cfg)])
  return { cfg: { days: Number((cfg[3] || [])[2]) || 20 }, own: rows(own, OWN), res: rows(res, RES), at: new Date().toISOString() }
}

// ---------------- đăng nhập theo sheet Tài khoản ----------------
const normPhone = (s: unknown) => {
  let d = String(s || '').replace(/\D/g, '')
  if (d.startsWith('84') && d.length === 11) d = '0' + d.slice(2)
  if (d.length === 9) d = '0' + d
  return d
}
type Acc = { phone: string; pass: string; name: string; role: string; active: boolean; key: string }
async function accounts(data?: string[][]): Promise<Acc[]> {
  data = data || (gs.enabled() ? (await gs.values([SHEETS.acc], { [SHEETS.acc]: 'G' }))[SHEETS.acc] : await sheet(SHEETS.acc))
  return data.slice(FIRST - 1).flatMap((d) => {
    const phone = normPhone(d[ACC.phone - 1]), pass = d[ACC.pass - 1] || ''
    if (!phone || !pass) return []
    const role = ROLES.includes((d[ACC.role - 1] || '').trim()) ? d[ACC.role - 1].trim() : 'Chỉ xem'
    const active = (d[ACC.state - 1] || '').trim() !== 'Khóa'
    // Khóa ký phiên phụ thuộc mật khẩu/quyền/trạng thái → đổi trên sheet là phiên cũ hết hiệu lực.
    const key = createHash('sha256').update(['xtvh-session', phone, pass, role, active].join('|')).digest('hex')
    return [{ phone, pass, name: (d[ACC.name - 1] || '').trim(), role, active, key }]
  })
}
const b64 = (s: string) => Buffer.from(s).toString('base64url')
const sign = (payload: string, key: string) => createHmac('sha256', key).update(payload).digest('base64url')
const pub = (a: Acc) => ({ phone: a.phone, name: a.name, role: a.role })

const fails = new Map<string, { n: number; until: number }>()
async function login(phone: string, password: string, accs?: string[][]) {
  phone = normPhone(phone)
  if (!phone || !password) throw new HttpError(400, 'Nhập số điện thoại và mật khẩu')
  const f = fails.get(phone)
  if (f && f.n >= 5 && f.until > Date.now()) throw new HttpError(429, 'Sai mật khẩu quá 5 lần. Thử lại sau 15 phút.')
  const a = (await accounts(accs)).find((x) => x.phone === phone)
  if (!a || a.pass !== String(password)) {
    fails.set(phone, { n: (f && f.until > Date.now() ? f.n : 0) + 1, until: Date.now() + 15 * 60_000 })
    await new Promise((r) => setTimeout(r, 700))
    throw new HttpError(401, 'Số điện thoại hoặc mật khẩu không đúng')
  }
  if (!a.active) throw new HttpError(403, 'Tài khoản đang bị khóa')
  fails.delete(phone)
  const payload = b64(JSON.stringify({ p: a.phone, e: Date.now() + SESSION_MS }))
  return { token: `p.${payload}.${sign(payload, a.key)}`, user: pub(a) }
}

async function auth(token: unknown, accs?: string[][]) {
  const [kind, payload, sig] = String(token || '').split('.')
  if (kind !== 'p' || !payload || !sig) throw new HttpError(401, 'Chưa đăng nhập')
  let s: { p: string; e: number }
  try { s = JSON.parse(Buffer.from(payload, 'base64url').toString()) } catch { throw new HttpError(401, 'Phiên không hợp lệ') }
  if (!s.e || s.e < Date.now()) throw new HttpError(401, 'Phiên đăng nhập đã hết hạn, hãy đăng nhập lại')
  const a = (await accounts(accs)).find((x) => x.phone === s.p)
  const ok = a && (() => { const x = Buffer.from(sign(payload, a.key)), y = Buffer.from(sig); return x.length === y.length && timingSafeEqual(x, y) })()
  if (!a || !ok || !a.active) throw new HttpError(401, 'Thông tin đăng nhập đã thay đổi trên Google Sheet, hãy đăng nhập lại')
  return a
}

// Chưa có Apps Script: dữ liệu sheet + các thay đổi nhập trên web đang chờ (Netlify Database).
// Không kết nối được cơ sở dữ liệu → quay về chế độ chỉ xem như trước.
async function view(sheet?: Awaited<ReturnType<typeof readAll>>) {
  sheet = sheet || (await readAll())
  try { return { data: await overlay(sheet), readOnly: false, queue: true } }
  catch (e) { console.error('pending overlay', e); return { data: sheet, readOnly: true } }
}

async function pubMode(body: any) {
  switch (body.action) {
    case 'login': { const r = await login(body.phone, body.password); return { ok: true, ...r, ...(await view()) } }
    case 'logout': return { ok: true }
    case 'read': { const a = await auth(body.token); return { ok: true, user: pub(a), ...(await view()) } }
    case 'write': {
      const a = await auth(body.token)
      const ops = Array.isArray(body.ops) ? body.ops.slice(0, 600) : []
      guard(a.role, ops)
      const sheet = await readAll()
      try { await queue(sheet, ops, a.phone) }
      catch (e) { if (e instanceof OpError) throw new HttpError(e.code, e.message); throw new HttpError(503, 'Không lưu được dữ liệu nhập trên web – thử lại sau') }
      return { ok: true, user: pub(a), ...(await view(sheet)) }
    }
  }
}

// ---------------- Google Sheets API (ghi thẳng, không cần Apps Script) ----------------
const WIDTH = { [SHEETS.res]: 'S', [SHEETS.own]: 'N', [SHEETS.cfg]: 'C', [SHEETS.acc]: 'G' }
const PHOTO = 5
const TEXTS = new Set(['cccd', 'phone'])
const INPUT: Record<string, string[]> = {
  res: ['fac', 'name', 'dob', 'cccd', 'addr', 'job', 'income', 'moveIn', 'status', 'regDate', 'expiry', 'fine', 'note'],
  own: ['fac', 'name', 'dob', 'cccd', 'type', 'phone', 'rooms', 'note'],
}

function apiData(v: Record<string, string[][]>) {
  const cfg = v[SHEETS.cfg] || []
  return { cfg: { days: Number((cfg[3] || [])[2]) || 20 }, own: rows(v[SHEETS.own] || [], OWN), res: rows(v[SHEETS.res] || [], RES), at: new Date().toISOString() }
}
async function apiRead(withAcc = true) {
  const names = [SHEETS.res, SHEETS.own, SHEETS.cfg, ...(withAcc ? [SHEETS.acc] : [])]
  const v = await gs.values(names, WIDTH)
  return { data: apiData(v), acc: v[SHEETS.acc] }
}

// Nội dung 1 ô theo đúng kiểu: ngày → ngày dd/mm/yyyy, CCCD/SĐT → chữ (giữ số 0 đầu), số → số, còn lại → chữ.
// Luôn ghi dạng giá trị (không bao giờ là công thức).
function cell(k: string, raw: unknown) {
  const v = String(raw ?? '').trim()
  if (!v) return { d: {}, f: 'userEnteredValue' }
  if (DATES.has(k)) {
    const m = v.match(/^(\d{4})-(\d\d)-(\d\d)$/)
    if (m) return {
      d: { userEnteredValue: { numberValue: (Date.UTC(+m[1], +m[2] - 1, +m[3]) - Date.UTC(1899, 11, 30)) / 86400000 }, userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'dd/mm/yyyy' } } },
      f: 'userEnteredValue,userEnteredFormat.numberFormat',
    }
  }
  if (TEXTS.has(k)) return { d: { userEnteredValue: { stringValue: v }, userEnteredFormat: { numberFormat: { type: 'TEXT' } } }, f: 'userEnteredValue,userEnteredFormat.numberFormat' }
  if (/^-?\d+(\.\d+)?$/.test(v) && !/^-?0\d/.test(v) && v.length < 15) return { d: { userEnteredValue: { numberValue: Number(v) } }, f: 'userEnteredValue' }
  return { d: { userEnteredValue: { stringValue: v.slice(0, 5000) } }, f: 'userEnteredValue' }
}

let chain: Promise<unknown> = Promise.resolve()
// Áp các thao tác của webapp lên sheet: chỉ ghi cột nhập liệu, không đụng cột công thức. Tương đương Apps Script.
function apiApply(ops: any[], role: string) {
  const run = async () => {
    for (const op of ops)
      if ((op.op === 'cfg' || op.op === 'wipe') && !ADMIN.has(role)) throw new HttpError(403, 'Chỉ tài khoản Quản trị hoặc Phó Quản trị được thực hiện thao tác này')
    const meta = await gs.meta(true)
    for (const n of [SHEETS.res, SHEETS.own]) if (!meta[n]) throw new HttpError(503, `Google Sheet chưa có sheet "${n}"`)
    const v = await gs.values([SHEETS.res, SHEETS.own], WIDTH)
    const S = {
      res: { sh: meta[SHEETS.res], map: RES, grid: v[SHEETS.res] },
      own: { sh: meta[SHEETS.own], map: OWN, grid: v[SHEETS.own] },
    }
    const reqs: unknown[] = []
    const get = (k: 'res' | 'own', row: number, f: string) => String((S[k].grid[row - 1] || [])[S[k].map[f] - 1] ?? '').trim()
    const setGrid = (k: 'res' | 'own', row: number, f: string, val: string) => {
      const g = S[k].grid
      while (g.length < row) g.push([])
      g[row - 1][S[k].map[f] - 1] = val
    }
    const put = (k: 'res' | 'own', row: number, col: number, c: { d: unknown; f: string }) =>
      reqs.push({ updateCells: { range: { sheetId: S[k].sh.id, startRowIndex: row - 1, endRowIndex: row, startColumnIndex: col - 1, endColumnIndex: col }, rows: [{ values: [c.d] }], fields: c.f } })
    const write = (k: 'res' | 'own', row: number, rec: Record<string, unknown>) => {
      for (const f of INPUT[k]) if (f in rec) { put(k, row, S[k].map[f], cell(f, rec[f])); setGrid(k, row, f, String(rec[f] ?? '').trim()) }
    }
    const clear = (k: 'res' | 'own', row: number) => {
      write(k, row, Object.fromEntries(INPUT[k].map((f) => [f, ''])))
      put(k, row, PHOTO, { d: {}, f: 'userEnteredValue' })
    }
    const check = (k: 'res' | 'own', row: number, expect: unknown) => {
      if (!row || row < FIRST || row > S[k].sh.rows) throw new HttpError(409, 'Dòng không hợp lệ')
      if (expect != null && get(k, row, 'name') !== String(expect).trim())
        throw new HttpError(409, 'Dữ liệu trên Google Sheet vừa thay đổi. Đã tải lại, hãy thử lại.')
    }
    const results: unknown[] = []
    for (const op of ops) {
      const k: 'res' | 'own' = op.kind === 'own' ? 'own' : 'res'
      const rec = op.rec && typeof op.rec === 'object' ? op.rec : {}
      switch (op.op) {
        case 'add': {
          if (!String(rec.name ?? '').trim() && !String(rec.fac ?? '').trim()) throw new HttpError(400, 'Thiếu dữ liệu')
          let row = 0
          for (let r = FIRST; r <= S[k].sh.rows; r++) if (!get(k, r, 'fac') && !get(k, r, 'name')) { row = r; break }
          if (!row) { // hết dòng trống: thêm dòng cuối và chép công thức của dòng trên
            const n = S[k].sh.rows, id = S[k].sh.id
            reqs.push({ insertDimension: { range: { sheetId: id, dimension: 'ROWS', startIndex: n, endIndex: n + 1 }, inheritFromBefore: true } })
            reqs.push({ copyPaste: { source: { sheetId: id, startRowIndex: n - 1, endRowIndex: n }, destination: { sheetId: id, startRowIndex: n, endRowIndex: n + 1 }, pasteType: 'PASTE_NORMAL' } })
            S[k].sh.rows = row = n + 1
            put(k, row, PHOTO, { d: {}, f: 'userEnteredValue' })
          }
          write(k, row, { ...Object.fromEntries(INPUT[k].map((f) => [f, ''])), ...rec })
          if (!get(k, row, 'name') && !get(k, row, 'fac')) setGrid(k, row, 'name', ' ')
          results.push({ row })
          break
        }
        case 'update': {
          const row = Number(op.row)
          check(k, row, op.expect)
          const old = get(k, row, 'fac')
          write(k, row, rec)
          const nf = String(rec.fac ?? '').trim()
          if (k === 'own' && nf && old && nf !== old)
            for (let r = FIRST; r <= S.res.grid.length; r++) if (get('res', r, 'fac') === old) write('res', r, { fac: nf })
          results.push({ row })
          break
        }
        case 'delete': {
          const row = Number(op.row)
          check(k, row, op.expect)
          const fac = get(k, row, 'fac')
          if (k === 'own' && fac) for (let r = FIRST; r <= S.res.grid.length; r++) if (get('res', r, 'fac') === fac) clear('res', r)
          clear(k, row)
          results.push({ row })
          break
        }
        case 'cfg': {
          const cfg = meta[SHEETS.cfg]
          if (!cfg) throw new HttpError(503, `Google Sheet chưa có sheet "${SHEETS.cfg}"`)
          reqs.push({ updateCells: { range: { sheetId: cfg.id, startRowIndex: 3, endRowIndex: 4, startColumnIndex: 2, endColumnIndex: 3 }, rows: [{ values: [{ userEnteredValue: { numberValue: Number(op.days) || 20 } }] }], fields: 'userEnteredValue' } })
          results.push({})
          break
        }
        case 'wipe': {
          for (const kk of ['res', 'own'] as const) for (let r = FIRST; r <= S[kk].grid.length; r++) if (get(kk, r, 'fac') || get(kk, r, 'name')) clear(kk, r)
          results.push({})
          break
        }
        default: throw new HttpError(400, 'Thao tác không hỗ trợ')
      }
    }
    await gs.batchUpdate(reqs)
    return results
  }
  const p = chain.then(run, run)
  chain = p.catch(() => {})
  return p
}

// Thay đổi đã nhập trên web khi chưa nối được sheet (hàng đợi Netlify Database) → ghi thẳng lên sheet.
const apiFlush = () =>
  flush(async (ops) => {
    try { await apiApply(ops, 'Nhập liệu'); return { ok: true } }
    catch (e) { return { ok: false, code: e instanceof HttpError || e instanceof gs.SheetError ? e.code : 502 } }
  }).catch((e) => console.error('flush', e))

async function apiMode(body: any) {
  switch (body.action) {
    case 'login': {
      const { acc } = await apiRead(true)
      const r = await login(body.phone, body.password, acc)
      await apiFlush()
      return { ok: true, ...r, data: (await apiRead(false)).data }
    }
    case 'logout': return { ok: true }
    case 'read': {
      if (body.token) await apiFlush()
      const { data, acc } = await apiRead(true)
      const a = await auth(body.token, acc)
      return { ok: true, user: pub(a), data }
    }
    case 'write': {
      const a = await auth(body.token)
      const ops = Array.isArray(body.ops) ? body.ops.slice(0, 600) : []
      guard(a.role, ops)
      await apiFlush()
      const results = await apiApply(ops, a.role)
      return { ok: true, results, user: pub(a), data: (await apiRead(false)).data }
    }
  }
}

// ---------------- Apps Script ----------------
async function scriptMode(url: string, raw: string) {
  const r = await fetch(url, { method: 'POST', body: raw, headers: { 'content-type': 'text/plain;charset=utf-8' }, redirect: 'follow' })
  const text = await r.text()
  try { return JSON.parse(text) } catch {
    throw new HttpError(502, 'Apps Script không trả về dữ liệu hợp lệ – kiểm tra lại triển khai Ứng dụng web (quyền truy cập: Bất kỳ ai)')
  }
}

export default async (req: Request) => {
  if (req.method !== 'POST') return new Response('Method not allowed', { status: 405 })
  const raw = await req.text()
  if (raw.length > 5_000_000) return j({ ok: false, error: 'Dữ liệu quá lớn' }, 413)
  let body: any = {}
  try { body = JSON.parse(raw) } catch {}
  if (!ACTIONS.has(body.action)) return j({ ok: false, error: 'Yêu cầu không hợp lệ' }, 400)
  try {
    if (gs.enabled()) return j(scrub(await apiMode(body)))
    const url = process.env.GOOGLE_SCRIPT_URL
    if (url && body.token && (body.action === 'read' || body.action === 'write'))
      // Đẩy các thay đổi đã nhập trên web lúc chưa có Apps Script lên đúng cột Google Sheet.
      await flush((ops) => scriptMode(url, JSON.stringify({ action: 'write', token: body.token, ops }))).catch((e) => console.error('flush', e))
    const data = scrub(url ? await scriptMode(url, raw) : await pubMode(body))
    const status = data.ok ? 200 : [400, 401, 403, 409, 429, 503].includes(data.code) ? data.code : 502
    return j(data, status)
  } catch (e) {
    const code = e instanceof HttpError || e instanceof gs.SheetError ? e.code : 502
    // Không đọc được bản xuất bản và chưa có Apps Script → webapp hiện chế độ xem thử dữ liệu mẫu.
    const setup = !gs.enabled() && !process.env.GOOGLE_SCRIPT_URL && code === 503 && body.action !== 'write' ? { setup: false } : {}
    return j({ ok: false, code, error: String((e as Error)?.message || e), ...setup }, code)
  }
}

export const config = { path: '/api/sheet' }
