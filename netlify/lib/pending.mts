// Hàng đợi thay đổi nhập trên webapp khi CHƯA cài Apps Script (GOOGLE_SCRIPT_URL chưa đặt).
// - Thêm / sửa / xóa trên web được lưu vào Netlify Database và hiện ngay trên web (gắn nhãn "chờ lên sheet").
// - Dữ liệu sheet luôn được ưu tiên: dòng trên sheet bị sửa khác đi, hoặc sheet đã có đúng bản ghi đó,
//   thì thay đổi đang chờ tự bỏ.
// - Khi đặt GOOGLE_SCRIPT_URL: các thay đổi đang chờ tự ghi lên đúng cột Google Sheet rồi xóa khỏi hàng đợi.
import { asc, eq } from 'drizzle-orm'
import { db } from '../../db/index.js'
import { pendingOps } from '../../db/schema.js'

type Rec = Record<string, any>
type Data = { cfg: any; own: Rec[]; res: Rec[]; at: string }
type Op = { kind?: string; op?: string; row?: number; expect?: string; rec?: Rec; days?: number }
type P = typeof pendingOps.$inferSelect

export const FIELDS: Record<string, string[]> = {
  res: ['fac', 'name', 'dob', 'cccd', 'addr', 'job', 'income', 'moveIn', 'status', 'regDate', 'expiry', 'fine', 'note'],
  own: ['fac', 'name', 'dob', 'cccd', 'type', 'phone', 'rooms', 'note'],
}
const ALLOWED = (k: string) => [...FIELDS[k], 'photo']
const s = (v: unknown) => String(v ?? '').trim()
const lower = (v: unknown) => s(v).toLowerCase()

class OpError extends Error { code: number; constructor(code: number, m: string) { super(m); this.code = code } }

function clean(kind: string, rec: Rec = {}) {
  const o: Rec = {}
  for (const k of ALLOWED(kind)) if (k in rec) o[k] = k === 'photo' ? String(rec[k] || '').slice(0, 400_000) : s(rec[k]).slice(0, 500)
  return o
}
const snap = (kind: string, r: Rec) => Object.fromEntries(FIELDS[kind].map((k) => [k, s(r[k])]))
const same = (kind: string, a: Rec, b: Rec) => FIELDS[kind].every((k) => s(a[k]) === s(b[k]))
const applied = (kind: string, row: Rec, rec: Rec) => Object.keys(rec).every((k) => (k === 'photo' && !rec[k]) || s(row[k]) === s(rec[k]))

function onSheet(kind: string, list: Rec[], rec: Rec) {
  if (kind === 'own') return list.some((o) => lower(o.fac) && lower(o.fac) === lower(rec.fac))
  return list.some((r) => lower(r.name) === lower(rec.name) && (s(r.cccd) && s(rec.cccd) ? s(r.cccd) === s(rec.cccd) : lower(r.fac) === lower(rec.fac)))
}

const list = () => db.select().from(pendingOps).orderBy(asc(pendingOps.id))
const drop = (id: number) => db.delete(pendingOps).where(eq(pendingOps.id, id))

// Gộp hàng đợi vào dữ liệu đọc từ sheet. Trả về dữ liệu đã gộp + số thay đổi còn chờ.
export async function overlay(sheet: Data): Promise<Data & { pending: number }> {
  const data: Data = { ...sheet, own: sheet.own.map((x) => ({ ...x })), res: sheet.res.map((x) => ({ ...x })) }
  const raw: Record<string, Map<number, Rec>> = {
    own: new Map(sheet.own.map((x) => [x.row, x])), res: new Map(sheet.res.map((x) => [x.row, x])),
  }
  const stale: number[] = []
  let n = 0
  for (const p of await list()) {
    const kind = p.kind === 'own' ? 'own' : 'res', L = data[kind], rec = (p.rec || {}) as Rec
    if (p.action === 'add') {
      if (onSheet(kind, sheet[kind], rec)) { stale.push(p.id); continue }
      L.push({ ...Object.fromEntries(ALLOWED(kind).map((k) => [k, ''])), ...rec, row: -p.id, pending: 'add' })
      n++
      continue
    }
    const base = raw[kind].get(p.row || 0)
    if (!base || !same(kind, base, (p.base || {}) as Rec)) { stale.push(p.id); continue } // sheet đã đổi → sheet thắng
    const i = L.findIndex((x) => x.row === p.row)
    if (i < 0) { stale.push(p.id); continue }
    if (p.action === 'delete') { L.splice(i, 1); n++; continue }
    if (applied(kind, base, rec)) { stale.push(p.id); continue }
    Object.assign(L[i], rec, { pending: 'update' })
    n++
  }
  if (stale.length) await Promise.all(stale.map(drop))
  return { ...data, pending: n }
}

// Ghi thao tác từ web vào hàng đợi. sheet = dữ liệu đọc từ sheet (chưa gộp).
export async function queue(sheet: Data, ops: Op[], who: string) {
  for (const op of ops) {
    if (op.op === 'cfg' || op.op === 'wipe')
      throw new OpError(503, 'Đổi số ngày nhắc, khôi phục sao lưu và xóa toàn bộ cần kết nối tài khoản dịch vụ Google (GOOGLE_SERVICE_ACCOUNT_JSON) để ghi thẳng lên Google Sheet.')
  }
  for (const op of ops) {
    const kind = op.kind === 'own' ? 'own' : op.kind === 'res' ? 'res' : ''
    if (!kind || !['add', 'update', 'delete'].includes(String(op.op))) throw new OpError(400, 'Thao tác không hỗ trợ')
    const rec = clean(kind, op.rec)
    if (op.op === 'add') {
      if (!s(rec.name) && !s(rec.fac)) throw new OpError(400, 'Thiếu dữ liệu')
      await db.insert(pendingOps).values({ kind, action: 'add', rec, createdBy: who })
      continue
    }
    const row = Number(op.row) || 0, cur = (await overlay(sheet))[kind].find((x) => x.row === row)
    if (!cur || (op.expect != null && s(cur.name) !== s(op.expect)))
      throw new OpError(409, 'Dữ liệu trên Google Sheet vừa thay đổi. Đã tải lại, hãy thử lại.')

    if (row < 0) { // bản ghi thêm trên web, chưa lên sheet
      const [p] = await db.select().from(pendingOps).where(eq(pendingOps.id, -row))
      if (!p) throw new OpError(409, 'Bản ghi vừa được cập nhật lên Google Sheet. Đã tải lại, hãy thử lại.')
      if (op.op === 'delete') await drop(p.id)
      else await db.update(pendingOps).set({ rec: { ...(p.rec as Rec), ...rec } }).where(eq(pendingOps.id, p.id))
    } else {
      const base = sheet[kind].find((x) => x.row === row)!
      const all = await list(), prev = all.find((p) => p.kind === kind && p.row === row && p.action !== 'add')
      if (op.op === 'delete') {
        if (prev) await drop(prev.id)
        await db.insert(pendingOps).values({ kind, action: 'delete', row, expect: s(base.name), base: snap(kind, base), createdBy: who })
      } else if (prev && prev.action === 'update') {
        await db.update(pendingOps).set({ rec: { ...(prev.rec as Rec), ...rec } }).where(eq(pendingOps.id, prev.id))
      } else {
        await db.insert(pendingOps).values({ kind, action: 'update', row, expect: s(base.name), base: snap(kind, base), rec, createdBy: who })
      }
    }

    // Cơ sở: đổi tên → đổi theo ở người cư trú; xóa → xóa luôn người cư trú của cơ sở (giống Apps Script).
    if (kind === 'own') {
      const oldFac = s(cur.fac), newFac = op.op === 'update' && 'fac' in rec ? s(rec.fac) : oldFac
      if (op.op === 'delete' || newFac !== oldFac) {
        const v2 = await overlay(sheet)
        for (const r of v2.res.filter((x) => s(x.fac) === oldFac))
          await queue(sheet, [op.op === 'delete' ? { kind: 'res', op: 'delete', row: r.row, expect: r.name } : { kind: 'res', op: 'update', row: r.row, expect: r.name, rec: { fac: newFac } }], who)
      }
    }
  }
}

// Có Apps Script: đẩy hàng đợi lên sheet (mỗi lần tối đa 15 thao tác), dùng phiên của người đang đăng nhập.
export async function flush(send: (ops: Op[]) => Promise<any>) {
  let items: P[]
  try { items = (await list()).slice(0, 15) } catch { return }
  for (const p of items) {
    const [claimed] = await db.delete(pendingOps).where(eq(pendingOps.id, p.id)).returning()
    if (!claimed) continue // yêu cầu khác đã đẩy
    const op: Op = p.action === 'add'
      ? { kind: p.kind, op: 'add', rec: p.rec as Rec }
      : { kind: p.kind, op: p.action, row: p.row!, expect: p.expect ?? undefined, ...(p.action === 'update' ? { rec: p.rec as Rec } : {}) }
    let r: any
    try { r = await send([op]) } catch { r = { ok: false, code: 502 } }
    if (r && r.ok) continue
    if (r && r.code === 409) continue // dòng trên sheet đã đổi → sheet thắng, bỏ thay đổi
    const { id, ...back } = claimed // lỗi khác (mất kết nối, không có quyền) → trả lại hàng đợi, thử lần sau
    await db.insert(pendingOps).values(back)
    return
  }
}

export { OpError }
