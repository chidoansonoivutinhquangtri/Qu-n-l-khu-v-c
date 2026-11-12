// Ghi/đọc THẲNG Google Sheet qua Google Sheets API bằng tài khoản dịch vụ (service account) – không cần Apps Script.
//
// Cấu hình (biến môi trường Netlify):
//   GOOGLE_SERVICE_ACCOUNT_JSON  – toàn bộ nội dung tệp khóa JSON của service account
//                                  (hoặc dùng GOOGLE_SERVICE_ACCOUNT_EMAIL + GOOGLE_PRIVATE_KEY)
//   GOOGLE_SHEET_ID              – (tùy chọn) ID file Google Sheet, mặc định là file của xóm trọ bên dưới
// File Google Sheet phải được Chia sẻ quyền "Người chỉnh sửa" cho email của service account.
import { createSign } from 'node:crypto'

const DEFAULT_ID = '1poy8BnWtqzxMPxetQRpgHchmu6P-XM4HA8HT-lt_JWE'
const API = 'https://sheets.googleapis.com/v4/spreadsheets'

export class SheetError extends Error { code: number; constructor(code: number, m: string) { super(m); this.code = code } }

type Cred = { email: string; key: string }
function cred(): Cred | null {
  const raw = (process.env.GOOGLE_SERVICE_ACCOUNT_JSON || '').trim()
  if (raw) {
    try {
      const j = JSON.parse(raw)
      if (j.client_email && j.private_key) return { email: j.client_email, key: String(j.private_key).replace(/\\n/g, '\n') }
    } catch {}
    throw new SheetError(500, 'GOOGLE_SERVICE_ACCOUNT_JSON không phải tệp khóa JSON hợp lệ của service account')
  }
  const email = (process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL || '').trim(), key = (process.env.GOOGLE_PRIVATE_KEY || '').trim()
  return email && key ? { email, key: key.replace(/\\n/g, '\n') } : null
}
export const enabled = () => {
  try { return !!cred() } catch { return true } // cấu hình sai vẫn vào chế độ API để báo lỗi rõ ràng
}
export const serviceEmail = () => { try { return cred()?.email || '' } catch { return '' } }
const sheetId = () => {
  const v = (process.env.GOOGLE_SHEET_ID || DEFAULT_ID).trim()
  return v.match(/\/d\/([\w-]{20,})/)?.[1] || v
}

let tok: { v: string; exp: number; email: string } | null = null
async function token() {
  const c = cred()!
  if (tok && tok.email === c.email && tok.exp > Date.now() + 60_000) return tok.v
  const now = Math.floor(Date.now() / 1000)
  const b = (o: unknown) => Buffer.from(JSON.stringify(o)).toString('base64url')
  const unsigned = `${b({ alg: 'RS256', typ: 'JWT' })}.${b({ iss: c.email, scope: 'https://www.googleapis.com/auth/spreadsheets', aud: 'https://oauth2.googleapis.com/token', iat: now, exp: now + 3600 })}`
  let sig: string
  try { sig = createSign('RSA-SHA256').update(unsigned).sign(c.key, 'base64url') }
  catch { throw new SheetError(500, 'Khóa riêng (private_key) của service account không hợp lệ') }
  const r = await fetch('https://oauth2.googleapis.com/token', {
    method: 'POST',
    headers: { 'content-type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({ grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer', assertion: `${unsigned}.${sig}` }),
    signal: AbortSignal.timeout(8000),
  })
  const d: any = await r.json().catch(() => ({}))
  if (!r.ok || !d.access_token) throw new SheetError(503, 'Google từ chối đăng nhập service account – kiểm tra lại tệp khóa JSON (khóa có thể đã bị xóa)')
  tok = { v: d.access_token, exp: Date.now() + (Number(d.expires_in) || 3600) * 1000, email: c.email }
  return tok.v
}

async function call(path: string, init: RequestInit = {}) {
  const r = await fetch(`${API}/${sheetId()}${path}`, {
    ...init,
    headers: { authorization: `Bearer ${await token()}`, 'content-type': 'application/json' },
    signal: AbortSignal.timeout(15000),
  })
  if (r.ok) return r.json()
  const msg = ((await r.json().catch(() => ({}))) as any)?.error?.message || ''
  console.error('sheets api', r.status, msg)
  if (r.status === 403 || r.status === 404)
    throw new SheetError(503, `Service account chưa có quyền với Google Sheet. Mở Google Sheet → Chia sẻ → thêm ${serviceEmail()} với quyền "Người chỉnh sửa".`)
  if (r.status === 429) throw new SheetError(503, 'Google Sheet đang bị giới hạn số lượt truy cập – thử lại sau ít giây')
  if (r.status === 400) throw new SheetError(409, `Google Sheet từ chối thao tác: ${msg}`)
  throw new SheetError(503, 'Không kết nối được Google Sheets API – thử lại sau')
}

const q = (name: string) => `'${name.replace(/'/g, "''")}'`

// Đọc nhiều vùng 1 lượt (giá trị hiển thị, giống bản xuất bản). Trả về mảng 2 chiều bắt đầu từ ô A1 mỗi sheet.
export async function values(names: string[], widths: Record<string, string> = {}) {
  const p = new URLSearchParams({ valueRenderOption: 'FORMATTED_VALUE', majorDimension: 'ROWS' })
  for (const n of names) p.append('ranges', `${q(n)}!A1:${widths[n] || 'Z'}`)
  const d = await call(`/values:batchGet?${p}`)
  const out: Record<string, string[][]> = {}
  names.forEach((n, i) => { out[n] = (d.valueRanges?.[i]?.values || []).map((r: unknown[]) => r.map((x) => String(x ?? ''))) })
  return out
}

let metaCache: { at: number; id: string; map: Record<string, { id: number; rows: number }> } | null = null
export async function meta(fresh = false) {
  if (!fresh && metaCache && metaCache.id === sheetId() && Date.now() - metaCache.at < 60_000) return metaCache.map
  const d = await call('?fields=sheets.properties(sheetId,title,gridProperties.rowCount)')
  const map: Record<string, { id: number; rows: number }> = {}
  for (const s of d.sheets || []) map[String(s.properties.title).trim()] = { id: s.properties.sheetId, rows: s.properties.gridProperties?.rowCount || 0 }
  metaCache = { at: Date.now(), id: sheetId(), map }
  return map
}

export const batchUpdate = (requests: unknown[]) =>
  requests.length ? call(':batchUpdate', { method: 'POST', body: JSON.stringify({ requests }) }) : Promise.resolve({})
