import { pgTable, serial, text, integer, jsonb, timestamp } from 'drizzle-orm/pg-core'

// Thay đổi nhập trên webapp khi CHƯA cài Apps Script (GOOGLE_SCRIPT_URL).
// Hiển thị ngay trên web; khi đặt GOOGLE_SCRIPT_URL sẽ tự ghi lên đúng cột Google Sheet rồi xóa khỏi bảng này.
export const pendingOps = pgTable('pending_ops', {
  id: serial().primaryKey(),
  kind: text().notNull(), // 'res' (Người cư trú) | 'own' (Cơ sở)
  action: text().notNull(), // 'add' | 'update' | 'delete'
  row: integer(), // dòng trên sheet (update/delete); null khi thêm mới
  expect: text(), // họ tên trên dòng đó lúc sửa
  base: jsonb(), // dữ liệu dòng sheet lúc sửa – sheet đổi khác thì sheet được ưu tiên
  rec: jsonb().notNull().default({}),
  createdBy: text('created_by'),
  createdAt: timestamp('created_at').defaultNow(),
})
