import { integer, jsonb, pgTable, serial, text, timestamp } from 'drizzle-orm/pg-core'

// Thay đổi nhập trên web đang chờ ghi lên Google Sheet.
export const pendingOps = pgTable('pending_ops', {
  id: serial('id').primaryKey(),
  kind: text('kind').notNull(),
  action: text('action').notNull(),
  row: integer('row'),
  expect: text('expect'),
  base: jsonb('base'),
  rec: jsonb('rec').notNull().default({}),
  createdBy: text('created_by'),
  createdAt: timestamp('created_at').defaultNow(),
})
