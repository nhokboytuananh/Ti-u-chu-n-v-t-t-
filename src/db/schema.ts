import { pgTable, text, jsonb, timestamp } from "drizzle-orm/pg-core";

export const materials = pgTable("materials", {
  id: text("id").primaryKey(),
  code: text("code").notNull(),
  name: text("name").notNull(),
  content: text("content").notNull(),
  notes: text("notes"),
  excelData: jsonb("excel_data").notNull(),
  images: jsonb("images").notNull(),
  docRequirements: jsonb("doc_requirements"),
  createdAt: timestamp("created_at").defaultNow().notNull(),
  updatedAt: timestamp("updated_at").defaultNow().notNull(),
  updatedBy: text("updated_by"),
});

export const materialAuditLogs = pgTable("material_audit_logs", {
  id: text("id").primaryKey(),
  materialId: text("material_id").notNull(),
  action: text("action").notNull(), // 'create', 'update'
  previousData: jsonb("previous_data"),
  newData: jsonb("new_data"),
  updatedBy: text("updated_by"),
  updatedAt: timestamp("updated_at").defaultNow().notNull(),
});

export const packages = pgTable("packages", {
  id: text("id").primaryKey(),
  name: text("name").notNull(),
  materialIds: jsonb("material_ids").notNull(),
  hiddenTags: jsonb("hidden_tags"),
  hiddenTables: jsonb("hidden_tables"),
  createdAt: timestamp("created_at").defaultNow().notNull(),
});
