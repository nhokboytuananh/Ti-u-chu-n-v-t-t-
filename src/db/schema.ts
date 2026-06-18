import { pgTable, text, jsonb, timestamp } from "drizzle-orm/pg-core";

export const materials = pgTable("materials", {
  id: text("id").primaryKey(),
  code: text("code").notNull(),
  name: text("name").notNull(),
  content: text("content").notNull(),
  excelData: jsonb("excel_data").notNull(),
  images: jsonb("images").notNull(),
  createdAt: timestamp("created_at").defaultNow().notNull(),
});

export const packages = pgTable("packages", {
  id: text("id").primaryKey(),
  name: text("name").notNull(),
  materialIds: jsonb("material_ids").notNull(),
  hiddenTags: jsonb("hidden_tags"),
  createdAt: timestamp("created_at").defaultNow().notNull(),
});
