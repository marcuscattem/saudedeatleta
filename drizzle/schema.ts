import { int, mysqlEnum, mysqlTable, text, timestamp, varchar } from "drizzle-orm/mysql-core";

/**
 * Core user table backing auth flow.
 * Extend this file with additional tables as your product grows.
 * Columns use camelCase to match both database fields and generated types.
 */
export const users = mysqlTable("users", {
  /**
   * Surrogate primary key. Auto-incremented numeric value managed by the database.
   * Use this for relations between tables.
   */
  id: int("id").autoincrement().primaryKey(),
  /** Manus OAuth identifier (openId) returned from the OAuth callback. Unique per user. */
  openId: varchar("openId", { length: 64 }).notNull().unique(),
  name: text("name"),
  email: varchar("email", { length: 320 }),
  loginMethod: varchar("loginMethod", { length: 64 }),
  role: mysqlEnum("role", ["user", "admin"]).default("user").notNull(),
  createdAt: timestamp("createdAt").defaultNow().notNull(),
  updatedAt: timestamp("updatedAt").defaultNow().onUpdateNow().notNull(),
  lastSignedIn: timestamp("lastSignedIn").defaultNow().notNull(),
});

export type User = typeof users.$inferSelect;
export type InsertUser = typeof users.$inferInsert;

// Tabela para Antropometria
export const antropometrias = mysqlTable("antropometrias", {
  id: int("id").autoincrement().primaryKey(),
  userId: int("userId").notNull(),
  participantId: varchar("participantId", { length: 255 }).notNull(),
  date: timestamp("date").notNull(),
  bracoMeasurements: text("bracoMeasurements").notNull(), // JSON array
  cinturaMeasurements: text("cinturaMeasurements").notNull(), // JSON array
  panturrilhaMeasurements: text("panturrilhaMeasurements").notNull(), // JSON array
  excelUrl: text("excelUrl"), // URL do arquivo Excel armazenado
  createdAt: timestamp("createdAt").defaultNow().notNull(),
  updatedAt: timestamp("updatedAt").defaultNow().onUpdateNow().notNull(),
});

export type Antropometria = typeof antropometrias.$inferSelect;
export type InsertAntropometria = typeof antropometrias.$inferInsert;

// Tabela para Força de Preensão Manual (FPM)
export const fpmEvaluations = mysqlTable("fpmEvaluations", {
  id: int("id").autoincrement().primaryKey(),
  userId: int("userId").notNull(),
  participantId: varchar("participantId", { length: 255 }).notNull(),
  date: timestamp("date").notNull(),
  dominantHand: varchar("dominantHand", { length: 20 }).notNull(),
  bestLeg: varchar("bestLeg", { length: 20 }).notNull(),
  rightMeasurements: text("rightMeasurements").notNull(), // JSON array
  leftMeasurements: text("leftMeasurements").notNull(), // JSON array
  excelUrl: text("excelUrl"), // URL do arquivo Excel armazenado
  createdAt: timestamp("createdAt").defaultNow().notNull(),
  updatedAt: timestamp("updatedAt").defaultNow().onUpdateNow().notNull(),
});

export type FpmEvaluation = typeof fpmEvaluations.$inferSelect;
export type InsertFpmEvaluation = typeof fpmEvaluations.$inferInsert;

// Tabela para ISAK 1
export const isakEvaluations = mysqlTable("isakEvaluations", {
  id: int("id").autoincrement().primaryKey(),
  userId: int("userId").notNull(),
  participantId: varchar("participantId", { length: 255 }).notNull(),
  date: timestamp("date").notNull(),
  measurements: text("measurements").notNull(), // JSON object com todas as medidas
  excelUrl: text("excelUrl"), // URL do arquivo Excel armazenado
  createdAt: timestamp("createdAt").defaultNow().notNull(),
  updatedAt: timestamp("updatedAt").defaultNow().onUpdateNow().notNull(),
});

export type IsakEvaluation = typeof isakEvaluations.$inferSelect;
export type InsertIsakEvaluation = typeof isakEvaluations.$inferInsert;