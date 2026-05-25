import { eq } from "drizzle-orm";
import { drizzle } from "drizzle-orm/mysql2";
import { InsertUser, users, antropometrias, InsertAntropometria, fpmEvaluations, InsertFpmEvaluation, isakEvaluations, InsertIsakEvaluation } from "../drizzle/schema";
import { ENV } from './_core/env';

let _db: ReturnType<typeof drizzle> | null = null;

function getInsertId(result: unknown) {
  const insertResult = Array.isArray(result) ? result[0] : result;

  if (insertResult && typeof insertResult === "object" && "insertId" in insertResult) {
    const insertId = (insertResult as { insertId?: unknown }).insertId;
    return typeof insertId === "number" ? insertId : 0;
  }

  return 0;
}

// Lazily create the drizzle instance so local tooling can run without a DB.
export async function getDb() {
  if (!_db && process.env.DATABASE_URL) {
    try {
      _db = drizzle(process.env.DATABASE_URL);
    } catch (error) {
      console.warn("[Database] Failed to connect:", error);
      _db = null;
    }
  }
  return _db;
}

export async function upsertUser(user: InsertUser): Promise<void> {
  if (!user.openId) {
    throw new Error("User openId is required for upsert");
  }

  const db = await getDb();
  if (!db) {
    console.warn("[Database] Cannot upsert user: database not available");
    return;
  }

  try {
    const values: InsertUser = {
      openId: user.openId,
    };
    const updateSet: Record<string, unknown> = {};

    const textFields = ["name", "email", "loginMethod"] as const;
    type TextField = (typeof textFields)[number];

    const assignNullable = (field: TextField) => {
      const value = user[field];
      if (value === undefined) return;
      const normalized = value ?? null;
      values[field] = normalized;
      updateSet[field] = normalized;
    };

    textFields.forEach(assignNullable);

    if (user.lastSignedIn !== undefined) {
      values.lastSignedIn = user.lastSignedIn;
      updateSet.lastSignedIn = user.lastSignedIn;
    }
    if (user.role !== undefined) {
      values.role = user.role;
      updateSet.role = user.role;
    } else if (user.openId === ENV.ownerOpenId) {
      values.role = 'admin';
      updateSet.role = 'admin';
    }

    if (!values.lastSignedIn) {
      values.lastSignedIn = new Date();
    }

    if (Object.keys(updateSet).length === 0) {
      updateSet.lastSignedIn = new Date();
    }

    await db.insert(users).values(values).onDuplicateKeyUpdate({
      set: updateSet,
    });
  } catch (error) {
    console.error("[Database] Failed to upsert user:", error);
    throw error;
  }
}

export async function getUserByOpenId(openId: string) {
  const db = await getDb();
  if (!db) {
    console.warn("[Database] Cannot get user: database not available");
    return undefined;
  }

  const result = await db.select().from(users).where(eq(users.openId, openId)).limit(1);

  return result.length > 0 ? result[0] : undefined;
}

export async function createAntropometria(data: InsertAntropometria) {
  const db = await getDb();
  if (!db) throw new Error("Database not available");
  
  const result = await db.insert(antropometrias).values({
    ...data,
    bracoMeasurements: typeof data.bracoMeasurements === 'string' ? data.bracoMeasurements : JSON.stringify(data.bracoMeasurements),
    cinturaMeasurements: typeof data.cinturaMeasurements === 'string' ? data.cinturaMeasurements : JSON.stringify(data.cinturaMeasurements),
    panturrilhaMeasurements: typeof data.panturrilhaMeasurements === 'string' ? data.panturrilhaMeasurements : JSON.stringify(data.panturrilhaMeasurements),
  });
  
  return { insertId: getInsertId(result) };
}

export async function createFpmEvaluation(data: InsertFpmEvaluation) {
  const db = await getDb();
  if (!db) throw new Error("Database not available");
  
  const result = await db.insert(fpmEvaluations).values({
    ...data,
    rightMeasurements: typeof data.rightMeasurements === 'string' ? data.rightMeasurements : JSON.stringify(data.rightMeasurements),
    leftMeasurements: typeof data.leftMeasurements === 'string' ? data.leftMeasurements : JSON.stringify(data.leftMeasurements),
  });
  
  return { insertId: getInsertId(result) };
}

export async function createIsakEvaluation(data: InsertIsakEvaluation) {
  const db = await getDb();
  if (!db) throw new Error("Database not available");
  
  const result = await db.insert(isakEvaluations).values({
    ...data,
    measurements: typeof data.measurements === 'string' ? data.measurements : JSON.stringify(data.measurements),
  });
  
  return { insertId: getInsertId(result) };
}

export async function getUserAntropometrias(userId: number) {
  const db = await getDb();
  if (!db) return [];
  
  const results = await db.select().from(antropometrias).where(eq(antropometrias.userId, userId));
  return results.map(r => ({
    ...r,
    bracoMeasurements: JSON.parse(r.bracoMeasurements),
    cinturaMeasurements: JSON.parse(r.cinturaMeasurements),
    panturrilhaMeasurements: JSON.parse(r.panturrilhaMeasurements),
  }));
}

export async function getUserFpmEvaluations(userId: number) {
  const db = await getDb();
  if (!db) return [];
  
  const results = await db.select().from(fpmEvaluations).where(eq(fpmEvaluations.userId, userId));
  return results.map(r => ({
    ...r,
    rightMeasurements: JSON.parse(r.rightMeasurements),
    leftMeasurements: JSON.parse(r.leftMeasurements),
  }));
}

export async function getUserIsakEvaluations(userId: number) {
  const db = await getDb();
  if (!db) return [];
  
  const results = await db.select().from(isakEvaluations).where(eq(isakEvaluations.userId, userId));
  return results.map(r => ({
    ...r,
    measurements: JSON.parse(r.measurements),
  }));
}
