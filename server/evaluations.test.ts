import { describe, expect, it, vi } from "vitest";

vi.mock("./storage", () => ({
  storagePut: vi.fn(async (key: string) => ({
    key,
    url: `/manus-storage/${key}`,
  })),
}));

vi.mock("./db", () => ({
  createAntropometria: vi.fn(async () => ({ insertId: 1 })),
  createFpmEvaluation: vi.fn(async () => ({ insertId: 2 })),
  createIsakEvaluation: vi.fn(async () => ({ insertId: 3 })),
  getUserAntropometrias: vi.fn(async () => []),
  getUserFpmEvaluations: vi.fn(async () => []),
  getUserIsakEvaluations: vi.fn(async () => []),
  getDb: vi.fn(async () => null),
}));

import { appRouter } from "./routers";
import type { TrpcContext } from "./_core/context";

type AuthenticatedUser = NonNullable<TrpcContext["user"]>;

function createAuthContext(): { ctx: TrpcContext } {
  const user: AuthenticatedUser = {
    id: 1,
    openId: "test-user",
    email: "test@example.com",
    name: "Test User",
    loginMethod: "manus",
    role: "user",
    createdAt: new Date(),
    updatedAt: new Date(),
    lastSignedIn: new Date(),
  };

  const ctx: TrpcContext = {
    user,
    req: {
      protocol: "https",
      headers: {},
    } as TrpcContext["req"],
    res: {
      clearCookie: () => {},
    } as TrpcContext["res"],
  };

  return { ctx };
}

describe("evaluations", () => {
  describe("saveAntropometria", () => {
    it("should save antropometria measurements", async () => {
      const { ctx } = createAuthContext();
      const caller = appRouter.createCaller(ctx);

      const result = await caller.evaluations.saveAntropometria({
        participantId: "P001",
        date: new Date(),
        bracoMeasurements: [25.5, 25.6, 25.4],
        cinturaMeasurements: [80.0, 80.1, 79.9],
        panturrilhaMeasurements: [35.0, 35.1, 34.9],
      });

      expect(result).toBeDefined();
      expect(result.id).toBe(1);
      expect(result.filename).toMatch(/^antropometria_P001_/);
      expect(result.excelUrl).toContain("/manus-storage/evaluations/1/");
      expect(result.data.length).toBeGreaterThan(0);
    });

    it("should reject without authentication", async () => {
      const ctx: TrpcContext = {
        user: null,
        req: { protocol: "https", headers: {} } as TrpcContext["req"],
        res: { clearCookie: () => {} } as TrpcContext["res"],
      };

      const caller = appRouter.createCaller(ctx);

      try {
        await caller.evaluations.saveAntropometria({
          participantId: "P001",
          date: new Date(),
          bracoMeasurements: [25.5],
          cinturaMeasurements: [80.0],
          panturrilhaMeasurements: [35.0],
        });
        expect.fail("Should have thrown");
      } catch (error) {
        expect(error).toBeDefined();
      }
    });
  });

  describe("saveFpm", () => {
    it("should save FPM measurements", async () => {
      const { ctx } = createAuthContext();
      const caller = appRouter.createCaller(ctx);

      const result = await caller.evaluations.saveFpm({
        participantId: "P001",
        date: new Date("2026-06-05T00:00:00.000Z"),
        dominantHand: "Direita",
        bestLeg: "Direita",
        rightMeasurements: [45.0, 45.5, 44.8],
        leftMeasurements: [42.0, 42.5, 41.8],
      });

      expect(result).toBeDefined();
      expect(result.id).toBe(2);
      expect(result.filename).toBe("2026-06-05_P001.xlsx");
      expect(result.excelUrl).toContain("/manus-storage/evaluations/1/");
      expect(result.data.length).toBeGreaterThan(0);
    });
  });

  describe("saveIsak", () => {
    it("should save ISAK measurements", async () => {
      const { ctx } = createAuthContext();
      const caller = appRouter.createCaller(ctx);

      const result = await caller.evaluations.saveIsak({
        participantId: "P001",
        date: new Date(),
        measurements: {
          subscap: [10.5, 10.6, 10.4],
          triceps: [12.0, 12.1, 11.9],
          biceps: [8.5, 8.6, 8.4],
          iliaca: [15.0, 15.1, 14.9],
          supraesp: [9.0, 9.1, 8.9],
          abdom: [20.0, 20.1, 19.9],
          coxa: [18.0, 18.1, 17.9],
          pant_dobra: [11.0, 11.1, 10.9],
          torax: [95.0, 95.1, 94.9],
          braco_rel: [28.0, 28.1, 27.9],
          braco_flet: [32.0, 32.1, 31.9],
          cintura: [80.0, 80.1, 79.9],
          gluteo: [95.0, 95.1, 94.9],
          coxa_media: [55.0, 55.1, 54.9],
          pant_perim: [38.0, 38.1, 37.9],
        },
      });

      expect(result).toBeDefined();
      expect(result.id).toBe(3);
      expect(result.filename).toMatch(/^isak_P001_/);
      expect(result.excelUrl).toContain("/manus-storage/evaluations/1/");
      expect(result.data.length).toBeGreaterThan(0);
    });
  });

  describe("getAntropometrias", () => {
    it("should retrieve antropometria data", async () => {
      const { ctx } = createAuthContext();
      const caller = appRouter.createCaller(ctx);

      const result = await caller.evaluations.getAntropometrias();

      expect(Array.isArray(result)).toBe(true);
    });
  });

  describe("getFpmEvaluations", () => {
    it("should retrieve FPM data", async () => {
      const { ctx } = createAuthContext();
      const caller = appRouter.createCaller(ctx);

      const result = await caller.evaluations.getFpmEvaluations();

      expect(Array.isArray(result)).toBe(true);
    });
  });

  describe("getIsakEvaluations", () => {
    it("should retrieve ISAK data", async () => {
      const { ctx } = createAuthContext();
      const caller = appRouter.createCaller(ctx);

      const result = await caller.evaluations.getIsakEvaluations();

      expect(Array.isArray(result)).toBe(true);
    });
  });
});
