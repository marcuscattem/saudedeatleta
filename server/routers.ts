import { COOKIE_NAME, ONE_YEAR_MS } from "@shared/const";
import { getSessionCookieOptions } from "./_core/cookies";
import { systemRouter } from "./_core/systemRouter";
import { publicProcedure, router, protectedProcedure } from "./_core/trpc";
import { z } from "zod";
import { createAntropometria, createFpmEvaluation, createIsakEvaluation, getUserAntropometrias, getUserFpmEvaluations, getUserIsakEvaluations, getDb, getUserByOpenId, upsertUser } from "./db";
import { generateAntropometriaExcel, generateFpmExcel, generateIsakExcel } from "./excel-generator";
import { eq } from "drizzle-orm";
import { antropometrias, fpmEvaluations, isakEvaluations, type Antropometria, type FpmEvaluation, type IsakEvaluation } from "../drizzle/schema";
import { storagePut } from "./storage";
import { sdk } from "./_core/sdk";
import { signInWithSupabase, signUpWithSupabase } from "./supabase-auth";
import type { TrpcContext } from "./_core/context";

const EXCEL_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

function dateStamp(date = new Date()) {
  return date.toISOString().split("T")[0];
}

function safeFilenamePart(value: string) {
  return value.trim().replace(/[^a-zA-Z0-9_-]+/g, "_").replace(/^_+|_+$/g, "") || "participante";
}

function collectionFilename(date: Date, participantId: string) {
  return `${dateStamp(date)}_${safeFilenamePart(participantId)}.xlsx`;
}

function createExcelResponse(buffer: Buffer, filename: string, excelUrl?: string | null) {
  return {
    data: buffer.toString("base64"),
    filename,
    excelUrl: excelUrl ?? null,
  };
}

async function uploadEvaluationExcel(filename: string, buffer: Buffer, userId: number) {
  const storageResult = await storagePut(
    `evaluations/${userId}/${filename}`,
    buffer,
    EXCEL_CONTENT_TYPE,
  );

  return storageResult.url;
}

async function createSupabaseSession(
  ctx: Pick<TrpcContext, "req" | "res">,
  user: { openId: string; name: string; email: string | null },
) {
  await upsertUser({
    openId: user.openId,
    name: user.name,
    email: user.email,
    loginMethod: "supabase",
    lastSignedIn: new Date(),
  });

  const savedUser = await getUserByOpenId(user.openId);
  if (!savedUser) {
    throw new Error("Banco de dados não configurado para salvar o usuário");
  }

  const sessionToken = await sdk.createSessionToken(user.openId, {
    name: user.name,
    expiresInMs: ONE_YEAR_MS,
  });

  ctx.res.cookie(COOKIE_NAME, sessionToken, {
    ...getSessionCookieOptions(ctx.req),
    maxAge: ONE_YEAR_MS,
  });

  return savedUser;
}

export const appRouter = router({
  system: systemRouter,
  auth: router({
    me: publicProcedure.query(opts => opts.ctx.user),
    signUp: publicProcedure
      .input(z.object({
        name: z.string().min(1),
        email: z.string().email(),
        password: z.string().min(6),
      }))
      .mutation(async ({ ctx, input }) => {
        const user = await signUpWithSupabase(input);
        return createSupabaseSession(ctx, user);
      }),
    signIn: publicProcedure
      .input(z.object({
        email: z.string().email(),
        password: z.string().min(6),
      }))
      .mutation(async ({ ctx, input }) => {
        const user = await signInWithSupabase(input);
        return createSupabaseSession(ctx, user);
      }),
    logout: publicProcedure.mutation(({ ctx }) => {
      const cookieOptions = getSessionCookieOptions(ctx.req);
      ctx.res.clearCookie(COOKIE_NAME, { ...cookieOptions, maxAge: -1 });
      return {
        success: true,
      } as const;
    }),
  }),

  evaluations: router({
    saveAntropometria: protectedProcedure
      .input(z.object({
        participantId: z.string(),
        date: z.date(),
        bracoMeasurements: z.array(z.number()),
        cinturaMeasurements: z.array(z.number()),
        panturrilhaMeasurements: z.array(z.number()),
      }))
      .mutation(async ({ ctx, input }) => {
        if (!ctx.user) throw new Error("Not authenticated");
        const filename = `antropometria_${safeFilenamePart(input.participantId)}_${dateStamp()}.xlsx`;
        const excelData: Antropometria = {
          id: 0,
          userId: ctx.user.id,
          participantId: input.participantId,
          date: input.date,
          bracoMeasurements: JSON.stringify(input.bracoMeasurements),
          cinturaMeasurements: JSON.stringify(input.cinturaMeasurements),
          panturrilhaMeasurements: JSON.stringify(input.panturrilhaMeasurements),
          excelUrl: null,
          createdAt: new Date(),
          updatedAt: new Date(),
        };
        const excelBuffer = await generateAntropometriaExcel(excelData);
        const excelUrl = await uploadEvaluationExcel(filename, excelBuffer, ctx.user.id);
        const result = await createAntropometria({
          userId: ctx.user.id,
          participantId: input.participantId,
          date: input.date,
          bracoMeasurements: JSON.stringify(input.bracoMeasurements),
          cinturaMeasurements: JSON.stringify(input.cinturaMeasurements),
          panturrilhaMeasurements: JSON.stringify(input.panturrilhaMeasurements),
          excelUrl,
        });
        return { success: true, id: result.insertId, ...createExcelResponse(excelBuffer, filename, excelUrl) };
      }),
    
    saveFpm: protectedProcedure
      .input(z.object({
        participantId: z.string(),
        date: z.date(),
        dominantHand: z.string(),
        bestLeg: z.string(),
        rightMeasurements: z.array(z.number()),
        leftMeasurements: z.array(z.number()),
      }))
      .mutation(async ({ ctx, input }) => {
        if (!ctx.user) throw new Error("Not authenticated");
        const filename = collectionFilename(input.date, input.participantId);
        const excelData: FpmEvaluation = {
          id: 0,
          userId: ctx.user.id,
          participantId: input.participantId,
          date: input.date,
          dominantHand: input.dominantHand,
          bestLeg: input.bestLeg,
          rightMeasurements: JSON.stringify(input.rightMeasurements),
          leftMeasurements: JSON.stringify(input.leftMeasurements),
          excelUrl: null,
          createdAt: new Date(),
          updatedAt: new Date(),
        };
        const excelBuffer = await generateFpmExcel(excelData);
        const excelUrl = await uploadEvaluationExcel(filename, excelBuffer, ctx.user.id);
        const result = await createFpmEvaluation({
          userId: ctx.user.id,
          participantId: input.participantId,
          date: input.date,
          dominantHand: input.dominantHand,
          bestLeg: input.bestLeg,
          rightMeasurements: JSON.stringify(input.rightMeasurements),
          leftMeasurements: JSON.stringify(input.leftMeasurements),
          excelUrl,
        });
        return { success: true, id: result.insertId, ...createExcelResponse(excelBuffer, filename, excelUrl) };
      }),
    
    saveIsak: protectedProcedure
      .input(z.object({
        participantId: z.string(),
        date: z.date(),
        measurements: z.record(z.string(), z.array(z.number())),
      }))
      .mutation(async ({ ctx, input }) => {
        if (!ctx.user) throw new Error("Not authenticated");
        const filename = `isak_${safeFilenamePart(input.participantId)}_${dateStamp()}.xlsx`;
        const excelData: IsakEvaluation = {
          id: 0,
          userId: ctx.user.id,
          participantId: input.participantId,
          date: input.date,
          measurements: JSON.stringify(input.measurements),
          excelUrl: null,
          createdAt: new Date(),
          updatedAt: new Date(),
        };
        const excelBuffer = await generateIsakExcel(excelData);
        const excelUrl = await uploadEvaluationExcel(filename, excelBuffer, ctx.user.id);
        const result = await createIsakEvaluation({
          userId: ctx.user.id,
          participantId: input.participantId,
          date: input.date,
          measurements: JSON.stringify(input.measurements),
          excelUrl,
        });
        return { success: true, id: result.insertId, ...createExcelResponse(excelBuffer, filename, excelUrl) };
      }),
    
    getAntropometrias: protectedProcedure
      .query(async ({ ctx }) => {
        if (!ctx.user) throw new Error("Not authenticated");
        return await getUserAntropometrias(ctx.user.id);
      }),
    
    getFpmEvaluations: protectedProcedure
      .query(async ({ ctx }) => {
        if (!ctx.user) throw new Error("Not authenticated");
        return await getUserFpmEvaluations(ctx.user.id);
      }),
    
    getIsakEvaluations: protectedProcedure
      .query(async ({ ctx }) => {
        if (!ctx.user) throw new Error("Not authenticated");
        return await getUserIsakEvaluations(ctx.user.id);
      }),

    downloadAntropometriaExcel: protectedProcedure
      .input(z.object({ id: z.number() }))
      .mutation(async ({ ctx, input }) => {
        if (!ctx.user) throw new Error("Not authenticated");
        const db = await getDb();
        if (!db) throw new Error("Database not available");

        const data = await db
          .select()
          .from(antropometrias)
          .where(eq(antropometrias.id, input.id))
          .limit(1);

        if (!data[0] || data[0].userId !== ctx.user.id) {
          throw new Error("Not authorized");
        }

        const buffer = await generateAntropometriaExcel(data[0]);
        return createExcelResponse(
          buffer,
          `antropometria_${safeFilenamePart(data[0].participantId)}_${dateStamp()}.xlsx`,
          data[0].excelUrl,
        );
      }),

    downloadFpmExcel: protectedProcedure
      .input(z.object({ id: z.number() }))
      .mutation(async ({ ctx, input }) => {
        if (!ctx.user) throw new Error("Not authenticated");
        const db = await getDb();
        if (!db) throw new Error("Database not available");

        const data = await db
          .select()
          .from(fpmEvaluations)
          .where(eq(fpmEvaluations.id, input.id))
          .limit(1);

        if (!data[0] || data[0].userId !== ctx.user.id) {
          throw new Error("Not authorized");
        }

        const buffer = await generateFpmExcel(data[0]);
        return createExcelResponse(
          buffer,
          collectionFilename(data[0].date, data[0].participantId),
          data[0].excelUrl,
        );
      }),

    downloadIsakExcel: protectedProcedure
      .input(z.object({ id: z.number() }))
      .mutation(async ({ ctx, input }) => {
        if (!ctx.user) throw new Error("Not authenticated");
        const db = await getDb();
        if (!db) throw new Error("Database not available");

        const data = await db
          .select()
          .from(isakEvaluations)
          .where(eq(isakEvaluations.id, input.id))
          .limit(1);

        if (!data[0] || data[0].userId !== ctx.user.id) {
          throw new Error("Not authorized");
        }

        const buffer = await generateIsakExcel(data[0]);
        return createExcelResponse(
          buffer,
          `isak_${safeFilenamePart(data[0].participantId)}_${dateStamp()}.xlsx`,
          data[0].excelUrl,
        );
      }),
  }),
});

export type AppRouter = typeof appRouter;
