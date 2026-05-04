import { Router } from "express";
import { z } from "zod";

import { prisma } from "../../lib/prisma.js";
import { requireAuth } from "../../middleware/auth.js";
import { computeDoseRate, getCoefficientsPayload } from "../../services/dose-rate-calc.js";

const doseRateCalcRouter = Router();

const postBodySchema = z.object({
  mode: z.enum(["RATE", "INTEG"]),
  potentiometerId: z.union([z.literal(1), z.literal(2)]),
  ionChamberId: z.number().int().min(1).max(9),
  readValue: z.number().finite(),
  temperatureC: z.number().finite(),
  pressureHpa: z.number().finite().positive(),
  tpEnabled: z.boolean(),
  uncertaintyOrg: z.enum(["ANTM", "JQA", "none"]),
  cableOn: z.boolean(),
  targetDoseGy: z.number().nonnegative().optional(),
  decay: z
    .object({
      halfLifeDays: z.number().positive(),
      elapsedDays: z.number().nonnegative(),
    })
    .optional()
    .nullable(),
});

/**
 * GET /api/dosimetry/dose-rate-calc — 換算定数・機器テーブル（校正係数）と ex_kanri 参照
 */
doseRateCalcRouter.get("/", requireAuth, async (_req, res, next) => {
  try {
    let kanriSources: { sikibetu: string; kousinn: string | null; ricvm: string | null; hppp: string | null }[] = [];
    try {
      kanriSources = await prisma.exKanriTb.findMany({
        orderBy: { sikibetu: "asc" },
        select: { sikibetu: true, kousinn: true, ricvm: true, hppp: true },
      });
    } catch {
      // DB 未接続・テーブル未作成時も係数は返す
    }
    res.json({ ...getCoefficientsPayload(), kanriSources });
  } catch (e) {
    next(e);
  }
});

/**
 * POST /api/dosimetry/dose-rate-calc — 減衰・表示値付きの計算結果
 */
doseRateCalcRouter.post("/", requireAuth, (req, res, next) => {
  try {
    const body = postBodySchema.parse(req.body);
    const result = computeDoseRate({
      mode: body.mode,
      potentiometerId: body.potentiometerId,
      ionChamberId: body.ionChamberId,
      readValue: body.readValue,
      temperatureC: body.temperatureC,
      pressureHpa: body.pressureHpa,
      tpEnabled: body.tpEnabled,
      uncertaintyOrg: body.uncertaintyOrg,
      cableOn: body.cableOn,
      targetDoseGy: body.targetDoseGy,
      decay: body.decay ?? null,
    });
    res.json({ ok: true, result });
  } catch (e) {
    if (e instanceof z.ZodError) {
      res.status(400).json({ ok: false, error: e.flatten() });
      return;
    }
    if (e instanceof RangeError) {
      res.status(400).json({ ok: false, error: e.message });
      return;
    }
    next(e);
  }
});

export { doseRateCalcRouter };
