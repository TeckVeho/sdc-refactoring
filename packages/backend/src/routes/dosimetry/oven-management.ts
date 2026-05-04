import { Router } from "express";
import type { Prisma } from "@prisma/client";
import { z } from "zod";

import { prisma } from "../../lib/prisma.js";
import { requireAuth } from "../../middleware/auth.js";
import { tempCorrection } from "../../services/temp-correction.js";

const ovenManagementRouter = Router();

const postBodySchema = z.object({
  ovenkey: z.string().min(1).max(64),
  json: z.record(z.unknown()),
});

const patchBodySchema = z.object({
  json: z.record(z.unknown()),
});

function isObjectJson(v: Prisma.JsonValue | null): v is Record<string, Prisma.JsonValue> {
  return v !== null && typeof v === "object" && !Array.isArray(v);
}

function mergeJson(
  base: Prisma.JsonValue | null,
  patch: Record<string, unknown>,
): Prisma.InputJsonValue {
  const b = isObjectJson(base) ? { ...base } : {};
  return { ...b, ...patch } as Prisma.InputJsonValue;
}

/** 管理点線量シート AG 列相当: (-3.2478+23.386*H/G)*(-0.006*K+1.0558) */
ovenManagementRouter.get("/correction-preview", requireAuth, (req, res) => {
  const sokutei = Number(req.query.sokutei);
  const atusa = Number(req.query.atusa);
  const ondo = Number(req.query.ondo);
  if (!Number.isFinite(sokutei) || !Number.isFinite(atusa) || !Number.isFinite(ondo)) {
    res.status(400).json({ error: "sokutei, atusa, ondo は数値で指定してください" });
    return;
  }
  if (atusa === 0) {
    res.status(400).json({ error: "厚さ(atusa)が0です" });
    return;
  }
  const ratio = sokutei / atusa;
  const correctedKanri = tempCorrection(ratio, ondo);
  res.json({
    data: {
      /** 測定値(ABS) / 素子厚さ(mm) — temp-correction 第1引数 */
      absPerThickness: ratio,
      /** 補正後管理点（仕様書 AG 列） */
      correctedKanriPoint: correctedKanri,
    },
  });
});

ovenManagementRouter.get("/", requireAuth, async (_req, res) => {
  const [exOvenTb, syouj3ov] = await Promise.all([
    prisma.exOvenTb.findMany({
      orderBy: { ovenkey: "asc" },
    }),
    prisma.syouj3ov.findMany({
      take: 500,
      orderBy: { ovno: "asc" },
    }),
  ]);

  res.json({
    data: {
      exOvenTb: exOvenTb.map((r) => ({
        ovenkey: r.ovenkey,
        json: r.json,
      })),
      syouj3ov,
    },
  });
});

ovenManagementRouter.post("/", requireAuth, async (req, res) => {
  const parsed = postBodySchema.safeParse(req.body);
  if (!parsed.success) {
    res.status(400).json({ error: "Invalid body", details: parsed.error.flatten() });
    return;
  }
  const { ovenkey, json } = parsed.data;

  const row = await prisma.exOvenTb.upsert({
    where: { ovenkey },
    create: {
      ovenkey,
      json: json as Prisma.InputJsonValue,
    },
    update: {
      json: json as Prisma.InputJsonValue,
    },
  });

  res.status(201).json({
    data: { ovenkey: row.ovenkey, json: row.json },
  });
});

ovenManagementRouter.patch("/:ovenkey", requireAuth, async (req, res) => {
  const ovenkey = req.params.ovenkey;
  if (!ovenkey) {
    res.status(400).json({ error: "ovenkey が必要です" });
    return;
  }

  const parsed = patchBodySchema.safeParse(req.body);
  if (!parsed.success) {
    res.status(400).json({ error: "Invalid body", details: parsed.error.flatten() });
    return;
  }

  const existing = await prisma.exOvenTb.findUnique({ where: { ovenkey } });
  if (!existing) {
    res.status(404).json({ error: "Not found" });
    return;
  }

  const merged = mergeJson(existing.json, parsed.data.json);

  const row = await prisma.exOvenTb.update({
    where: { ovenkey },
    data: { json: merged },
  });

  res.json({
    data: { ovenkey: row.ovenkey, json: row.json },
  });
});

export { ovenManagementRouter };
