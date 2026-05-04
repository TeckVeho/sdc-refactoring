import { randomUUID } from "node:crypto";

import { Prisma } from "@prisma/client";
import { Router } from "express";
import { z } from "zod";

import { prisma } from "../../lib/prisma.js";
import { requireAuth } from "../../middleware/auth.js";

const planRouter = Router();

const putBodySchema = z.object({
  json: z.unknown(),
});

const postBodySchema = z.object({
  planId: z.string().min(1).optional(),
  json: z.unknown().optional(),
});

function todayYmd(): string {
  return new Date().toISOString().slice(0, 10);
}

function defaultGanttJson(): Prisma.InputJsonValue {
  return {
    data: [
      {
        id: 1,
        text: "新規タスク",
        start_date: todayYmd(),
        duration: 1,
        open: true,
      },
    ],
    links: [],
  };
}

/** GET /api/irradiation/machine1/plan — 計画ヘッダ一覧 + 明細 */
planRouter.get("/", requireAuth, async (_req, res, next) => {
  try {
    const headers = await prisma.exR1Keikaku1.findMany({
      orderBy: { planId: "asc" },
    });
    const planIds = headers.map((h) => h.planId);
    const details =
      planIds.length === 0
        ? []
        : await prisma.exR1Keikaku2.findMany({
            where: { planId: { in: planIds } },
            orderBy: { detailId: "asc" },
          });

    const byPlan = new Map<string, typeof details>();
    for (const d of details) {
      if (d.planId == null) continue;
      const list = byPlan.get(d.planId) ?? [];
      list.push(d);
      byPlan.set(d.planId, list);
    }

    const plans = headers.map((h) => ({
      planId: h.planId,
      json: h.json,
      details: (byPlan.get(h.planId) ?? []).map((d) => ({
        detailId: d.detailId,
        planId: d.planId,
        json: d.json,
      })),
    }));

    res.json({ data: { plans } });
  } catch (e) {
    next(e);
  }
});

/** PUT /api/irradiation/machine1/plan/:planId — ヘッダ json を丸ごと置換（MVP） */
planRouter.put("/:planId", requireAuth, async (req, res, next) => {
  const planId = req.params.planId;
  if (!planId || planId.trim() === "") {
    res.status(400).json({ error: "planId が不正です" });
    return;
  }

  const parsed = putBodySchema.safeParse(req.body);
  if (!parsed.success) {
    res.status(400).json({ error: "リクエスト本文が不正です" });
    return;
  }

  try {
    const json = parsed.data.json as Prisma.InputJsonValue;
    await prisma.exR1Keikaku1.upsert({
      where: { planId },
      create: { planId, json },
      update: { json },
    });
    res.json({ data: { planId, json } });
  } catch (e) {
    next(e);
  }
});

/** POST /api/irradiation/machine1/plan — 新規計画（planId 省略時は自動採番） */
planRouter.post("/", requireAuth, async (req, res, next) => {
  const parsed = postBodySchema.safeParse(req.body);
  if (!parsed.success) {
    res.status(400).json({ error: "リクエスト本文が不正です" });
    return;
  }

  const planId = parsed.data.planId?.trim() || randomUUID();
  const json = (parsed.data.json ?? defaultGanttJson()) as Prisma.InputJsonValue;

  try {
    await prisma.exR1Keikaku1.upsert({
      where: { planId },
      create: { planId, json },
      update: { json },
    });
    res.status(201).json({ data: { planId, json } });
  } catch (e) {
    next(e);
  }
});

export { planRouter };
