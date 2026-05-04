import { Prisma } from "@prisma/client";
import { Router } from "express";
import { z } from "zod";

import { prisma } from "../../lib/prisma.js";
import { requireAuth } from "../../middleware/auth.js";

const calendarRouter = Router();

calendarRouter.use(requireAuth);

const ymd8 = z.string().regex(/^\d{8}$/);
const yyyymm = z.string().regex(/^\d{6}$/);

const postBody = z.object({
  ymd: ymd8,
  reason: z.string().optional(),
});

const patchBody = z.object({
  reason: z.string(),
});

function yyyymmRangeToYmdBounds(yyyyMmFrom: string, yyyyMmTo: string): { min: string; max: string } {
  let a = yyyyMmFrom;
  let b = yyyyMmTo;
  if (a > b) {
    [a, b] = [b, a];
  }
  const y1 = a.slice(0, 4);
  const m1 = a.slice(4, 6);
  const y2 = b.slice(0, 4);
  const m2 = b.slice(4, 6);
  const min = `${y1}${m1}01`;
  const lastD = new Date(Number.parseInt(y2, 10), Number.parseInt(m2, 10), 0).getDate();
  const max = `${y2}${m2}${String(lastD).padStart(2, "0")}`;
  return { min, max };
}

calendarRouter.get("/", async (req, res, next) => {
  try {
    const yearQ = req.query.year;
    const fromQ = req.query.from;
    const toQ = req.query.to;

    let min: string;
    let max: string;

    if (typeof yearQ === "string" && /^\d{4}$/.test(yearQ)) {
      min = `${yearQ}0101`;
      max = `${yearQ}1231`;
    } else if (typeof fromQ === "string" && typeof toQ === "string") {
      const fromParsed = yyyymm.safeParse(fromQ);
      const toParsed = yyyymm.safeParse(toQ);
      if (!fromParsed.success || !toParsed.success) {
        res.status(400).json({ error: "from と to は YYYYMM 形式で指定してください" });
        return;
      }
      const bounds = yyyymmRangeToYmdBounds(fromQ, toQ);
      min = bounds.min;
      max = bounds.max;
    } else {
      res.status(400).json({ error: "year=YYYY または from=YYYYMM&to=YYYYMM を指定してください" });
      return;
    }

    const rows = await prisma.exYasumiX.findMany({
      where: { ymd: { gte: min, lte: max } },
      orderBy: { ymd: "asc" },
    });

    res.json({ data: rows });
  } catch (err) {
    next(err);
  }
});

calendarRouter.post("/", async (req, res, next) => {
  const parsed = postBody.safeParse(req.body);
  if (!parsed.success) {
    res.status(400).json({ error: "Invalid request body" });
    return;
  }

  try {
    const created = await prisma.exYasumiX.create({
      data: {
        ymd: parsed.data.ymd,
        reason: parsed.data.reason ?? null,
      },
    });
    res.status(201).json({ data: created });
  } catch (err) {
    if (err instanceof Prisma.PrismaClientKnownRequestError && err.code === "P2002") {
      const e = new Error("この日付は既に登録されています") as Error & { status: number };
      e.status = 409;
      next(e);
      return;
    }
    next(err);
  }
});

calendarRouter.patch("/:ymd", async (req, res, next) => {
  const ymd = req.params.ymd;
  if (!/^\d{8}$/.test(ymd)) {
    res.status(400).json({ error: "ymd は8桁の日付文字列で指定してください" });
    return;
  }

  const parsed = patchBody.safeParse(req.body);
  if (!parsed.success) {
    res.status(400).json({ error: "Invalid request body" });
    return;
  }

  try {
    const updated = await prisma.exYasumiX.update({
      where: { ymd },
      data: { reason: parsed.data.reason },
    });
    res.json({ data: updated });
  } catch (err) {
    if (err instanceof Prisma.PrismaClientKnownRequestError && err.code === "P2025") {
      res.status(404).json({ error: "該当する休日が見つかりません" });
      return;
    }
    next(err);
  }
});

calendarRouter.delete("/:ymd", async (req, res, next) => {
  const ymd = req.params.ymd;
  if (!/^\d{8}$/.test(ymd)) {
    res.status(400).json({ error: "ymd は8桁の日付文字列で指定してください" });
    return;
  }

  try {
    await prisma.exYasumiX.delete({ where: { ymd } });
    res.json({ data: { ok: true as const, ymd } });
  } catch (err) {
    if (err instanceof Prisma.PrismaClientKnownRequestError && err.code === "P2025") {
      res.status(404).json({ error: "該当する休日が見つかりません" });
      return;
    }
    next(err);
  }
});

export { calendarRouter };
