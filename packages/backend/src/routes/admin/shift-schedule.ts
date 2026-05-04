import { Router } from "express";
import { z } from "zod";

import { prisma } from "../../lib/prisma.js";
import { requireAuth } from "../../middleware/auth.js";
import { shiftCycleFromYearMonth } from "../../services/shift-cycle.js";

const shiftScheduleRouter = Router();

/** 仕様書 §2.1 KinmuMoto 抜粋（Web 既定チップ用） */
export const DEFAULT_SHIFT_SYMBOLS: { symbol: string; memo: string }[] = [
  { symbol: "点検", memo: "装置点検" },
  { symbol: "積", memo: "" },
  { symbol: "S希", memo: "S（早出）希望" },
  { symbol: "S", memo: "早出勤務" },
  { symbol: "休希", memo: "休み希望" },
  { symbol: "代休", memo: "代替休日" },
  { symbol: "早出", memo: "早出勤務" },
  { symbol: "有休", memo: "有給休暇" },
  { symbol: "夜L", memo: "夜勤L" },
  { symbol: "S残1", memo: "S残業1号機" },
  { symbol: "S残2", memo: "S残業2号機" },
  { symbol: "1", memo: "1号機勤務" },
  { symbol: "2", memo: "2号機勤務" },
  { symbol: "3", memo: "3号機勤務" },
  { symbol: "残1", memo: "残業1号機" },
  { symbol: "残2", memo: "残業2号機" },
  { symbol: "3残1", memo: "3号機残業1" },
  { symbol: "3残2", memo: "3号機残業2" },
  { symbol: "①", memo: "特殊勤務①" },
  { symbol: "②", memo: "特殊勤務②" },
  { symbol: "夜2", memo: "夜勤2号機" },
  { symbol: "夜1", memo: "夜勤1号機" },
  { symbol: "明", memo: "明け番" },
  { symbol: "休", memo: "公休" },
];

const querySchema = z.object({
  year: z.coerce.number().int().min(1980).max(2100),
  month: z.coerce.number().int().min(1).max(12),
});

const postBodySchema = z.object({
  year: z.number().int().min(1980).max(2100),
  month: z.number().int().min(1).max(12),
  label: z.string().max(200).optional(),
  /** shano -> dateIso -> 勤務記号 */
  grid: z.record(z.string(), z.record(z.string(), z.string())).optional(),
});

function cycleStartDate(year: number, month: number): Date {
  const { startIso } = shiftCycleFromYearMonth(year, month);
  const [y, m, d] = startIso.split("-").map(Number);
  return new Date(y!, m! - 1, d!);
}

/** GET /api/admin/shift-schedule?year=&month= */
shiftScheduleRouter.get("/", requireAuth, async (req, res) => {
  const parsed = querySchema.safeParse(req.query);
  if (!parsed.success) {
    res.status(400).json({ error: "year, month が必要です（1〜12）" });
    return;
  }

  const { year, month } = parsed.data;
  const cycle = shiftCycleFromYearMonth(year, month);
  const cStart = cycleStartDate(year, month);

  try {
    const [roster, dbSymbols, dbCirculation, saved, allShain] = await Promise.all([
      prisma.shiftEmployee.findMany({
        orderBy: { shano: "asc" },
        select: { id: true, shano: true, note: true },
      }),
      prisma.workTime.findMany({
        orderBy: { symbol: "asc" },
        select: { symbol: true, hours: true, memo: true },
      }),
      prisma.circulationList.findMany({
        orderBy: { sortOrder: "asc" },
        select: { id: true, name: true, sortOrder: true },
      }),
      prisma.shiftSchedule.findFirst({
        where: { cycleStart: cStart },
        select: { id: true, cycleStart: true, label: true, raw: true },
      }),
      prisma.shainmst.findMany({
        take: 48,
        orderBy: { shano: "asc" },
        select: { shano: true, shaname: true },
      }),
    ]);

    const rosterShanos = new Set(roster.map((r) => r.shano));
    const employees =
      roster.length > 0
        ? (
            await prisma.shainmst.findMany({
              where: { shano: { in: [...rosterShanos] } },
              orderBy: { shano: "asc" },
              select: { shano: true, shaname: true },
            })
          ).map((s) => ({
            shano: s.shano,
            shaname: s.shaname,
            inRoster: true as const,
          }))
        : allShain.map((s) => ({
            shano: s.shano,
            shaname: s.shaname,
            inRoster: false as const,
          }));

    const symbols =
      dbSymbols.length > 0
        ? dbSymbols.map((w) => ({
            symbol: w.symbol,
            memo: w.memo ?? "",
            hours: w.hours != null ? String(w.hours) : null,
          }))
        : DEFAULT_SHIFT_SYMBOLS.map((d) => ({
            symbol: d.symbol,
            memo: d.memo,
            hours: null as string | null,
          }));

    const grid =
      saved?.raw &&
      typeof saved.raw === "object" &&
      saved.raw !== null &&
      "grid" in saved.raw &&
      typeof (saved.raw as { grid?: unknown }).grid === "object"
        ? ((saved.raw as { grid: Record<string, Record<string, string>> }).grid ?? {})
        : {};

    res.json({
      data: {
        cycle,
        scheduleId: saved?.id ?? null,
        label: saved?.label ?? null,
        symbols,
        circulationList: dbCirculation,
        employees,
        grid,
        rosterSize: roster.length,
      },
    });
  } catch (e) {
    res.status(500).json({ error: e instanceof Error ? e.message : "Server error" });
  }
});

/** POST /api/admin/shift-schedule — グリッド保存（Upsert） */
shiftScheduleRouter.post("/", requireAuth, async (req, res) => {
  const parsed = postBodySchema.safeParse(req.body);
  if (!parsed.success) {
    res.status(400).json({ error: "Invalid body" });
    return;
  }

  const { year, month, label, grid } = parsed.data;
  const cStart = cycleStartDate(year, month);

  try {
    const existing = await prisma.shiftSchedule.findFirst({
      where: { cycleStart: cStart },
      select: { id: true },
    });

    const raw = { grid: grid ?? {}, savedAt: new Date().toISOString() };

    if (existing) {
      await prisma.shiftSchedule.update({
        where: { id: existing.id },
        data: { label: label ?? undefined, raw },
      });
      res.json({ ok: true, id: existing.id, updated: true });
      return;
    }

    const created = await prisma.shiftSchedule.create({
      data: {
        cycleStart: cStart,
        label: label ?? null,
        raw,
      },
      select: { id: true },
    });
    res.json({ ok: true, id: created.id, updated: false });
  } catch (e) {
    res.status(500).json({ error: e instanceof Error ? e.message : "Server error" });
  }
});

export { shiftScheduleRouter };
