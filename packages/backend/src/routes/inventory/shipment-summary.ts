import { Router } from "express";
import { z } from "zod";

import { prisma } from "../../lib/prisma.js";
import { requireAuth } from "../../middleware/auth.js";

const shipmentSummaryRouter = Router();

/** 入荷日・出荷日の表記揺れに対し、比較用に先頭8桁の数字列へ正規化 */
function normalizeYyyymmdd(s: string | null | undefined): string {
  if (s == null) return "";
  const digits = String(s).replace(/\D/g, "");
  return digits.length >= 8 ? digits.slice(0, 8) : digits;
}

function parseYmdToUtcDate(ymd: string): Date {
  const y = Number(ymd.slice(0, 4));
  const m = Number(ymd.slice(4, 6)) - 1;
  const d = Number(ymd.slice(6, 8));
  return new Date(Date.UTC(y, m, d));
}

/** 両端含む日数（同日なら1） */
function inclusiveDayCount(fromYmd: string, toYmd: string): number {
  const a = parseYmdToUtcDate(fromYmd).getTime();
  const b = parseYmdToUtcDate(toYmd).getTime();
  if (b < a) return 0;
  return Math.floor((b - a) / 86400000) + 1;
}

function machineLabel(m: "1" | "2" | "3"): string {
  return `${m}号機`;
}

function toMachineId(s: string | null | undefined): "1" | "2" | "3" | null {
  if (s == null) return null;
  const t = s.trim();
  if (t === "1" || t === "01") return "1";
  if (t === "2" || t === "02") return "2";
  if (t === "3" || t === "03") return "3";
  if (/^[123]$/.test(t)) return t as "1" | "2" | "3";
  return null;
}

function firstQueryString(v: unknown): string | undefined {
  if (v == null || v === "") return undefined;
  if (Array.isArray(v)) {
    return typeof v[0] === "string" ? v[0] : undefined;
  }
  return typeof v === "string" ? v : undefined;
}

const querySchema = z
  .object({
    view: z.enum(["arrival", "shipment", "uptime"]),
    date: z.string().regex(/^\d{8}$/).optional(),
    from: z.string().regex(/^\d{8}$/).optional(),
    to: z.string().regex(/^\d{8}$/).optional(),
  })
  .superRefine((q, ctx) => {
    if (q.from && q.to && q.from > q.to) {
      ctx.addIssue({ code: z.ZodIssueCode.custom, message: "from must be <= to" });
    }
  });

shipmentSummaryRouter.get("/", requireAuth, async (req, res) => {
  const parsed = querySchema.safeParse({
    view: firstQueryString(req.query.view),
    date: firstQueryString(req.query.date),
    from: firstQueryString(req.query.from),
    to: firstQueryString(req.query.to),
  });
  if (!parsed.success) {
    res.status(400).json({ error: "Invalid query" });
    return;
  }
  const { view, date, from, to } = parsed.data;

  const singleDate = date ?? (from && to && from === to ? from : undefined);
  const range =
    from && to
      ? { from, to }
      : date
        ? { from: date, to: date }
        : null;

  try {
    if (view === "arrival" || view === "shipment") {
      if (!singleDate) {
        res.status(400).json({ error: "date (YYYYMMDD) is required" });
        return;
      }
    } else {
      if (!range) {
        res.status(400).json({ error: "date or from+to (YYYYMMDD) is required" });
        return;
      }
    }

    if (view === "arrival" && singleDate) {
      const y = singleDate.slice(0, 4);
      const rows = await prisma.zaiko.findMany({
        where: {
          nyukabi: { contains: y },
        },
        select: { syouso: true, kainame: true, nyukasu: true, nyukabi: true },
      });
      const onDate = rows.filter((r) => normalizeYyyymmdd(r.nyukabi) === singleDate);

      type Machine = "1" | "2" | "3";
      const byMachine: Record<Machine, Map<string, number>> = {
        "1": new Map(),
        "2": new Map(),
        "3": new Map(),
      };
      for (const r of onDate) {
        const m = toMachineId(r.syouso);
        if (!m) continue;
        const name = (r.kainame && r.kainame.trim()) || "(未設定)";
        const add = r.nyukasu ?? 0;
        byMachine[m].set(name, (byMachine[m].get(name) ?? 0) + add);
      }

      const machines: Array<{
        machine: "1" | "2" | "3";
        label: string;
        rows: Array<{ kainame: string; nyukasu: number }>;
        totalNyukasu: number;
      }> = (["1", "2", "3"] as const).map((m) => {
        const map = byMachine[m];
        const list = Array.from(map.entries())
          .map(([kainame, nyukasu]) => ({ kainame, nyukasu }))
          .sort((a, b) => a.kainame.localeCompare(b.kainame, "ja"));
        const totalNyukasu = list.reduce((s, r) => s + r.nyukasu, 0);
        return { machine: m, label: machineLabel(m), rows: list, totalNyukasu };
      });

      res.json({
        data: {
          view: "arrival" as const,
          asOf: new Date().toISOString(),
          date: singleDate,
          series: { machines },
        },
      });
      return;
    }

    if (view === "shipment" && singleDate) {
      const y = singleDate.slice(0, 4);
      const syRows = await prisma.syukar.findMany({
        where: {
          syudate: { contains: y },
        },
        select: { kaisyacd: true, syudate: true, syukasu: true },
      });
      const onDate = syRows.filter((r) => normalizeYyyymmdd(r.syudate) === singleDate);

      const byCompany = new Map<string, { kaisyacd: string; syukasu: number }>();
      for (const r of onDate) {
        const k = r.kaisyacd;
        const prev = byCompany.get(k);
        const add = r.syukasu ?? 0;
        if (prev) {
          prev.syukasu += add;
        } else {
          byCompany.set(k, { kaisyacd: k, syukasu: add });
        }
      }

      const codes = Array.from(byCompany.keys());
      const toku = await prisma.tokumst.findMany({
        where: { kaisyacd: { in: codes } },
        select: { kaisyacd: true, coname: true, kairname: true },
      });
      const nameByCd = new Map(
        toku.map((t) => [t.kaisyacd, t.coname.trim() || t.kairname?.trim() || t.kaisyacd]),
      );

      const rows = Array.from(byCompany.values())
        .map((row) => ({
          kaisyacd: row.kaisyacd,
          companyName: nameByCd.get(row.kaisyacd) ?? row.kaisyacd,
          syukasu: row.syukasu,
        }))
        .sort((a, b) => b.syukasu - a.syukasu);

      const totalSyukasu = rows.reduce((s, r) => s + r.syukasu, 0);

      res.json({
        data: {
          view: "shipment" as const,
          asOf: new Date().toISOString(),
          date: singleDate,
          series: { rows, totalSyukasu },
        },
      });
      return;
    }

    if (view === "uptime" && range) {
      const { from: fromY, to: toY } = range;
      const yStart = Math.min(parseInt(fromY.slice(0, 4), 10), parseInt(toY.slice(0, 4), 10));
      const yEnd = Math.max(parseInt(fromY.slice(0, 4), 10), parseInt(toY.slice(0, 4), 10));
      const yearPrefixes: string[] = [];
      for (let y = yStart; y <= yEnd; y++) {
        yearPrefixes.push(String(y));
      }

      const inRangeYmd = (n: string) => n.length === 8 && n >= fromY && n <= toY;

      // --- Sengnr1（sdate: YYYYMMDD 想定） 年接頭辞で候補を絞り込み日付範囲を評価 ---
      const s1all = await prisma.sengnr1.findMany({
        where: { OR: yearPrefixes.map((y) => ({ sdate: { startsWith: y } })) },
        select: { id: true, sdate: true, stime: true, timer: true, event: true },
      });
      const s1 = s1all.filter((r) => inRangeYmd(normalizeYyyymmdd(r.sdate)));
      let timerSum = 0;
      for (const r of s1) {
        if (r.timer != null) {
          timerSum += Number(r.timer);
        }
      }
      const lastS1 =
        s1.length === 0
          ? null
          : s1.reduce(
              (a, b) => (a.id > b.id ? a : b),
              s1[0] as (typeof s1)[0],
            );
      const lastS1Out: { sdate: string; stime: string | null } | null = lastS1
        ? { sdate: lastS1.sdate, stime: lastS1.stime }
        : null;

      // --- Kyouj2: rectime「YYYYMMDD HHMMSS」等。年接頭辞で候補を絞る ---
      const k2all = await prisma.kyouj2.findMany({
        where: { OR: yearPrefixes.map((y) => ({ rectime: { startsWith: y } })) },
        select: { id: true, rectime: true, realspd: true },
      });
      const k2 = k2all.filter((r) => {
        const d = normalizeYyyymmdd(r.rectime);
        return d.length === 8 && inRangeYmd(d);
      });
      const lastK2row =
        k2.length === 0 ? null : k2.reduce((a, b) => (a.id > b.id ? a : b), k2[0] as (typeof k2)[0]);
      const lastK2 = lastK2row?.rectime ?? null;

      // --- Sengnr3 ---
      const s3all = await prisma.sengnr3.findMany({
        where: { OR: yearPrefixes.map((y) => ({ sdate: { startsWith: y } })) },
        select: { id: true, sdate: true, sekitime: true, event: true },
      });
      const s3 = s3all.filter((r) => inRangeYmd(normalizeYyyymmdd(r.sdate)));
      const lastS3row =
        s3.length === 0 ? null : s3.reduce((a, b) => (a.id > b.id ? a : b), s3[0] as (typeof s3)[0]);
      const lastS3Out: { sdate: string; sekitime: string | null } | null = lastS3row
        ? { sdate: lastS3row.sdate, sekitime: lastS3row.sekitime }
        : null;

      const days = inclusiveDayCount(fromY, toY);
      /**
       * 簡易稼働率（MVP）定義:
       * 期間中の1号機(sengnr1)＋3号機(sengnr3)のイベント合計件数を、
       * 日数×基準件数(100)で割り 0〜1 にクリップ。Excel 帳票の厳密な照射/停止ロジックとは一致しません。
       */
      const eventLikeCount = s1.length + s3.length;
      const simpleUptimeRate = Math.min(1, eventLikeCount / Math.max(1, days * 100));

      res.json({
        data: {
          view: "uptime" as const,
          asOf: new Date().toISOString(),
          from: fromY,
          to: toY,
          series: {
            periodDays: days,
            sengnr1: {
              count: s1.length,
              timerSum,
              last: lastS1Out,
            },
            kyouj2: {
              count: k2.length,
              lastRectime: lastK2,
            },
            sengnr3: {
              count: s3.length,
              last: lastS3Out,
            },
            simpleUptimeRate,
          },
        },
      });
      return;
    }

    res.status(400).json({ error: "Invalid request" });
  } catch (e) {
    res.status(500).json({ error: e instanceof Error ? e.message : "Internal error" });
  }
});

export { shipmentSummaryRouter };
