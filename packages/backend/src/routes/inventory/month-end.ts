import { Router } from "express";
import { z } from "zod";

import { prisma } from "../../lib/prisma.js";
import { requireAuth } from "../../middleware/auth.js";

const monthEndRouter = Router();

function firstQueryString(v: unknown): string | undefined {
  if (v == null || v === "") return undefined;
  if (Array.isArray(v)) {
    return typeof v[0] === "string" ? v[0] : undefined;
  }
  return typeof v === "string" ? v : undefined;
}

function normalizeYyyymmdd(s: string | null | undefined): string {
  if (s == null) return "";
  const digits = String(s).replace(/\D/g, "");
  return digits.length >= 8 ? digits.slice(0, 8) : digits;
}

function toMachineId(s: string): "1" | "2" | "3" | null {
  const t = s.trim();
  if (t === "1" || t === "01") return "1";
  if (t === "2" || t === "02") return "2";
  if (t === "3" || t === "03") return "3";
  if (/^[123]$/.test(t)) return t as "1" | "2" | "3";
  return null;
}

type InvBucket = "1" | "2" | "3" | "EB" | "LOCA" | "other";

function inventoryBucket(syouso: string | null | undefined): InvBucket {
  if (syouso == null || syouso.trim() === "") return "other";
  const t = syouso.trim();
  const u = t.toUpperCase();
  if (u.includes("LOCA")) return "LOCA";
  if (u === "EB" || u === "4" || t === "eb") return "EB";
  const m = toMachineId(t);
  if (m) return m;
  return "other";
}

function bucketLabel(b: InvBucket): string {
  switch (b) {
    case "1":
      return "1号機";
    case "2":
      return "2号機";
    case "3":
      return "3号機";
    case "EB":
      return "EB";
    case "LOCA":
      return "LOCA";
    default:
      return "その他";
  }
}

/** 締日時点で入荷済みとみなす行（入荷日が取れない行は含める） */
function includeRowAsOf(nyukabi: string | null | undefined, asOf: string): boolean {
  const n = normalizeYyyymmdd(nyukabi);
  if (n.length < 8) return true;
  return n <= asOf;
}

function yearPrefixesForAsOf(asOf: string): string[] {
  const y = parseInt(asOf.slice(0, 4), 10);
  if (!Number.isFinite(y)) return [];
  return [String(y), String(y - 1)];
}

const asOfSchema = z.string().regex(/^\d{8}$/);

monthEndRouter.get("/", requireAuth, async (req, res) => {
  const rawAsOf = firstQueryString(req.query.asOf) ?? firstQueryString(req.query.closingDate);
  const parsedAsOf = asOfSchema.safeParse(rawAsOf ?? "");
  if (!parsedAsOf.success) {
    res.status(400).json({ error: "asOf or closingDate (YYYYMMDD) is required" });
    return;
  }
  const asOf = parsedAsOf.data;

  const prefixes = yearPrefixesForAsOf(asOf);
  const yearOr = [{ nyukabi: null }, ...prefixes.map((p) => ({ nyukabi: { startsWith: p } }))];

  type RowSelect = {
    uno: string;
    kaisyacd: string;
    sehncd: string;
    syouso: string | null;
    kainame: string | null;
    nyukabi: string | null;
    nyukasu: number | null;
    pass: string | null;
    incnt: number | null;
    nouki: string | null;
  };

  type Line = RowSelect & {
    source: "zaiko" | "zaikor";
    bucket: InvBucket;
    bucketLabel: string;
  };

  try {
    const [zaikoRows, zaikorRows] = await Promise.all([
      prisma.zaiko.findMany({
        where: { OR: yearOr },
        select: {
          uno: true,
          kaisyacd: true,
          sehncd: true,
          syouso: true,
          kainame: true,
          nyukabi: true,
          nyukasu: true,
          pass: true,
          incnt: true,
          nouki: true,
        },
      }),
      prisma.zaikor.findMany({
        where: { OR: yearOr },
        select: {
          uno: true,
          kaisyacd: true,
          sehncd: true,
          syouso: true,
          kainame: true,
          nyukabi: true,
          nyukasu: true,
          pass: true,
          incnt: true,
          nouki: true,
        },
      }),
    ]);

    const byUno = new Map<string, Line>();

    for (const r of zaikoRows) {
      if (!includeRowAsOf(r.nyukabi, asOf)) continue;
      const b = inventoryBucket(r.syouso);
      byUno.set(r.uno, {
        ...r,
        source: "zaiko",
        bucket: b,
        bucketLabel: bucketLabel(b),
      });
    }

    for (const r of zaikorRows) {
      if (!includeRowAsOf(r.nyukabi, asOf)) continue;
      if (byUno.has(r.uno)) continue;
      const b = inventoryBucket(r.syouso);
      byUno.set(r.uno, {
        ...r,
        source: "zaikor",
        bucket: b,
        bucketLabel: bucketLabel(b),
      });
    }

    const lines = Array.from(byUno.values()).sort((a, b) => {
      const c = a.kaisyacd.localeCompare(b.kaisyacd, "ja");
      if (c !== 0) return c;
      return a.sehncd.localeCompare(b.sehncd, "ja");
    });

    const order: InvBucket[] = ["1", "2", "3", "EB", "LOCA", "other"];
    const summaryAgg = new Map<InvBucket, { rowCount: number; totalNyukasu: number }>();
    for (const b of order) {
      summaryAgg.set(b, { rowCount: 0, totalNyukasu: 0 });
    }
    for (const line of lines) {
      const cur = summaryAgg.get(line.bucket)!;
      cur.rowCount += 1;
      cur.totalNyukasu += line.nyukasu ?? 0;
    }

    const summary = order.map((b) => {
      const s = summaryAgg.get(b)!;
      return {
        bucket: b,
        label: bucketLabel(b),
        rowCount: s.rowCount,
        /** MVP: スキーマに在庫数量が無いため入荷数合計を装置別の代理指標とする */
        stockProxyNyukasu: s.totalNyukasu,
        /** 仕掛金額は未実装（本番は照射仕掛シート相当） */
        wipAmount: 0,
      };
    });

    const totalNyukasu = lines.reduce((s, l) => s + (l.nyukasu ?? 0), 0);

    res.json({
      data: {
        asOf,
        closingDate: asOf,
        yearMonth: `${asOf.slice(0, 4)}-${asOf.slice(4, 6)}`,
        generatedAt: new Date().toISOString(),
        mvpNote:
          "簡易版: zaiko ∪ zaikor（uno 重複時は zaiko 優先）。締日以降の入荷日付き行は除外。装置別は syouso から判定。在庫数量・金額は DB スキーマ未対応のため入荷数を代理集計。LOCA 手入力・仕掛按分は含みません。",
        lines,
        summary,
        totals: { rowCount: lines.length, totalNyukasu },
      },
    });
  } catch (e) {
    res.status(500).json({ error: e instanceof Error ? e.message : "Internal error" });
  }
});

export { monthEndRouter };
