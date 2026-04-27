import { Router } from "express";
import type { Prisma } from "@prisma/client";
import { z } from "zod";

import { prisma } from "../../lib/prisma.js";
import { requireAuth } from "../../middleware/auth.js";

const doseSearchRouter = Router();

/** Ex線量検索: 在庫=zaiko+syouj2 / 履歴=zaikor+syoujr2 だが、Prisma の Syouj2 は列が最小。MVP は zaiko/zaikor + Syouk1（受付に紐づく指定線量等）+ Syouj2（明細1件目） */
const MAX_LIMIT = 8_000;

const querySchema = z.object({
  source: z.enum(["zaiko", "zaikor"]).default("zaiko"),
  dateFrom: z.string().regex(/^\d{8}$/).optional(),
  dateTo: z.string().regex(/^\d{8}$/).optional(),
  /** 10桁1件 または 10-10 桁の範囲 */
  uno: z.string().optional(),
  unoFrom: z.string().min(1).optional(),
  unoTo: z.string().min(1).optional(),
  kaisyacd: z.string().min(1).optional(),
  sehncd: z.string().min(1).optional(),
  limit: z.coerce.number().int().min(1).max(MAX_LIMIT).default(2_000),
});

function firstQueryString(v: unknown): string | undefined {
  if (v == null || v === "") return undefined;
  if (Array.isArray(v)) {
    return typeof v[0] === "string" ? v[0] : undefined;
  }
  return typeof v === "string" ? v : undefined;
}

function onlyDigits(s: string): string {
  return s.replace(/\D/g, "");
}

/**
 * 受付番号: 単一（10桁）または範囲 `開始-終了`（仕様の Excel 相当）
 */
function parseUnoRange(
  uno: string | undefined,
  unoFrom: string | undefined,
  unoTo: string | undefined,
):
  | { type: "single"; value: string }
  | { type: "range"; from: string; to: string }
  | { type: "none" } {
  if (unoFrom && unoTo) {
    return { type: "range", from: onlyDigits(unoFrom).slice(0, 10), to: onlyDigits(unoTo).slice(0, 10) };
  }
  const raw = uno?.trim();
  if (raw) {
    if (raw.includes("-")) {
      const [a, b] = raw.split("-", 2).map((x) => onlyDigits(x).slice(0, 10));
      if (a.length === 10 && b.length === 10) {
        if (a <= b) return { type: "range", from: a, to: b };
        return { type: "range", from: b, to: a };
      }
    }
    const d = onlyDigits(raw).slice(0, 10);
    if (d.length === 10) {
      return { type: "single", value: d };
    }
  }
  return { type: "none" };
}

function countActiveConditions(x: {
  hasDateRange: boolean;
  hasUno: boolean;
  hasKaisya: boolean;
  hasSehn: boolean;
}): number {
  let n = 0;
  if (x.hasDateRange) n += 1;
  if (x.hasUno) n += 1;
  if (x.hasKaisya) n += 1;
  if (x.hasSehn) n += 1;
  return n;
}

type DoseSearchRow = {
  uno: string;
  kaisyacd: string;
  sehncd: string;
  syouso: string | null;
  kainame: string | null;
  nyukabi: string | null;
  nouki: string | null;
  pass: string | null;
  source: "zaiko" | "zaikor";
  siteisn: string | null;
  senritu: string | null;
  syokind: string | null;
  syostat: string | null;
  syosuu: number | null;
  zhansuu: number | null;
  /** Syouj2 先頭1件 */
  jno: string | null;
  j2meisai: string | null;
  j2status: string | null;
};

type HeaderRow = {
  uno: string;
  kaisyacd: string;
  sehncd: string;
  syouso: string | null;
  kainame: string | null;
  nyukabi: string | null;
  nouki: string | null;
  pass: string | null;
};

function buildHeaderWhere(w: {
  hasDateRange: boolean;
  dateFrom: string;
  dateTo: string;
  ur: ReturnType<typeof parseUnoRange>;
  kaisyacdPadded: string | undefined;
  sehncdPadded: string | undefined;
}): Prisma.ZaikoWhereInput {
  const base: Prisma.ZaikoWhereInput = {};

  if (w.hasDateRange) {
    base.nyukabi = { gte: w.dateFrom, lte: w.dateTo };
  }

  if (w.ur.type === "single") {
    base.uno = w.ur.value;
  } else if (w.ur.type === "range") {
    base.uno = { gte: w.ur.from, lte: w.ur.to };
  }

  if (w.kaisyacdPadded) {
    base.kaisyacd = w.kaisyacdPadded;
  }
  if (w.sehncdPadded) {
    base.sehncd = w.sehncdPadded;
  }

  return base;
}

function mapHeader(z: "zaiko" | "zaikor", h: HeaderRow) {
  return {
    uno: h.uno,
    kaisyacd: h.kaisyacd,
    sehncd: h.sehncd,
    syouso: h.syouso,
    kainame: h.kainame,
    nyukabi: h.nyukabi,
    nouki: h.nouki,
    pass: h.pass,
    source: z,
  } as const;
}

/** GET /api/dosimetry/dose-search */
doseSearchRouter.get("/", requireAuth, async (req, res) => {
  const parsed = querySchema.safeParse({
    source: firstQueryString(req.query.source),
    dateFrom: firstQueryString(req.query.dateFrom),
    dateTo: firstQueryString(req.query.dateTo),
    uno: firstQueryString(req.query.uno),
    unoFrom: firstQueryString(req.query.unoFrom),
    unoTo: firstQueryString(req.query.unoTo),
    kaisyacd: firstQueryString(req.query.kaisyacd),
    sehncd: firstQueryString(req.query.sehncd),
    limit: firstQueryString(req.query.limit),
  });
  if (!parsed.success) {
    res.status(400).json({ error: "Invalid query" });
    return;
  }
  const q = parsed.data;
  const hasDateRange = Boolean(q.dateFrom && q.dateTo);
  const ur = parseUnoRange(q.uno, q.unoFrom, q.unoTo);
  const hasUno = ur.type !== "none";
  const kaisyacdPadded = q.kaisyacd
    ? q.kaisyacd.trim().replace(/\D/g, "").padStart(4, "0").slice(-4)
    : undefined;
  const sehncdPadded = q.sehncd
    ? q.sehncd.trim().replace(/\D/g, "").padStart(3, "0").slice(-3)
    : undefined;
  const hasKaisya = Boolean(kaisyacdPadded);
  const hasSehn = Boolean(sehncdPadded);

  const nCond = countActiveConditions({ hasDateRange, hasUno, hasKaisya, hasSehn });
  if (nCond < 2) {
    res.status(400).json({
      error: "MVP: 日付範囲・受付番号・会社コード・製品コードのうち2条件以上を指定してください",
    });
    return;
  }

  if (hasDateRange && q.dateFrom! > q.dateTo!) {
    res.status(400).json({ error: "日付範囲が不正です" });
    return;
  }
  if (ur.type === "range" && ur.from > ur.to) {
    res.status(400).json({ error: "受付番号の範囲が不正です" });
    return;
  }

  const whereZaiko = buildHeaderWhere({
    hasDateRange,
    dateFrom: q.dateFrom ?? "",
    dateTo: q.dateTo ?? "",
    ur,
    kaisyacdPadded: hasKaisya ? kaisyacdPadded : undefined,
    sehncdPadded: hasSehn ? sehncdPadded : undefined,
  });
  const whereZaikor = whereZaiko as Prisma.ZaikorWhereInput;
  const take = q.limit + 1;
  const orderBy: Prisma.ZaikoOrderByWithRelationInput = { uno: "asc" };

  try {
    const [headers, kansokuMasters] = await Promise.all([
      q.source === "zaiko"
        ? prisma.zaiko.findMany({ where: whereZaiko, orderBy, take, select: {
            uno: true,
            kaisyacd: true,
            sehncd: true,
            syouso: true,
            kainame: true,
            nyukabi: true,
            nouki: true,
            pass: true,
          } })
        : prisma.zaikor.findMany({ where: whereZaikor, orderBy, take, select: {
            uno: true,
            kaisyacd: true,
            sehncd: true,
            syouso: true,
            kainame: true,
            nyukabi: true,
            nouki: true,
            pass: true,
          } }),
      prisma.kansoku.findMany({
        take: 1_000,
        orderBy: { kid: "asc" },
        select: { kid: true, sokutei: true, bikou: true },
      }),
    ]);

    const truncated = headers.length > q.limit;
    const slice = truncated ? headers.slice(0, q.limit) : headers;
    const unos = slice.map((r) => r.uno);
    if (unos.length === 0) {
      res.json({ data: { rows: [] as DoseSearchRow[], truncated, kansokuMasters } });
      return;
    }

    const [skList, j2all] = await Promise.all([
      prisma.syouk1.findMany({
        where: { uno: { in: unos } },
        select: {
          uno: true,
          siteisn: true,
          senritu: true,
          syokind: true,
          syostat: true,
          syosuu: true,
          zhansuu: true,
        },
      }),
      prisma.syouj2.findMany({
        where: { uno: { in: unos } },
        orderBy: { jno: "asc" },
        select: { jno: true, uno: true, meisai: true, status: true },
      }),
    ]);

    const byUnoS = new Map(skList.map((r) => [r.uno, r]));
    const byUnoJ2 = new Map<string, (typeof j2all)[0]>();
    for (const r of j2all) {
      if (!r.uno) continue;
      if (!byUnoJ2.has(r.uno)) byUnoJ2.set(r.uno, r);
    }

    const rows: DoseSearchRow[] = slice.map((h) => {
      const s = byUnoS.get(h.uno);
      const j2 = byUnoJ2.get(h.uno);
      return {
        ...mapHeader(q.source, h),
        siteisn: s?.siteisn ?? null,
        senritu: s?.senritu != null ? String(s.senritu) : null,
        syokind: s?.syokind ?? null,
        syostat: s?.syostat ?? null,
        syosuu: s?.syosuu ?? null,
        zhansuu: s?.zhansuu ?? null,
        jno: j2?.jno ?? null,
        j2meisai: j2?.meisai ?? null,
        j2status: j2?.status ?? null,
      };
    });

    res.json({ data: { rows, truncated, kansokuMasters } });
  } catch (e) {
    res.status(500).json({ error: e instanceof Error ? e.message : "Internal error" });
  }
});

export { doseSearchRouter };
