import { Router } from "express";
import { z } from "zod";
import type { Prisma } from "@prisma/client";

import { prisma } from "../../lib/prisma.js";
import { requireAuth } from "../../middleware/auth.js";

const productionRouter = Router();

const querySchema = z.object({
  syouso: z.enum(["1", "2", "3", "EB"]).optional(),
  limit: z.coerce.number().int().min(1).max(2000).default(200),
});

function firstQueryString(v: unknown): string | undefined {
  if (v == null || v === "") return undefined;
  if (Array.isArray(v)) {
    return typeof v[0] === "string" ? v[0] : undefined;
  }
  return typeof v === "string" ? v : undefined;
}

/** クエリ `syouso` を zaiko.syouso の表記揺れに合わせる */
function syousoWhere(
  q: "1" | "2" | "3" | "EB",
): Prisma.ZaikoWhereInput {
  if (q === "EB") {
    return {
      OR: [
        { syouso: { equals: "EB" } },
        { syouso: { equals: "eb" } },
        { syouso: { equals: "4" } },
      ],
    };
  }
  return {
    OR: [
      { syouso: { equals: q } },
      { syouso: { equals: `0${q}` } },
      { syouso: { equals: `${q}.0` } },
    ],
  };
}

type KeikakuFields = {
  syukkabi?: string;
  bikou1?: string;
  kakunin?: string;
  syuhouhou?: string;
};

type YoyakuFields = {
  yoyakubi?: string;
  yoyakuno?: string;
  bikou?: string;
};

function strVal(o: Record<string, unknown>, ...keys: string[]): string | undefined {
  for (const k of keys) {
    const v = o[k];
    if (v === undefined || v === null) continue;
    const s = String(v).trim();
    if (s !== "") return s.replace(/\s+/g, " ");
  }
  return undefined;
}

function readKeikakuObject(o: Record<string, unknown>, sink: Map<string, KeikakuFields>): void {
  const rawUno = o.uno ?? o.UNO;
  if (rawUno === undefined || rawUno === null) return;
  const uno = String(rawUno).trim();
  if (!uno) return;
  const cur = sink.get(uno) ?? {};
  const next: KeikakuFields = { ...cur };
  const a = strVal(o, "syukkabi", "SYUKKABI");
  const b = strVal(o, "bikou1", "BIKOU1");
  const c = strVal(o, "kakunin", "KAKUNIN");
  const d = strVal(o, "syuhouhou", "SYUHHOUHOU", "syukkahou", "SYUKKAHOU");
  if (a != null) next.syukkabi = a;
  if (b != null) next.bikou1 = b;
  if (c != null) next.kakunin = c;
  if (d != null) next.syuhouhou = d;
  sink.set(uno, next);
}

function readKeikakuJson(j: Prisma.JsonValue, sink: Map<string, KeikakuFields>): void {
  if (j === null || typeof j !== "object" || Array.isArray(j)) return;
  const o = j as Record<string, unknown>;
  readKeikakuObject(o, sink);
  try {
    const s = JSON.stringify(j);
    const p = JSON.parse(s) as unknown;
    if (p && typeof p === "object" && !Array.isArray(p)) {
      readKeikakuObject(p as Record<string, unknown>, sink);
    }
  } catch {
    // ignore
  }
}

async function buildKeikakuByUno(): Promise<Map<string, KeikakuFields>> {
  const rows = await prisma.exKeikakuX.findMany({ select: { json: true } });
  const map = new Map<string, KeikakuFields>();
  for (const r of rows) {
    const j = r.json;
    if (j === null) continue;
    if (typeof j === "object" && j !== null && !Array.isArray(j)) {
      readKeikakuJson(j, map);
    } else {
      try {
        const s = JSON.stringify(j);
        const p = JSON.parse(s) as unknown;
        if (p && typeof p === "object" && !Array.isArray(p)) {
          readKeikakuJson(p as Prisma.JsonValue, map);
        }
      } catch {
        // ignore
      }
    }
  }
  return map;
}

function readYoyakuObject(o: Record<string, unknown>, sink: Map<string, YoyakuFields>): void {
  const rawUno = o.uno ?? o.UNO;
  if (rawUno === undefined || rawUno === null) return;
  const uno = String(rawUno).trim();
  if (!uno) return;
  const cur = sink.get(uno) ?? {};
  const next: YoyakuFields = { ...cur };
  const a = strVal(o, "yoyakubi", "YOYAKUBI");
  const b = strVal(o, "yoyakuno", "YOYAKUNO");
  const c = strVal(o, "bikou", "BIKOU");
  if (a != null) next.yoyakubi = a;
  if (b != null) next.yoyakuno = b;
  if (c != null) next.bikou = c;
  sink.set(uno, next);
}

function readYoyakuJson(j: Prisma.JsonValue, sink: Map<string, YoyakuFields>): void {
  if (j === null || typeof j !== "object" || Array.isArray(j)) return;
  readYoyakuObject(j as Record<string, unknown>, sink);
}

async function buildYoyakuByUno(): Promise<Map<string, YoyakuFields>> {
  const rows = await prisma.exYoyakuX.findMany({ select: { json: true } });
  const map = new Map<string, YoyakuFields>();
  for (const r of rows) {
    const j = r.json;
    if (j === null) continue;
    if (typeof j === "object" && j !== null && !Array.isArray(j)) {
      readYoyakuJson(j, map);
    } else {
      try {
        const s = JSON.stringify(j);
        const p = JSON.parse(s) as unknown;
        if (p && typeof p === "object" && !Array.isArray(p)) {
          readYoyakuObject(p as Record<string, unknown>, map);
        }
      } catch {
        // ignore
      }
    }
  }
  return map;
}

export type ProductionListRow = {
  uno: string;
  kaisyacd: string;
  sehncd: string;
  syouso: string | null;
  kainame: string | null;
  nyukabi: string | null;
  nouki: string | null;
  pass: string | null;
  nyukasu: number | null;
  incnt: number | null;
  /** 指定線量（Syouk1） */
  siteisn: string | null;
  /** 未照射数に相当（Syouk1.zhansuu、未設定時は null） */
  misyousu: number | null;
  /** 入数と解釈した照射側数量（Syouk1.syosuu） */
  syosuu: number | null;
  /** 測定・上限系（Syouk1.senritu、仕様 jyougsn に近い列が無い場合の代替） */
  senritu: string | null;
  syokind: string | null;
  syostat: string | null;
  /** 生産計画 JSON（ExKeikakuX） */
  syukkabi: string | null;
  bikou1: string | null;
  kakunin: string | null;
  syuhouhou: string | null;
  /** 予約 JSON（ExYoyakuX） */
  yoyakubi: string | null;
  yoyakuno: string | null;
  yoyakuBikou: string | null;
};

/** GET /api/irradiation/production */
productionRouter.get("/", requireAuth, async (req, res) => {
  const parsed = querySchema.safeParse({
    syouso: firstQueryString(req.query.syouso),
    limit: firstQueryString(req.query.limit),
  });
  if (!parsed.success) {
    res.status(400).json({ error: "Invalid query" });
    return;
  }
  const { syouso: syousoQ, limit } = parsed.data;

  const where: Prisma.ZaikoWhereInput = syousoQ
    ? syousoWhere(syousoQ)
    : {};

  try {
    const [total, zaRows, kei, yoy] = await Promise.all([
      prisma.zaiko.count({ where }),
      prisma.zaiko.findMany({
        where,
        orderBy: { uno: "asc" },
        take: limit,
        select: {
          uno: true,
          kaisyacd: true,
          sehncd: true,
          syouso: true,
          kainame: true,
          nouki: true,
          pass: true,
          nyukabi: true,
          nyukasu: true,
          incnt: true,
        },
      }),
      buildKeikakuByUno(),
      buildYoyakuByUno(),
    ]);

    const unos = zaRows.map((z) => z.uno);
    const syList =
      unos.length === 0
        ? []
        : await prisma.syouk1.findMany({
            where: { uno: { in: unos } },
            select: {
              uno: true,
              siteisn: true,
              syosuu: true,
              zhansuu: true,
              senritu: true,
              syokind: true,
              syostat: true,
            },
          });
    const byUnoSyouk = new Map(syList.map((s) => [s.uno, s]));

    const rows: ProductionListRow[] = zaRows.map((z) => {
      const s = byUnoSyouk.get(z.uno);
      const k = kei.get(z.uno);
      const y = yoy.get(z.uno);
      const sen =
        s?.senritu != null
          ? String(s.senritu)
          : null;
      return {
        uno: z.uno,
        kaisyacd: z.kaisyacd,
        sehncd: z.sehncd,
        syouso: z.syouso,
        kainame: z.kainame,
        nyukabi: z.nyukabi,
        nouki: z.nouki,
        pass: z.pass,
        nyukasu: z.nyukasu,
        incnt: z.incnt,
        siteisn: s?.siteisn ?? null,
        misyousu: s?.zhansuu ?? null,
        syosuu: s?.syosuu ?? null,
        senritu: sen,
        syokind: s?.syokind ?? null,
        syostat: s?.syostat ?? null,
        syukkabi: k?.syukkabi ?? null,
        bikou1: k?.bikou1 ?? null,
        kakunin: k?.kakunin ?? null,
        syuhouhou: k?.syuhouhou ?? null,
        yoyakubi: y?.yoyakubi ?? null,
        yoyakuno: y?.yoyakuno ?? null,
        yoyakuBikou: y?.bikou ?? null,
      };
    });

    res.json({
      data: {
        rows,
        meta: { total },
      },
    });
  } catch (e) {
    res.status(500).json({ error: e instanceof Error ? e.message : "Internal error" });
  }
});

export { productionRouter };
