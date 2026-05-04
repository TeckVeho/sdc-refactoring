import { Router } from "express";
import type { Prisma } from "@prisma/client";
import { z } from "zod";

import { prisma } from "../../lib/prisma.js";
import { requireAuth } from "../../middleware/auth.js";

const unissuedSearchRouter = Router();

/** DB 取得の上限（json に未発行フラグが無い簡易版では全件走査になりうるためキャップする） */
const FETCH_CAP = 8_000;

/** レスポンス行数の上限（クライアント負荷軽減） */
const RESPONSE_CAP = 2_000;

function firstQueryString(v: unknown): string | undefined {
  if (v == null || v === "") return undefined;
  if (Array.isArray(v)) {
    return typeof v[0] === "string" ? v[0] : undefined;
  }
  return typeof v === "string" ? v : undefined;
}

const querySchema = z.object({
  minUno: z.string().min(1).optional(),
  kaisyacd: z.string().min(1).optional(),
  maxShipDate: z.string().regex(/^\d{8}$/).optional(),
});

const patchBodySchema = z.object({
  jsonPatch: z.record(z.unknown()),
});

function normalizeYyyymmdd(s: string | null | undefined): string {
  if (s == null) return "";
  const digits = String(s).replace(/\D/g, "");
  return digits.length >= 8 ? digits.slice(0, 8) : digits;
}

/**
 * ExKeikakuX の json は列ベース移行でキー揺れがありうるため、
 * 代表的な別名を順に参照する。
 */
function pickField(obj: Record<string, unknown>, keys: string[]): unknown {
  for (const k of keys) {
    if (k in obj && obj[k] !== undefined) return obj[k];
  }
  const lowerKeys = new Set(keys.map((x) => x.toLowerCase()));
  for (const [ok, v] of Object.entries(obj)) {
    if (lowerKeys.has(ok.toLowerCase())) return v;
  }
  return undefined;
}

function strField(obj: Record<string, unknown>, keys: string[]): string {
  const v = pickField(obj, keys);
  if (v === null || v === undefined) return "";
  return String(v).trim();
}

/**
 * 報告書未発行検索の簡易判定（本番 Excel は zaiko/zaikor と UNION だが、MVP では ExKeikakuX.json のみ）。
 *
 * - 仕様上の「未発行」は syinji = 0（報告書発行フラグ）。
 * - 「発行済だが Fax 未送信」も一覧対象（syinji = 1 かつ houfax = 未）。
 * - json に syinji / hakkou 相当が無い行は、VBA 同等の結合結果が無いため
 *   「判定不能」として FETCH/RESPONSE の上限付きで返す（全件制限付きの意図）。
 */
function classifyUnissuedRow(j: Record<string, unknown>): "unissued" | "unknown" | "exclude" {
  const syinji = strField(j, ["syinji", "SYINJI", "hakkou", "HAKKOU", "報告書発行フラグ"]);
  const houfax = strField(j, ["houfax", "HOUFAX"]);

  if (syinji === "") {
    return "unknown";
  }
  if (syinji === "0") {
    return "unissued";
  }
  if (syinji === "1" && houfax === "未") {
    return "unissued";
  }
  return "exclude";
}

function parseJsonObject(raw: Prisma.JsonValue | null): Record<string, unknown> | null {
  if (raw === null) return null;
  if (typeof raw === "object" && !Array.isArray(raw)) {
    return raw as Record<string, unknown>;
  }
  try {
    const p = JSON.parse(JSON.stringify(raw)) as unknown;
    if (p && typeof p === "object" && !Array.isArray(p)) {
      return p as Record<string, unknown>;
    }
  } catch {
    /* ignore */
  }
  return null;
}

function shallowMergeJson(
  base: Prisma.JsonValue | null,
  patch: Record<string, unknown>,
): Prisma.InputJsonValue {
  const obj =
    base !== null && typeof base === "object" && !Array.isArray(base)
      ? { ...(base as Record<string, unknown>) }
      : {};
  return { ...obj, ...patch } as Prisma.InputJsonValue;
}

function rowMatchesQuery(
  j: Record<string, unknown>,
  q: { minUno?: string; kaisyacd?: string; maxShipDate?: string },
): boolean {
  if (q.minUno) {
    const uno = strField(j, ["uno", "UNO"]);
    if (uno < q.minUno) return false;
  }
  if (q.kaisyacd) {
    const k = strField(j, ["kaisyacd", "KAISYACD", "会社コード"]);
    if (k !== q.kaisyacd.trim()) return false;
  }
  if (q.maxShipDate) {
    const shipRaw = strField(j, ["syukkabi", "SYUKKABI", "最終出荷日"]);
    const ship = normalizeYyyymmdd(shipRaw);
    if (ship && ship > q.maxShipDate) return false;
  }
  return true;
}

unissuedSearchRouter.get("/", requireAuth, async (req, res) => {
  const parsed = querySchema.safeParse({
    minUno: firstQueryString(req.query.minUno),
    kaisyacd: firstQueryString(req.query.kaisyacd),
    maxShipDate: firstQueryString(req.query.maxShipDate),
  });
  if (!parsed.success) {
    res.status(400).json({ error: "Invalid query" });
    return;
  }
  const q = parsed.data;

  try {
    const dbRows = await prisma.exKeikakuX.findMany({
      take: FETCH_CAP,
      orderBy: { rowkey: "asc" },
      select: { rowkey: true, json: true },
    });

    const out: Array<{
      rowkey: string;
      uno: string;
      kaisyacd: string;
      kainame: string;
      sehncd: string;
      syinji: string;
      houfax: string;
      hyouji: string;
      rotno2: string;
      syukkabi: string;
      json: Record<string, unknown>;
    }> = [];

    let truncated = dbRows.length >= FETCH_CAP;

    for (const r of dbRows) {
      const j = parseJsonObject(r.json);
      if (!j) continue;

      const kind = classifyUnissuedRow(j);
      if (kind === "exclude") continue;
      if (kind === "unknown") {
        // 判定不能行も件数上限内で返す（仕様コメント参照）
      }
      if (!rowMatchesQuery(j, q)) continue;

      out.push({
        rowkey: r.rowkey,
        uno: strField(j, ["uno", "UNO"]),
        kaisyacd: strField(j, ["kaisyacd", "KAISYACD"]),
        kainame: strField(j, ["kainame", "KAINAME", "顧客名"]),
        sehncd: strField(j, ["sehncd", "SEHNCD", "製品コード"]),
        syinji: strField(j, ["syinji", "SYINJI", "hakkou", "HAKKOU"]),
        houfax: strField(j, ["houfax", "HOUFAX"]),
        hyouji: strField(j, ["hyouji", "HYOUJI", "表示設定"]),
        rotno2: strField(j, ["rotno2", "ROTNO2", "ロット番号"]),
        syukkabi: strField(j, ["syukkabi", "SYUKKABI"]),
        json: j,
      });

      if (out.length >= RESPONSE_CAP) {
        truncated = true;
        break;
      }
    }

    res.json({
      data: {
        rows: out,
        truncated,
      },
    });
  } catch (e) {
    res.status(500).json({ error: e instanceof Error ? e.message : "Server error" });
  }
});

unissuedSearchRouter.patch("/rows/:rowkey", requireAuth, async (req, res) => {
  const rowkey = typeof req.params.rowkey === "string" ? req.params.rowkey.trim() : "";
  if (!rowkey) {
    res.status(400).json({ error: "rowkey required" });
    return;
  }

  const body = patchBodySchema.safeParse(req.body);
  if (!body.success) {
    res.status(400).json({ error: "Invalid body (expected { jsonPatch: object })" });
    return;
  }

  const patch = body.data.jsonPatch;
  if (typeof patch !== "object" || patch === null || Array.isArray(patch)) {
    res.status(400).json({ error: "jsonPatch must be a plain object" });
    return;
  }

  try {
    const existing = await prisma.exKeikakuX.findUnique({
      where: { rowkey },
      select: { json: true },
    });
    if (!existing) {
      res.status(404).json({ error: "Not found" });
      return;
    }

    const merged = shallowMergeJson(existing.json, patch as Record<string, unknown>);

    await prisma.exKeikakuX.update({
      where: { rowkey },
      data: { json: merged },
    });

    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: e instanceof Error ? e.message : "Server error" });
  }
});

export { unissuedSearchRouter };
