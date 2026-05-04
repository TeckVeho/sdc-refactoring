import { Router } from "express";
import type { Prisma } from "@prisma/client";
import { z } from "zod";

import { prisma } from "../../lib/prisma.js";
import { requireAuth } from "../../middleware/auth.js";

const ric3RepackRouter = Router();

/** ExR3Stg に保存するワークスペース ID（Excel ExRic3詰替作業相当） */
const RIC3_REPACK_STG_ID = "ric3-repack";

const ric3RepackItemSchema = z.object({
  id: z.string(),
  lastRepackDate: z.string(),
  irradiationStatus: z.string(),
  irradiationCode: z.string(),
  doseMeterNo: z.string(),
  receiptNo: z.string(),
  companyName: z.string(),
  dueDate: z.string(),
  shipDate: z.string(),
  passCount: z.string(),
  note: z.string(),
  checked: z.boolean(),
});

const ric3ProductRowSchema = z.object({
  kaisyacd: z.string(),
  sehncd: z.string(),
  composite: z.string().optional(),
  companyName: z.string(),
  productName: z.string(),
  tumikae: z.string(),
  originalTumikae: z.string(),
});

const ric3StateSchema = z.object({
  version: z.literal(1),
  items: z.array(ric3RepackItemSchema),
  products: z.array(ric3ProductRowSchema),
});

export type Ric3RepackState = z.infer<typeof ric3StateSchema>;

function defaultState(): Ric3RepackState {
  return { version: 1, items: [], products: [] };
}

function parseState(raw: Prisma.JsonValue | null): Ric3RepackState {
  if (raw === null) return defaultState();
  try {
    const p = JSON.parse(JSON.stringify(raw)) as unknown;
    const r = ric3StateSchema.safeParse(p);
    if (r.success) return r.data;
  } catch {
    /* ignore */
  }
  return defaultState();
}

async function loadState(): Promise<Ric3RepackState> {
  const row = await prisma.exR3Stg.findUnique({
    where: { stgId: RIC3_REPACK_STG_ID },
    select: { json: true },
  });
  return parseState(row?.json ?? null);
}

async function saveState(state: Ric3RepackState): Promise<void> {
  await prisma.exR3Stg.upsert({
    where: { stgId: RIC3_REPACK_STG_ID },
    create: { stgId: RIC3_REPACK_STG_ID, json: state as Prisma.InputJsonValue },
    update: { json: state as Prisma.InputJsonValue },
  });
}

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

function right4(s: string | null | undefined): string {
  if (s == null) return "";
  const t = String(s).trim();
  return t.length <= 4 ? t : t.slice(-4);
}

function stripKabushiki(name: string): string {
  return name.replace(/株式会社/g, "").trim();
}

/** 数字8桁を mm/dd 表示（Excel 仕様の納期表示に近づける） */
function noukiToMmdd(raw: string | null | undefined): string {
  if (raw == null) return "";
  const d = String(raw).replace(/\D/g, "");
  if (d.length >= 8) {
    const mm = d.slice(4, 6);
    const dd = d.slice(6, 8);
    return `${mm}/${dd}`;
  }
  return String(raw).trim();
}

function normalizeShipYmd(raw: string): string {
  const t = raw.replace(/\D/g, "");
  return t.length >= 8 ? t.slice(0, 8) : "";
}

/** 出荷日(YYYYMMDD) の前日を YYYY/MM/DD で返す（無ければ空） */
function lastRepackFromShipYmd(ymd: string): string {
  if (ymd.length !== 8) return "";
  const y = Number(ymd.slice(0, 4));
  const m = Number(ymd.slice(4, 6)) - 1;
  const d = Number(ymd.slice(6, 8));
  const dt = new Date(Date.UTC(y, m, d));
  dt.setUTCDate(dt.getUTCDate() - 1);
  const yy = dt.getUTCFullYear();
  const mm = String(dt.getUTCMonth() + 1).padStart(2, "0");
  const dd = String(dt.getUTCDate()).padStart(2, "0");
  return `${yy}/${mm}/${dd}`;
}

/** 照射状況コード（仕様 syoush_f）→ 表示 */
function irradiationLabelFromCode(code: string): { label: string; codeOut: string } {
  const c = code.trim();
  if (c === "i" || c === "0" || c === "C") return { label: "未", codeOut: c };
  if (c === "1") return { label: "中", codeOut: c };
  if (c === "2") return { label: "済", codeOut: c };
  if (c === "") return { label: "未", codeOut: "" };
  return { label: "??", codeOut: c };
}

/**
 * syoukj3.syokind または日付から照射状況を推定（列が Oracle 原版と異なる場合のフォールバック）
 */
function irradiationFromRow(
  syokind: string | null | undefined,
  sdate: string | null | undefined,
  edate: string | null | undefined,
): { label: string; codeOut: string } {
  const sk = (syokind ?? "").trim();
  if (sk === "i" || sk === "0" || sk === "C" || sk === "1" || sk === "2") {
    return irradiationLabelFromCode(sk);
  }
  if (sk !== "") {
    return irradiationLabelFromCode(sk);
  }
  const e = (edate ?? "").replace(/\D/g, "");
  if (e.length >= 8) return { label: "済", codeOut: "2" };
  const s = (sdate ?? "").replace(/\D/g, "");
  if (s.length >= 8) return { label: "中", codeOut: "1" };
  return { label: "未", codeOut: "" };
}

type KeikakuR3 = Map<string, { shipDate: string; note: string }>;

function readKeikakuJsonR3(j: Prisma.JsonValue, sink: KeikakuR3): void {
  if (j === null || typeof j !== "object" || Array.isArray(j)) return;
  const o = j as Record<string, unknown>;
  const souti = String(o.souti ?? o.SOUTI ?? "").trim();
  if (souti !== "3" && souti !== "3.0") return;
  const rawUno = o.uno ?? o.UNO;
  if (rawUno === undefined || rawUno === null) return;
  const uno = String(rawUno).trim();
  const ship = String(o.syukkabi ?? o.SYUKKABI ?? "")
    .trim()
    .replace(/\s+/g, " ");
  const note = String(o.bikou1 ?? o.BIKOU1 ?? "")
    .trim()
    .replace(/\s+/g, " ");
  sink.set(uno, { shipDate: ship, note });
}

async function buildKeikakuUnoMapR3(): Promise<KeikakuR3> {
  const rows = await prisma.exKeikakuX.findMany({ select: { json: true } });
  const map: KeikakuR3 = new Map();
  for (const r of rows) {
    const j = r.json;
    if (j === null) continue;
    if (typeof j === "object" && !Array.isArray(j)) {
      readKeikakuJsonR3(j, map);
    } else {
      try {
        const p = JSON.parse(JSON.stringify(j)) as unknown;
        if (p && typeof p === "object" && !Array.isArray(p)) {
          readKeikakuJsonR3(p as Prisma.JsonValue, map);
        }
      } catch {
        /* ignore */
      }
    }
  }
  return map;
}

function parseExSeihinzJson(j: Prisma.JsonValue): { kaisyacd: string; sehncd: string; tumikae: string } | null {
  if (j === null || typeof j !== "object" || Array.isArray(j)) return null;
  const o = j as Record<string, unknown>;
  const kaisyacd = strField(o, ["kaisyacd", "KAISYACD"]);
  const sehncd = strField(o, ["sehncd", "SEHNCD", "製品コード"]);
  const tumikae = strField(o, ["tumikae", "TUMIKAE", "詰替え要不要"]);
  if (!kaisyacd || !sehncd) return null;
  return { kaisyacd, sehncd, tumikae: tumikae || "0" };
}

function tumikaeKey(kaisyacd: string, sehncd: string): string {
  return `${kaisyacd.trim()}\t${sehncd.trim()}`;
}

/** ExSeihinz から tumikae===1 の (kaisyacd,sehncd) 集合。0件ならフィルタ無し。 */
async function loadTumikaeRequiredKeys(): Promise<Set<string> | null> {
  const rows = await prisma.exSeihinz.findMany({
    select: { rowkey: true, json: true },
    take: 20_000,
  });
  const keys = new Set<string>();
  for (const r of rows) {
    const parsed = r.json ? parseExSeihinzJson(r.json) : null;
    if (parsed && parsed.tumikae === "1") {
      keys.add(tumikaeKey(parsed.kaisyacd, parsed.sehncd));
    }
  }
  if (keys.size === 0) return null;
  return keys;
}

async function buildItemsFromDb(): Promise<Ric3RepackState["items"]> {
  const [syoukRows, keikaku, tumikaeKeys] = await Promise.all([
    prisma.syoukj3.findMany({ take: 500 }),
    buildKeikakuUnoMapR3(),
    loadTumikaeRequiredKeys(),
  ]);

  const zaikoList = await prisma.zaiko.findMany({
    where: {
      uno: { in: syoukRows.map((s) => s.uno) },
    },
  });
  const zaikoByUno = new Map(zaikoList.map((z) => [z.uno, z]));

  const items: Ric3RepackState["items"] = [];
  let seq = 0;
  for (const s of syoukRows) {
    const z = zaikoByUno.get(s.uno);
    if (!z) continue;
    if (tumikaeKeys && !tumikaeKeys.has(tumikaeKey(z.kaisyacd, z.sehncd))) {
      continue;
    }
    const k = keikaku.get(s.uno.trim());
    const shipRaw = k?.shipDate ?? "";
    const shipYmd = normalizeShipYmd(shipRaw);
    const { label: irrLabel, codeOut: irrCode } = irradiationFromRow(s.syokind, s.sdate, s.edate);
    const company = stripKabushiki((s.kainame ?? z.kainame ?? "").trim());
    const dose = right4(s.syono);
    const receipt = right4(s.uno);
    items.push({
      id: `item-${s.uno}-${seq++}`,
      lastRepackDate: lastRepackFromShipYmd(shipYmd),
      irradiationStatus: irrLabel,
      irradiationCode: irrCode,
      doseMeterNo: dose,
      receiptNo: receipt,
      companyName: company || "—",
      dueDate: noukiToMmdd(z.nouki),
      shipDate: shipYmd ? `${shipYmd.slice(0, 4)}/${shipYmd.slice(4, 6)}/${shipYmd.slice(6, 8)}` : shipRaw,
      passCount: (z.pass ?? "").trim() || "—",
      note: (k?.note ?? s.bikou ?? "").trim() || "—",
      checked: false,
    });
  }

  const statusOrder = (x: string) => (x === "未" ? 0 : x === "中" ? 1 : x === "済" ? 2 : 3);
  items.sort((a, b) => {
    const o = statusOrder(b.irradiationStatus) - statusOrder(a.irradiationStatus);
    if (o !== 0) return o;
    const d = a.lastRepackDate.localeCompare(b.lastRepackDate);
    if (d !== 0) return d;
    return a.doseMeterNo.localeCompare(b.doseMeterNo);
  });
  return items;
}

async function buildProductsFromDb(): Promise<Ric3RepackState["products"]> {
  const zaikoR3 = await prisma.zaiko.findMany({
    where: { syouso: { in: ["3", "03"] } },
    select: { kaisyacd: true, sehncd: true },
    distinct: ["kaisyacd", "sehncd"],
    take: 3000,
  });
  const keySet = new Set(zaikoR3.map((z) => tumikaeKey(z.kaisyacd, z.sehncd)));

  const [sehmstRows, tokumstRows, exRows] = await Promise.all([
    prisma.sehmst.findMany({ take: 8000 }),
    prisma.tokumst.findMany(),
    prisma.exSeihinz.findMany({ select: { rowkey: true, json: true }, take: 20_000 }),
  ]);
  const tokMap = new Map(tokumstRows.map((t) => [t.kaisyacd, t]));
  const exByKey = new Map<string, { tumikae: string }>();
  for (const r of exRows) {
    const p = r.json ? parseExSeihinzJson(r.json) : null;
    if (p) {
      exByKey.set(tumikaeKey(p.kaisyacd, p.sehncd), { tumikae: p.tumikae || "0" });
    }
  }

  const products: Ric3RepackState["products"] = [];
  for (const s of sehmstRows) {
    const k = tumikaeKey(s.kaisyacd, s.sehncd);
    if (keySet.size > 0 && !keySet.has(k)) continue;
    const t = tokMap.get(s.kaisyacd);
    const ex = exByKey.get(k);
    const tm = ex?.tumikae ?? "0";
    products.push({
      kaisyacd: s.kaisyacd,
      sehncd: s.sehncd,
      composite: `${s.kaisyacd}${s.sehncd}`,
      companyName: (t?.coname ?? "").trim(),
      productName: (s.seiname ?? "").trim(),
      tumikae: tm,
      originalTumikae: tm,
    });
  }
  products.sort((a, b) => {
    const c = a.kaisyacd.localeCompare(b.kaisyacd);
    if (c !== 0) return c;
    return a.sehncd.localeCompare(b.sehncd);
  });
  return products;
}

const patchProductsSchema = z.object({
  rows: z.array(
    z.object({
      kaisyacd: z.string().min(1),
      sehncd: z.string().min(1),
      tumikae: z.string(),
    }),
  ),
});

const patchItemsSchema = z.object({
  items: z.array(
    z.object({
      id: z.string().min(1),
      checked: z.boolean().optional(),
      note: z.string().optional(),
    }),
  ),
});

ric3RepackRouter.get("/", requireAuth, async (_req, res) => {
  try {
    const state = await loadState();
    res.json({ data: { ...state, updatedAt: new Date().toISOString() } });
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: "Failed to load Ric3 repack state" });
  }
});

/** 積替品一覧を DB から再取得して ExR3Stg に保存（TumikaeHinn 相当） */
ric3RepackRouter.post("/actions/refresh-items", requireAuth, async (_req, res) => {
  try {
    const prev = await loadState();
    const items = await buildItemsFromDb();
    const next: Ric3RepackState = { ...prev, version: 1, items };
    await saveState(next);
    res.json({ data: { ...next, updatedAt: new Date().toISOString() } });
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: "Failed to refresh items" });
  }
});

/** 積替製品テーブルを DB から再取得（積替製品TB表示 相当） */
ric3RepackRouter.post("/actions/refresh-products", requireAuth, async (_req, res) => {
  try {
    const prev = await loadState();
    const products = await buildProductsFromDb();
    const next: Ric3RepackState = { ...prev, version: 1, products };
    await saveState(next);
    res.json({ data: { ...next, updatedAt: new Date().toISOString() } });
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: "Failed to refresh products" });
  }
});

/** 詰替要否の変更を JSON に反映し original を揃える（詰替品データ更新 の Stg 版） */
ric3RepackRouter.patch("/products", requireAuth, async (req, res) => {
  const parsed = patchProductsSchema.safeParse(req.body);
  if (!parsed.success) {
    res.status(400).json({ error: "Invalid body" });
    return;
  }
  try {
    const prev = await loadState();
    const byKey = new Map(parsed.data.rows.map((r) => [tumikaeKey(r.kaisyacd, r.sehncd), r.tumikae]));
    const products = prev.products.map((p) => {
      const nk = tumikaeKey(p.kaisyacd, p.sehncd);
      const nextTm = byKey.get(nk);
      if (nextTm === undefined) return p;
      return { ...p, tumikae: nextTm, originalTumikae: nextTm };
    });
    const next: Ric3RepackState = { ...prev, version: 1, products };
    await saveState(next);
    res.json({ data: { ...next, updatedAt: new Date().toISOString() } });
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: "Failed to update products" });
  }
});

/** 一覧のチェック・備考のみ更新 */
ric3RepackRouter.patch("/items", requireAuth, async (req, res) => {
  const parsed = patchItemsSchema.safeParse(req.body);
  if (!parsed.success) {
    res.status(400).json({ error: "Invalid body" });
    return;
  }
  try {
    const prev = await loadState();
    const patchById = new Map(parsed.data.items.map((x) => [x.id, x]));
    const items = prev.items.map((row) => {
      const p = patchById.get(row.id);
      if (!p) return row;
      return {
        ...row,
        checked: p.checked !== undefined ? p.checked : row.checked,
        note: p.note !== undefined ? p.note : row.note,
      };
    });
    const next: Ric3RepackState = { ...prev, version: 1, items };
    await saveState(next);
    res.json({ data: { ...next, updatedAt: new Date().toISOString() } });
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: "Failed to update items" });
  }
});

export { ric3RepackRouter };
