import { Router } from "express";
import { z } from "zod";

import { prisma } from "../../lib/prisma.js";
import { requireAuth } from "../../middleware/auth.js";

const customerStockRouter = Router();

/** 入荷日・出荷日の表記揺れに対し、比較用に先頭8桁の数字列へ正規化 */
function normalizeYyyymmdd(s: string | null | undefined): string {
  if (s == null) return "";
  const digits = String(s).replace(/\D/g, "");
  return digits.length >= 8 ? digits.slice(0, 8) : digits;
}

function normalizeKaisyacd(raw: string): string {
  const t = raw.trim();
  if (/^\d+$/.test(t) && t.length > 0 && t.length < 4) {
    return t.padStart(4, "0");
  }
  return t;
}

function yearPrefixesFromTo(fromYear: number, toYear: number): string[] {
  const out: string[] = [];
  for (let y = fromYear; y <= toYear; y++) {
    out.push(String(y));
  }
  return out;
}

function firstQueryString(v: unknown): string | undefined {
  if (v == null || v === "") return undefined;
  if (Array.isArray(v)) {
    return typeof v[0] === "string" ? v[0] : undefined;
  }
  return typeof v === "string" ? v : undefined;
}

function sortKeyUno(uno: string): number {
  const n = Number.parseFloat(uno.replace(/\D/g, "") || "0");
  return Number.isFinite(n) ? n : 0;
}

const querySchema = z.object({
  kaisyacd: z.string().min(1),
  asOf: z.string().regex(/^\d{8}$/),
});

/**
 * Ex顧客在庫報告相当: 指定会社・在庫基準日時点の zaiko 一覧。
 * 備考: Prisma スキーマが最小構成のため、Excel 原票の zaikor / syukar 紐付け（受付番号単位の加算）は未実装。
 */
customerStockRouter.get("/", requireAuth, async (req, res) => {
  const parsed = querySchema.safeParse({
    kaisyacd: firstQueryString(req.query.kaisyacd),
    asOf: firstQueryString(req.query.asOf),
  });
  if (!parsed.success) {
    res.status(400).json({ error: "Invalid query: kaisyacd and asOf (YYYYMMDD) are required" });
    return;
  }

  const kaisyacd = normalizeKaisyacd(parsed.data.kaisyacd);
  const asOf = parsed.data.asOf;
  const asYear = parseInt(asOf.slice(0, 4), 10);
  const nyYears = yearPrefixesFromTo(2000, asYear);

  try {
    const [toku, zaikoRaw] = await Promise.all([
      prisma.tokumst.findUnique({
        where: { kaisyacd },
        select: { coname: true, kairname: true },
      }),
      prisma.zaiko.findMany({
        where: {
          kaisyacd,
          OR: nyYears.map((y) => ({ nyukabi: { startsWith: y } })),
        },
        select: {
          uno: true,
          sehncd: true,
          kainame: true,
          nyukabi: true,
          nyukasu: true,
          pass: true,
          syouso: true,
        },
      }),
    ]);

    const companyName =
      toku?.coname?.trim() || toku?.kairname?.trim() || null;

    const zaikoRows = zaikoRaw.filter((r) => {
      const d = normalizeYyyymmdd(r.nyukabi);
      return d.length === 8 && d <= asOf;
    });

    const sehKeys = zaikoRows.map((r) => ({ kaisyacd, sehncd: r.sehncd }));
    const sehUnique = Array.from(
      new Map(sehKeys.map((k) => [`${k.kaisyacd}\t${k.sehncd}`, k])).values(),
    );

    const productNames = new Map<string, string>();
    if (sehUnique.length > 0) {
      const seh = await prisma.sehmst.findMany({
        where: { OR: sehUnique },
        select: { sehncd: true, seiname: true },
      });
      for (const s of seh) {
        productNames.set(s.sehncd, s.seiname.trim());
      }
    }

    const syYears = yearPrefixesFromTo(asYear, Math.max(asYear, new Date().getFullYear()));
    const syRaw = await prisma.syukar.findMany({
      where: {
        kaisyacd,
        OR: syYears.map((y) => ({ syudate: { startsWith: y } })),
      },
      select: { syudate: true, syukasu: true },
    });

    const shipmentsOnOrAfterAsOf = syRaw.filter((r) => {
      const d = normalizeYyyymmdd(r.syudate);
      return d.length === 8 && d >= asOf;
    });
    const shipmentAfterAsOfSum = shipmentsOnOrAfterAsOf.reduce(
      (s, r) => s + (r.syukasu ?? 0),
      0,
    );

    const sorted = [...zaikoRows].sort((a, b) => sortKeyUno(a.uno) - sortKeyUno(b.uno));

    const rows = sorted.map((r, i) => {
      const nyukasu = r.nyukasu ?? 0;
      return {
        no: i + 1,
        uno: r.uno,
        nyukabi: r.nyukabi,
        sehncd: r.sehncd,
        productName: productNames.get(r.sehncd) ?? null,
        kainame: r.kainame?.trim() ?? null,
        pass: r.pass?.trim() ?? null,
        syouso: r.syouso?.trim() ?? null,
        nyukasu: r.nyukasu,
        /** スキーマ拡張までの仮値: 入荷数を預り在庫の目安として返す */
        computedStock: nyukasu,
      };
    });

    const totalNyukasu = rows.reduce((s, r) => s + (r.nyukasu ?? 0), 0);
    const totalComputedStock = rows.reduce((s, r) => s + (r.computedStock ?? 0), 0);

    res.json({
      data: {
        asOf: new Date().toISOString(),
        basisDate: asOf,
        kaisyacd,
        companyName,
        note:
          "現行DBスキーマでは syukar.uno / zaikor.syukabi 等が無いため、Excel原票の受付番号別出荷加算・履歴合算は未反映です。",
        rows,
        totals: {
          rowCount: rows.length,
          nyukasu: totalNyukasu,
          computedStock: totalComputedStock,
        },
        shipmentAfterAsOfSum,
      },
    });
  } catch (e) {
    res.status(500).json({ error: e instanceof Error ? e.message : "Internal error" });
  }
});

export { customerStockRouter };
