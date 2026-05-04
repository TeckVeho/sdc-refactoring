import { Router } from "express";
import { z } from "zod";

import { prisma } from "../../lib/prisma.js";
import { requireAuth } from "../../middleware/auth.js";

const companyArrivalRouter = Router();

/** 入荷日の表記揺れに対し、比較用に先頭8桁の数字列へ正規化 */
function normalizeYyyymmdd(s: string | null | undefined): string {
  if (s == null) return "";
  const digits = String(s).replace(/\D/g, "");
  return digits.length >= 8 ? digits.slice(0, 8) : digits;
}

function firstQueryString(v: unknown): string | undefined {
  if (v == null || v === "") return undefined;
  if (Array.isArray(v)) {
    return typeof v[0] === "string" ? v[0] : undefined;
  }
  return typeof v === "string" ? v : undefined;
}

/** Excel 仕様に合わせ、数値のみの会社コードを4桁ゼロ埋め */
function normalizeKaisyacd(raw: string): string {
  const t = raw.trim();
  if (/^\d+$/.test(t) && t.length > 0 && t.length < 4) {
    return t.padStart(4, "0");
  }
  return t;
}

function rowMatchesArrivalDate(
  r: { nyukabi: string | null | undefined; uno: string },
  dateYmd: string,
): boolean {
  if (normalizeYyyymmdd(r.nyukabi) === dateYmd) return true;
  const digits = String(r.uno).replace(/\D/g, "");
  return digits.length >= 8 && digits.slice(0, 8) === dateYmd;
}

const querySchema = z.object({
  date: z.string().regex(/^\d{8}$/),
  kaisyacd: z.string().min(1).optional(),
});

companyArrivalRouter.get("/", requireAuth, async (req, res) => {
  const parsed = querySchema.safeParse({
    date: firstQueryString(req.query.date),
    kaisyacd: firstQueryString(req.query.kaisyacd),
  });
  if (!parsed.success) {
    res.status(400).json({ error: "Invalid query (date=YYYYMMDD required)" });
    return;
  }

  const { date, kaisyacd: kaisyacdRaw } = parsed.data;
  const kaisyacdNorm = kaisyacdRaw ? normalizeKaisyacd(kaisyacdRaw) : null;

  try {
    const y = date.slice(0, 4);
    const yearNarrow = {
      OR: [{ nyukabi: { contains: y } }, { uno: { startsWith: y } }],
    };
    const baseWhere = kaisyacdNorm
      ? { kaisyacd: kaisyacdNorm, ...yearNarrow }
      : yearNarrow;

    const rawRows = await prisma.zaikor.findMany({
      where: baseWhere,
      select: {
        kaisyacd: true,
        sehncd: true,
        nyukasu: true,
        nyukabi: true,
        uno: true,
      },
    });

    const rowsOnDate = rawRows.filter((r) => rowMatchesArrivalDate(r, date));

    if (!kaisyacdNorm) {
      const byCompany = new Map<string, number>();
      for (const r of rowsOnDate) {
        const k = r.kaisyacd;
        const add = r.nyukasu ?? 0;
        byCompany.set(k, (byCompany.get(k) ?? 0) + add);
      }

      const codes = Array.from(byCompany.keys());
      const toku =
        codes.length === 0
          ? []
          : await prisma.tokumst.findMany({
              where: { kaisyacd: { in: codes } },
              select: { kaisyacd: true, coname: true, kairname: true },
            });
      const nameByCd = new Map(
        toku.map((t) => [
          t.kaisyacd,
          (t.coname?.trim() || t.kairname?.trim() || t.kaisyacd) as string,
        ]),
      );

      const list = Array.from(byCompany.entries())
        .map(([kaisyacd, nyukasu]) => ({
          kaisyacd,
          companyName: nameByCd.get(kaisyacd) ?? kaisyacd,
          nyukasu,
        }))
        .sort((a, b) => b.nyukasu - a.nyukasu || a.kaisyacd.localeCompare(b.kaisyacd));

      const grand = list.reduce((s, r) => s + r.nyukasu, 0);

      res.json({
        data: {
          mode: "byCompany" as const,
          date,
          rows: list,
          totals: { nyukasu: grand, companyCount: list.length },
        },
      });
      return;
    }

    const byProduct = new Map<string, number>();
    for (const r of rowsOnDate) {
      const p = r.sehncd;
      const add = r.nyukasu ?? 0;
      byProduct.set(p, (byProduct.get(p) ?? 0) + add);
    }

    const sehncds = Array.from(byProduct.keys());
    const seh =
      sehncds.length === 0
        ? []
        : await prisma.sehmst.findMany({
            where: { kaisyacd: kaisyacdNorm, sehncd: { in: sehncds } },
            select: { sehncd: true, seiname: true },
          });
    const nameBySehn = new Map(seh.map((s) => [s.sehncd, s.seiname.trim() || s.sehncd]));

    const list = Array.from(byProduct.entries())
      .map(([sehncd, nyukasu]) => ({
        kaisyacd: kaisyacdNorm,
        sehncd,
        productName: nameBySehn.get(sehncd) ?? sehncd,
        nyukasu,
      }))
      .sort((a, b) => b.nyukasu - a.nyukasu || a.sehncd.localeCompare(b.sehncd));

    let companyName = kaisyacdNorm;
    const oneToku = await prisma.tokumst.findUnique({
      where: { kaisyacd: kaisyacdNorm },
      select: { coname: true, kairname: true },
    });
    if (oneToku) {
      companyName =
        oneToku.coname?.trim() || oneToku.kairname?.trim() || kaisyacdNorm;
    }

    const grand = list.reduce((s, r) => s + r.nyukasu, 0);

    res.json({
      data: {
        mode: "byProduct" as const,
        date,
        kaisyacd: kaisyacdNorm,
        companyName,
        rows: list,
        totals: { nyukasu: grand, productCount: list.length },
      },
    });
  } catch (e) {
    res.status(500).json({ error: e instanceof Error ? e.message : "Internal error" });
  }
});

export { companyArrivalRouter };
