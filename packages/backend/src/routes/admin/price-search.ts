import { Router } from "express";
import { z } from "zod";

import { prisma } from "../../lib/prisma.js";
import { requireAuth } from "../../middleware/auth.js";

const priceSearchRouter = Router();

export function priceRowkey(kaisyacd: string, sehncd: string): string {
  return `price:${kaisyacd.trim()}:${sehncd.trim()}`;
}

function parsePriceJson(json: unknown): Record<string, string> {
  if (!json || typeof json !== "object") return {};
  const out: Record<string, string> = {};
  for (const [k, v] of Object.entries(json as Record<string, unknown>)) {
    if (v == null) out[k] = "";
    else if (typeof v === "string" || typeof v === "number") out[k] = String(v);
    else out[k] = JSON.stringify(v);
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

const querySchema = z
  .object({
    uno: z.string().min(1).optional(),
    kaisyacd: z.string().min(1).optional(),
    sehncd: z.string().min(1).optional(),
  })
  .refine((q) => Boolean(q.uno || (q.kaisyacd && q.sehncd)), {
    message: "uno、または kaisyacd+sehncd が必要です",
  });

/** GET /api/admin/price-search */
priceSearchRouter.get("/", requireAuth, async (req, res) => {
  const parsed = querySchema.safeParse({
    uno: firstQueryString(req.query.uno)?.trim(),
    kaisyacd: firstQueryString(req.query.kaisyacd)?.trim(),
    sehncd: firstQueryString(req.query.sehncd)?.trim(),
  });
  if (!parsed.success) {
    res.status(400).json({ error: parsed.error.flatten().formErrors[0] ?? "Invalid query" });
    return;
  }

  try {
    let kaisyacd = parsed.data.kaisyacd;
    let sehncd = parsed.data.sehncd;
    let uno = parsed.data.uno;

    if (uno) {
      const zk = await prisma.zaiko.findUnique({
        where: { uno },
        select: { kaisyacd: true, sehncd: true },
      });
      if (!zk) {
        res.status(404).json({ error: "受付（zaiko）が見つかりません" });
        return;
      }
      kaisyacd = zk.kaisyacd;
      sehncd = zk.sehncd;
    }

    if (!kaisyacd || !sehncd) {
      res.status(400).json({ error: "会社コード・製品コードが必要です" });
      return;
    }

    const [zkRow, seh, priceRow, toku] = await Promise.all([
      uno
        ? prisma.zaiko.findUnique({ where: { uno } })
        : prisma.zaiko.findFirst({
            where: { kaisyacd, sehncd },
            orderBy: { uno: "desc" },
          }),
      prisma.sehmst.findFirst({
        where: { kaisyacd, sehncd },
      }),
      prisma.exSeihinz.findUnique({
        where: { rowkey: priceRowkey(kaisyacd, sehncd) },
      }),
      prisma.tokumst.findUnique({ where: { kaisyacd } }),
    ]);

    const priceFields = parsePriceJson(priceRow?.json ?? null);

    res.json({
      data: {
        uno: zkRow?.uno ?? uno ?? null,
        kaisyacd,
        sehncd,
        companyName: toku?.coname ?? null,
        companyShort: toku?.kairname ?? null,
        productName: seh?.seiname ?? null,
        zaiko: zkRow,
        priceRowkey: priceRowkey(kaisyacd, sehncd),
        price: {
          tanka: priceFields.tanka ?? "",
          tanni: priceFields.tanni ?? "",
          tourokubi: priceFields.tourokubi ?? "",
          dose: priceFields.dose ?? "",
          souti: priceFields.souti ?? "",
          pass: priceFields.pass ?? "",
          folder: priceFields.folder ?? "",
          file: priceFields.file ?? "",
          ...priceFields,
        },
        note: "単価マスタは ExSeihinz の JSON（price:会社:製品）で保持する MVP です。",
      },
    });
  } catch (e) {
    res.status(500).json({ error: e instanceof Error ? e.message : "Server error" });
  }
});

const putPriceSchema = z.object({
  kaisyacd: z.string().min(1),
  sehncd: z.string().min(1),
  fields: z.record(z.string(), z.string()),
});

/** PUT /api/admin/price-search — ExSeihinz UPSERT */
priceSearchRouter.put("/", requireAuth, async (req, res) => {
  const parsed = putPriceSchema.safeParse(req.body);
  if (!parsed.success) {
    res.status(400).json({ error: "Invalid body" });
    return;
  }

  const { kaisyacd, sehncd, fields } = parsed.data;
  const rk = priceRowkey(kaisyacd, sehncd);

  try {
    await prisma.exSeihinz.upsert({
      where: { rowkey: rk },
      create: { rowkey: rk, json: fields },
      update: { json: fields },
    });
    res.json({ ok: true, rowkey: rk });
  } catch (e) {
    res.status(500).json({ error: e instanceof Error ? e.message : "Server error" });
  }
});

export { priceSearchRouter };
