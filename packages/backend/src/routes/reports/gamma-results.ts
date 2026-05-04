import { Router } from "express";
import { z } from "zod";

import { prisma } from "../../lib/prisma.js";
import { requireAuth } from "../../middleware/auth.js";
import {
  ctimerInKkmdRange,
  kkmdRangeFromMonth,
  parseSiteisnKg,
  unoRangeFromYearMonth,
} from "../../services/gamma-results-aggregate.js";

const gammaResultsRouter = Router();

const aggregateBodySchema = z.object({
  year: z.number().int().min(1980).max(2100),
  month: z.number().int().min(1).max(12),
  unoMin: z.string().min(8).optional(),
  unoMax: z.string().min(8).optional(),
});

async function sumNyukasu(
  unoMin: string,
  unoMax: string,
  syouso: string,
): Promise<number> {
  const [z1, z2] = await Promise.all([
    prisma.zaiko.aggregate({
      where: {
        uno: { gte: unoMin, lte: unoMax },
        syouso,
      },
      _sum: { nyukasu: true },
    }),
    prisma.zaikor.aggregate({
      where: {
        uno: { gte: unoMin, lte: unoMax },
        syouso,
      },
      _sum: { nyukasu: true },
    }),
  ]);
  return (z1._sum.nyukasu ?? 0) + (z2._sum.nyukasu ?? 0);
}

/** POST /api/reports/gamma-results/aggregate — 月次集計（MVP） */
gammaResultsRouter.post("/aggregate", requireAuth, async (req, res) => {
  const parsed = aggregateBodySchema.safeParse(req.body);
  if (!parsed.success) {
    res.status(400).json({ error: "Invalid body" });
    return;
  }

  const { year, month } = parsed.data;
  const derived = unoRangeFromYearMonth(year, month);
  const unoMin = parsed.data.unoMin ?? derived.unoMin;
  const unoMax = parsed.data.unoMax ?? derived.unoMax;
  const { kkmd1, kkmd2 } = kkmdRangeFromMonth(month);

  try {
    const [ric1iri, ric2iri, ric3iri, sk1rows, n2, n3] = await Promise.all([
      sumNyukasu(unoMin, unoMax, "1"),
      sumNyukasu(unoMin, unoMax, "2"),
      sumNyukasu(unoMin, unoMax, "3"),
      prisma.syouk1.findMany({
        where: { uno: { gte: unoMin, lte: unoMax } },
        select: { uno: true, siteisn: true, syosuu: true, ctimer: true },
      }),
      prisma.syouk2.count({
        where: { uno: { gte: unoMin, lte: unoMax } },
      }),
      prisma.syoukj3.count({
        where: { uno: { gte: unoMin, lte: unoMax } },
      }),
    ]);

    let ric1Simple = 0;
    for (const r of sk1rows) {
      if (!ctimerInKkmdRange(r.ctimer, kkmd1, kkmd2)) continue;
      const sn = parseSiteisnKg(r.siteisn);
      if (sn == null) continue;
      const qty = r.syosuu ?? 1;
      ric1Simple += sn * qty;
    }

    res.json({
      data: {
        params: {
          year,
          month,
          unoMin,
          unoMax,
          kkmd1,
          kkmd2,
        },
        goukei: {
          ric1: {
            inboundQty: ric1iri,
            simpleProcessingApprox: Math.round(ric1Simple * 1000) / 1000,
          },
          ric2: {
            inboundQty: ric2iri,
            simpleProcessingApprox: n2,
          },
          ric3: {
            inboundQty: ric3iri,
            simpleProcessingApprox: n3,
          },
        },
        notes: [
          "2・3号機の処理量はスキーマが最小のため件数を返しています。",
          "混載考慮・標準ケース換算は将来追加してください。",
        ],
      },
    });
  } catch (e) {
    res.status(500).json({ error: e instanceof Error ? e.message : "Server error" });
  }
});

export { gammaResultsRouter };
