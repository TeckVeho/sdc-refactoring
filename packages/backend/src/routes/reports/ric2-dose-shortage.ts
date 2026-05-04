import { Router } from "express";
import { z } from "zod";

import { prisma } from "../../lib/prisma.js";
import { requireAuth } from "../../middleware/auth.js";

const ric2DoseShortageRouter = Router();

const dataQuerySchema = z.object({
  recentTake: z.coerce.number().int().min(1).max(50).optional().default(20),
  kansokuTake: z.coerce.number().int().min(1).max(200).optional().default(80),
});

/**
 * RIC2 線量不足報告書 — サマリー用データ（MVP）
 * 本番 Oracle の SYOUJ2.senkno / SYOUK2 多列は Prisma 最小スキーマに未含むため、
 * 参照用に Syouk2 / Zaiko / Kansoku / Syouj2 の件数・直近行を返す。
 */
ric2DoseShortageRouter.get("/data", requireAuth, async (req, res, next) => {
  try {
    const q = dataQuerySchema.parse(req.query);

    const [
      syouk2Count,
      zaikoCount,
      kansokuCount,
      syouj2Count,
      recentSyouk2,
      kansokuRows,
    ] = await Promise.all([
      prisma.syouk2.count(),
      prisma.zaiko.count(),
      prisma.kansoku.count(),
      prisma.syouj2.count(),
      prisma.syouk2.findMany({
        orderBy: { uno: "desc" },
        take: q.recentTake,
        select: {
          uno: true,
          syono: true,
          kainame: true,
          syokind: true,
          sdate: true,
          edate: true,
          bikou: true,
        },
      }),
      prisma.kansoku.findMany({
        orderBy: { kid: "asc" },
        take: q.kansokuTake,
        select: { kid: true, sokutei: true, bikou: true },
      }),
    ]);

    const unos = recentSyouk2.map((r) => r.uno);
    const zaikoByUno =
      unos.length === 0
        ? new Map<string, { kaisyacd: string; sehncd: string; kainame: string | null }>()
        : new Map(
            (
              await prisma.zaiko.findMany({
                where: { uno: { in: unos } },
                select: { uno: true, kaisyacd: true, sehncd: true, kainame: true },
              })
            ).map((z) => [
              z.uno,
              { kaisyacd: z.kaisyacd, sehncd: z.sehncd, kainame: z.kainame },
            ]),
          );

    const recentPlansWithZaiko = recentSyouk2.map((s) => {
      const z = zaikoByUno.get(s.uno);
      return {
        ...s,
        zaiko: z ?? null,
      };
    });

    res.json({
      data: {
        summary: {
          syouk2Count,
          zaikoCount,
          kansokuCount,
          syouj2Count,
        },
        recentPlansWithZaiko,
        kansokuRows,
        schemaNote:
          "MVP: senkno・照射計画8列などはDBスキーマ拡張後に追加可能。",
      },
    });
  } catch (e) {
    next(e);
  }
});

export { ric2DoseShortageRouter };
