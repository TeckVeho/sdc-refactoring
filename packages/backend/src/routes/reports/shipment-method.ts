import { Router } from "express";
import { z } from "zod";

import { prisma } from "../../lib/prisma.js";
import { requireAuth } from "../../middleware/auth.js";
import { fetchShipmentMethodBoard } from "../../services/shipment-method-board.js";
import { shipmentRowkey, syouhoMaxNumeric } from "../../services/shipment-method-registry.js";

const shipmentMethodRouter = Router();

const putBodySchema = z.object({
  rows: z
    .array(
      z.object({
        kaisyacd: z.string().min(1),
        hikitori: z.string(),
        housyube: z.string(),
      }),
    )
    .max(500),
});

/** GET /api/reports/shipment-method — 業者一覧（TOKUMST + ExSeihinj + SEHMST 集約） */
shipmentMethodRouter.get("/", requireAuth, async (_req, res) => {
  try {
    const board = await fetchShipmentMethodBoard(prisma);
    res.json({ data: board });
  } catch (e) {
    res.status(500).json({ error: e instanceof Error ? e.message : "Server error" });
  }
});

/** PUT /api/reports/shipment-method — 差分を ExSeihinj に UPSERT */
shipmentMethodRouter.put("/", requireAuth, async (req, res) => {
  const parsed = putBodySchema.safeParse(req.body);
  if (!parsed.success) {
    res.status(400).json({ error: "Invalid body" });
    return;
  }

  try {
    const flags = await prisma.sehmst.groupBy({
      by: ["kaisyacd"],
      _max: { syouho: true },
    });
    const flagMap = new Map<string, number>();
    for (const r of flags) {
      flagMap.set(r.kaisyacd, syouhoMaxNumeric(r._max.syouho ?? null));
    }

    let updated = 0;
    for (const row of parsed.data.rows) {
      const cd = row.kaisyacd.trim();
      const flag = flagMap.get(cd) ?? 0;
      if (flag <= 0 && row.housyube.trim() !== "不要") {
        res.status(400).json({
          error: `製品仕様台帳の報告書発行が不要です（${cd}）。報告書発行種別は「不要」のみ登録できます。`,
        });
        return;
      }

      const rk = shipmentRowkey(cd);
      await prisma.exSeihinj.upsert({
        where: { rowkey: rk },
        create: {
          rowkey: rk,
          json: { hikitori: row.hikitori, housyube: row.housyube },
        },
        update: {
          json: { hikitori: row.hikitori, housyube: row.housyube },
        },
      });
      updated += 1;
    }

    res.json({ ok: true, updated });
  } catch (e) {
    res.status(500).json({ error: e instanceof Error ? e.message : "Server error" });
  }
});

export { shipmentMethodRouter };
