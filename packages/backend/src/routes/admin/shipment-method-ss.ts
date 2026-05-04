import { Router } from "express";

import { prisma } from "../../lib/prisma.js";
import { requireAuth } from "../../middleware/auth.js";
import { fetchShipmentMethodBoard } from "../../services/shipment-method-board.js";

const shipmentMethodSsRouter = Router();

/** GET /api/admin/shipment-method-ss — SS 検証向け読取（データソースは本番と同一） */
shipmentMethodSsRouter.get("/", requireAuth, async (_req, res) => {
  try {
    const board = await fetchShipmentMethodBoard(prisma);
    res.json({
      data: {
        ...board,
        view: "ss" as const,
        ref: "docs/★SS検証_Ex出荷方法報告書発行登録_仕様書.md",
      },
    });
  } catch (e) {
    res.status(500).json({ error: e instanceof Error ? e.message : "Server error" });
  }
});

export { shipmentMethodSsRouter };
