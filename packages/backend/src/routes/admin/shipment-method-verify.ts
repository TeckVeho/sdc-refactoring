import { Router } from "express";

import { prisma } from "../../lib/prisma.js";
import { requireAuth } from "../../middleware/auth.js";
import { fetchShipmentMethodBoard } from "../../services/shipment-method-board.js";

const shipmentMethodVerifyRouter = Router();

/** GET /api/admin/shipment-method-verify — 本番登録画面と同一データの読取専用ビュー */
shipmentMethodVerifyRouter.get("/", requireAuth, async (_req, res) => {
  try {
    const board = await fetchShipmentMethodBoard(prisma);
    res.json({
      data: {
        ...board,
        view: "verify" as const,
        ref: "docs/★検証_Ex出荷方法報告書発行登録_仕様書.md",
      },
    });
  } catch (e) {
    res.status(500).json({ error: e instanceof Error ? e.message : "Server error" });
  }
});

export { shipmentMethodVerifyRouter };
