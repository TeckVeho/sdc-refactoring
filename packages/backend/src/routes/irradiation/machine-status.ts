import { Router } from "express";

import { requireAuth } from "../../middleware/auth.js";
import { getMachineStatusSnapshot } from "../../services/machine-status-snapshot.js";

const machineStatusRouter = Router();

/** GET /api/irradiation/machine-status — 1〜3号機スナップショット＋直近24hチャート点 */
machineStatusRouter.get("/", requireAuth, async (_req, res, next) => {
  try {
    const data = await getMachineStatusSnapshot();
    res.json({ data });
  } catch (e) {
    next(e);
  }
});

export { machineStatusRouter };
