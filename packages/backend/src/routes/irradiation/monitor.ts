import { Router } from "express";

import { requireAuth } from "../../middleware/auth.js";
import { getMachine1MonitorData } from "../../services/machine1-monitor.js";

const monitorRouter = Router();

/** GET /api/irradiation/machine1/monitor */
monitorRouter.get("/", requireAuth, async (_req, res, next) => {
  try {
    const data = await getMachine1MonitorData();
    res.json({ data });
  } catch (e) {
    next(e);
  }
});

export { monitorRouter };
