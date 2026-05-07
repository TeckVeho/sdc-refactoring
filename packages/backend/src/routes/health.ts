import { Router } from "express";

import { isRelaxedGuestAuth } from "../lib/dev-guest-auth.js";
import { isSampleDataMode } from "../lib/sample-data-mode.js";

const healthRouter = Router();

healthRouter.get("/", (_req, res) => {
  const sampleDataMode = isSampleDataMode();
  const relaxedGuestAuth = isRelaxedGuestAuth();
  res.json({
    data: {
      ok: true as const,
      sampleDataMode,
      relaxedGuestAuth,
      /** NEXT_PUBLIC と揃える用: API がゲストでも通るか */
      guestApiAccess: sampleDataMode || relaxedGuestAuth,
    },
  });
});

export { healthRouter };
