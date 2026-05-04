import { Router } from "express";
import { z } from "zod";

import { prisma } from "../../lib/prisma.js";
import { requireAuth } from "../../middleware/auth.js";
import { buildRic3DoseShortagePayload } from "../../services/ric3-dose-shortage.js";

const ric3DoseShortageRouter = Router();

const querySchema = z.object({
  senkNo: z.string().min(1).optional(),
});

function firstQueryString(v: unknown): string | undefined {
  if (v == null || v === "") return undefined;
  if (Array.isArray(v)) {
    return typeof v[0] === "string" ? v[0] : undefined;
  }
  return typeof v === "string" ? v : undefined;
}

/** GET /api/reports/ric3-dose-shortage — RIC3 線量不足報告書（一般様式） */
ric3DoseShortageRouter.get("/", requireAuth, async (req, res) => {
  const parsed = querySchema.safeParse({
    senkNo: firstQueryString(req.query.senkNo),
  });
  if (!parsed.success) {
    res.status(400).json({ error: "Invalid query" });
    return;
  }

  try {
    const data = await buildRic3DoseShortagePayload(prisma, "standard", parsed.data.senkNo);
    res.json({ data });
  } catch (e) {
    res.status(500).json({ error: e instanceof Error ? e.message : "Server error" });
  }
});

export { ric3DoseShortageRouter };
