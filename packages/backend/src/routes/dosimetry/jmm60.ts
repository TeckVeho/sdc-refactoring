import { Router } from "express";
import { z } from "zod";

import { prisma } from "../../lib/prisma.js";
import { requireAuth } from "../../middleware/auth.js";
import { loadJmmPayload } from "../../services/jmm-sheet.js";

const jmm60Router = Router();

const querySchema = z.object({
  uno: z.string().min(1),
});

function firstQueryString(v: unknown): string | undefined {
  if (v == null || v === "") return undefined;
  if (Array.isArray(v)) {
    return typeof v[0] === "string" ? v[0] : undefined;
  }
  return typeof v === "string" ? v : undefined;
}

/** GET /api/dosimetry/jmm60?uno= — JMM60φ 記入用紙データ（受付マスタ読込） */
jmm60Router.get("/", requireAuth, async (req, res) => {
  const parsed = querySchema.safeParse({ uno: firstQueryString(req.query.uno) });
  if (!parsed.success) {
    res.status(400).json({ error: "クエリ uno が必要です" });
    return;
  }

  try {
    const data = await loadJmmPayload(prisma, parsed.data.uno, "60");
    if (!data) {
      res.status(404).json({ error: "受付が見つかりません" });
      return;
    }
    res.json({ data });
  } catch (e) {
    res.status(500).json({ error: e instanceof Error ? e.message : "Server error" });
  }
});

export { jmm60Router };
