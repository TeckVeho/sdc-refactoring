import { Router } from "express";
import { z } from "zod";

import { prisma } from "../../lib/prisma.js";
import { requireAuth } from "../../middleware/auth.js";

const companyCodeConvertRouter = Router();

const upsertSchema = z.object({
  gammacd: z.string().min(1),
  ebcd: z.string().min(1),
  memo: z.string().optional(),
});

function firstQueryString(v: unknown): string | undefined {
  if (v == null || v === "") return undefined;
  if (Array.isArray(v)) {
    return typeof v[0] === "string" ? v[0] : undefined;
  }
  return typeof v === "string" ? v : undefined;
}

/** GET /api/admin/company-code-convert?q= */
companyCodeConvertRouter.get("/", requireAuth, async (req, res) => {
  const q = firstQueryString(req.query.q)?.trim();

  try {
    const rows = await prisma.kcdcnvmst.findMany({
      take: 500,
      where: q
        ? {
            OR: [
              { gammacd: { contains: q } },
              { ebcd: { contains: q } },
              { memo: { contains: q } },
            ],
          }
        : undefined,
      orderBy: [{ gammacd: "asc" }, { ebcd: "asc" }],
    });

    res.json({ data: { rows } });
  } catch (e) {
    res.status(500).json({ error: e instanceof Error ? e.message : "Server error" });
  }
});

/** POST /api/admin/company-code-convert — UPSERT */
companyCodeConvertRouter.post("/", requireAuth, async (req, res) => {
  const parsed = upsertSchema.safeParse(req.body);
  if (!parsed.success) {
    res.status(400).json({ error: "Invalid body" });
    return;
  }

  const { gammacd, ebcd, memo } = parsed.data;

  try {
    await prisma.kcdcnvmst.upsert({
      where: {
        gammacd_ebcd: { gammacd, ebcd },
      },
      create: { gammacd, ebcd, memo: memo ?? null },
      update: { memo: memo ?? null },
    });
    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: e instanceof Error ? e.message : "Server error" });
  }
});

/** DELETE /api/admin/company-code-convert?gammacd=&ebcd= */
companyCodeConvertRouter.delete("/", requireAuth, async (req, res) => {
  const gammacd = firstQueryString(req.query.gammacd)?.trim();
  const ebcd = firstQueryString(req.query.ebcd)?.trim();
  if (!gammacd || !ebcd) {
    res.status(400).json({ error: "gammacd と ebcd が必要です" });
    return;
  }

  try {
    await prisma.kcdcnvmst.delete({
      where: {
        gammacd_ebcd: { gammacd, ebcd },
      },
    });
    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: e instanceof Error ? e.message : "Server error" });
  }
});

export { companyCodeConvertRouter };
