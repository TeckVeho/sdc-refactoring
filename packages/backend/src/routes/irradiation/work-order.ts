import { Router } from "express";
import { z } from "zod";

import { prisma } from "../../lib/prisma.js";
import { requireAuth } from "../../middleware/auth.js";

const workOrderRouter = Router();

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

function parseNum(raw: string | null | undefined): number | null {
  if (raw == null || raw === "") return null;
  const n = Number(String(raw).replace(/,/g, "").trim());
  return Number.isFinite(n) ? n : null;
}

/** GET /api/irradiation/machine1/work-order?uno= — 1号機作業指図書用データ */
workOrderRouter.get("/", requireAuth, async (req, res) => {
  const parsed = querySchema.safeParse({ uno: firstQueryString(req.query.uno) });
  if (!parsed.success) {
    res.status(400).json({ error: "クエリ uno が必要です" });
    return;
  }

  const uno = parsed.data.uno.trim();

  try {
    const [zk, sk1, employees, ratetbl] = await Promise.all([
      prisma.zaiko.findUnique({ where: { uno } }),
      prisma.syouk1.findUnique({ where: { uno } }),
      prisma.shainmst.findMany({
        take: 80,
        orderBy: { shano: "asc" },
        select: { shano: true, shaname: true },
      }),
      prisma.ratetbl.findMany({
        take: 40,
        orderBy: { ratekey: "asc" },
        select: { ratekey: true, rateval: true },
      }),
    ]);

    if (!zk && !sk1) {
      res.status(404).json({ error: "在庫・照射データが見つかりません" });
      return;
    }

    const senrituNum =
      sk1?.senritu != null ? parseNum(String(sk1.senritu)) : null;

    res.json({
      data: {
        meta: {
          formTitle: "１号機作業指図書（製品仕様書なし）自動印字用",
          specRef: "docs/Ex１号機作業指図書_仕様書.md",
        },
        uno,
        zaiko: zk,
        syouk1: sk1,
        derived: {
          kainame: zk?.kainame ?? sk1?.kainame ?? null,
          nyukaBi: zk?.nyukabi ?? null,
          siteiSn: parseNum(sk1?.siteisn ?? null),
          syosuu: sk1?.syosuu ?? null,
          senritu: senrituNum,
          syono: sk1?.syono ?? null,
          syokind: sk1?.syokind ?? null,
          bikou: sk1?.bikou ?? null,
          sdate: sk1?.sdate ?? null,
          edate: sk1?.edate ?? null,
        },
        employees,
        ratetbl: ratetbl.map((r) => ({
          ratekey: r.ratekey,
          rateval: r.rateval != null ? String(r.rateval) : null,
        })),
      },
    });
  } catch (e) {
    res.status(500).json({ error: e instanceof Error ? e.message : "Server error" });
  }
});

export { workOrderRouter };
