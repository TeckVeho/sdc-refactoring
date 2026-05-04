import { Router } from "express";
import { z } from "zod";
import { Prisma } from "@prisma/client";

import { prisma } from "../../lib/prisma.js";
import { requireAuth } from "../../middleware/auth.js";

const irradiationResultsRouter = Router();

const querySchema = z.object({
  source: z.enum(["all", "syouk1", "syouk2", "syoukj3"]).default("all"),
  uno: z.string().optional(),
  kainame: z.string().optional(),
  page: z.coerce.number().int().min(1).default(1),
  pageSize: z.coerce.number().int().min(1).max(200).default(50),
});

function firstQueryString(v: unknown): string | undefined {
  if (v == null || v === "") return undefined;
  if (Array.isArray(v)) {
    return typeof v[0] === "string" ? v[0] : undefined;
  }
  return typeof v === "string" ? v : undefined;
}

function whereSyouk1(uno: string | undefined, kainame: string | undefined): Prisma.Syouk1WhereInput {
  const w: Prisma.Syouk1WhereInput = {};
  const u = uno?.trim();
  const k = kainame?.trim();
  if (u) w.uno = { contains: u };
  if (k) w.kainame = { contains: k };
  return w;
}

function whereSyouk2(
  uno: string | undefined,
  kainame: string | undefined,
): Prisma.Syouk2WhereInput {
  const w: Prisma.Syouk2WhereInput = {};
  const u = uno?.trim();
  const k = kainame?.trim();
  if (u) w.uno = { contains: u };
  if (k) w.kainame = { contains: k };
  return w;
}

function whereSyoukj3(
  uno: string | undefined,
  kainame: string | undefined,
): Prisma.Syoukj3WhereInput {
  const w: Prisma.Syoukj3WhereInput = {};
  const u = uno?.trim();
  const k = kainame?.trim();
  if (u) w.uno = { contains: u };
  if (k) w.kainame = { contains: k };
  return w;
}

export type IrradiationResultRow = {
  source: "syouk1" | "syouk2" | "syoukj3";
  uno: string;
  syono: string | null;
  kainame: string | null;
  siteisn: string | null;
  syosuu: number | null;
  syoichi: string | null;
  syotime: string | null;
  hansuu: number | null;
  stimer: string | null;
  ktimer: string | null;
  senritu: string | null;
  syostat: string | null;
  ctimer: string | null;
  zhansuu: number | null;
  htimer: string | null;
  slotno: string | null;
  sdate: string | null;
  edate: string | null;
  updfflg: string | null;
  syokind: string | null;
  bikou: string | null;
};

function mapSyouk1(r: {
  uno: string;
  syono: string | null;
  kainame: string | null;
  siteisn: string | null;
  syosuu: number | null;
  syoichi: string | null;
  syotime: string | null;
  hansuu: number | null;
  stimer: string | null;
  ktimer: string | null;
  senritu: { toString(): string } | null;
  syostat: string | null;
  ctimer: string | null;
  zhansuu: number | null;
  htimer: string | null;
  slotno: string | null;
  sdate: string | null;
  edate: string | null;
  updfflg: string | null;
  syokind: string | null;
  bikou: string | null;
}): IrradiationResultRow {
  return {
    source: "syouk1",
    uno: r.uno,
    syono: r.syono ?? null,
    kainame: r.kainame ?? null,
    siteisn: r.siteisn ?? null,
    syosuu: r.syosuu ?? null,
    syoichi: r.syoichi ?? null,
    syotime: r.syotime ?? null,
    hansuu: r.hansuu ?? null,
    stimer: r.stimer ?? null,
    ktimer: r.ktimer ?? null,
    senritu: r.senritu != null ? String(r.senritu) : null,
    syostat: r.syostat ?? null,
    ctimer: r.ctimer ?? null,
    zhansuu: r.zhansuu ?? null,
    htimer: r.htimer ?? null,
    slotno: r.slotno ?? null,
    sdate: r.sdate ?? null,
    edate: r.edate ?? null,
    updfflg: r.updfflg ?? null,
    syokind: r.syokind ?? null,
    bikou: r.bikou ?? null,
  };
}

function mapSyouk2(
  r: {
    uno: string;
    syono: string | null;
    kainame: string | null;
    syokind: string | null;
    sdate: string | null;
    edate: string | null;
    bikou: string | null;
  },
): IrradiationResultRow {
  return {
    source: "syouk2",
    uno: r.uno,
    syono: r.syono ?? null,
    kainame: r.kainame ?? null,
    siteisn: null,
    syosuu: null,
    syoichi: null,
    syotime: null,
    hansuu: null,
    stimer: null,
    ktimer: null,
    senritu: null,
    syostat: null,
    ctimer: null,
    zhansuu: null,
    htimer: null,
    slotno: null,
    sdate: r.sdate ?? null,
    edate: r.edate ?? null,
    updfflg: null,
    syokind: r.syokind ?? null,
    bikou: r.bikou ?? null,
  };
}

function mapSyoukj3(
  r: {
    uno: string;
    syono: string | null;
    kainame: string | null;
    syokind: string | null;
    sdate: string | null;
    edate: string | null;
    bikou: string | null;
  },
): IrradiationResultRow {
  return { ...mapSyouk2(r), source: "syoukj3" };
}

type RawUnionRow = {
  src: string;
  uno: string;
  syono: string | null;
  kainame: string | null;
  siteisn: string | null;
  syosuu: number | null;
  syoichi: string | null;
  syotime: string | null;
  hansuu: number | null;
  stimer: string | null;
  ktimer: string | null;
  senritu: string | null;
  syostat: string | null;
  ctimer: string | null;
  zhansuu: number | null;
  htimer: string | null;
  slotno: string | null;
  sdate: string | null;
  edate: string | null;
  updfflg: string | null;
  syokind: string | null;
  bikou: string | null;
};

function mapRaw(r: RawUnionRow): IrradiationResultRow {
  const s = r.src;
  if (s !== "syouk1" && s !== "syouk2" && s !== "syoukj3") {
    throw new Error("invalid source");
  }
  return {
    source: s,
    uno: r.uno,
    syono: r.syono,
    kainame: r.kainame,
    siteisn: r.siteisn,
    syosuu: r.syosuu,
    syoichi: r.syoichi,
    syotime: r.syotime,
    hansuu: r.hansuu,
    stimer: r.stimer,
    ktimer: r.ktimer,
    senritu: r.senritu,
    syostat: r.syostat,
    ctimer: r.ctimer,
    zhansuu: r.zhansuu,
    htimer: r.htimer,
    slotno: r.slotno,
    sdate: r.sdate,
    edate: r.edate,
    updfflg: r.updfflg,
    syokind: r.syokind,
    bikou: r.bikou,
  };
}

function unoKainameSql(
  uno: string | undefined,
  kainame: string | undefined,
): Prisma.Sql {
  const u = uno?.trim();
  const k = kainame?.trim();
  if (u && k) {
    return Prisma.sql`AND uno LIKE ${`%${u}%`} AND kainame LIKE ${`%${k}%`}`;
  }
  if (u) {
    return Prisma.sql`AND uno LIKE ${`%${u}%`}`;
  }
  if (k) {
    return Prisma.sql`AND kainame LIKE ${`%${k}%`}`;
  }
  return Prisma.empty;
}

/** GET /api/dosimetry/irradiation-results */
irradiationResultsRouter.get("/", requireAuth, async (req, res) => {
  const parsed = querySchema.safeParse({
    source: firstQueryString(req.query.source),
    uno: firstQueryString(req.query.uno),
    kainame: firstQueryString(req.query.kainame),
    page: firstQueryString(req.query.page),
    pageSize: firstQueryString(req.query.pageSize),
  });
  if (!parsed.success) {
    res.status(400).json({ error: "Invalid query" });
    return;
  }
  const { source, uno, kainame, page, pageSize } = parsed.data;
  const offset = (page - 1) * pageSize;

  try {
    if (source === "syouk1") {
      const w = whereSyouk1(uno, kainame);
      const [total, list] = await Promise.all([
        prisma.syouk1.count({ where: w }),
        prisma.syouk1.findMany({
          where: w,
          orderBy: { uno: "asc" },
          skip: offset,
          take: pageSize,
        }),
      ]);
      res.json({
        data: {
          rows: list.map(mapSyouk1),
          meta: { total, page, pageSize, totalPages: Math.max(1, Math.ceil(total / pageSize)) },
        },
      });
      return;
    }

    if (source === "syouk2") {
      const w = whereSyouk2(uno, kainame);
      const [total, list] = await Promise.all([
        prisma.syouk2.count({ where: w }),
        prisma.syouk2.findMany({
          where: w,
          orderBy: { uno: "asc" },
          skip: offset,
          take: pageSize,
        }),
      ]);
      res.json({
        data: {
          rows: list.map(mapSyouk2),
          meta: { total, page, pageSize, totalPages: Math.max(1, Math.ceil(total / pageSize)) },
        },
      });
      return;
    }

    if (source === "syoukj3") {
      const w = whereSyoukj3(uno, kainame);
      const [total, list] = await Promise.all([
        prisma.syoukj3.count({ where: w }),
        prisma.syoukj3.findMany({
          where: w,
          orderBy: { uno: "asc" },
          skip: offset,
          take: pageSize,
        }),
      ]);
      res.json({
        data: {
          rows: list.map(mapSyoukj3),
          meta: { total, page, pageSize, totalPages: Math.max(1, Math.ceil(total / pageSize)) },
        },
      });
      return;
    }

    const filterSql = unoKainameSql(uno, kainame);
    const countRows = await prisma.$queryRaw<[{ c: bigint | number }]>(
      Prisma.sql`
        SELECT COUNT(*) AS c
        FROM (
          SELECT uno FROM syouk1 WHERE 1=1 ${filterSql}
          UNION ALL
          SELECT uno FROM syouk2 WHERE 1=1 ${filterSql}
          UNION ALL
          SELECT uno FROM syoukj3 WHERE 1=1 ${filterSql}
        ) t
      `,
    );
    const total = Number(countRows[0]?.c ?? 0);

    const rawList = await prisma.$queryRaw<RawUnionRow[]>(
      Prisma.sql`
        SELECT * FROM (
          SELECT
            'syouk1' AS src, uno, syono, kainame, siteisn, syosuu, syoichi, syotime, hansuu, stimer, ktimer,
            CAST(senritu AS CHAR(32)) AS senritu, syostat, ctimer, zhansuu, htimer, slotno, sdate, edate, updfflg, syokind, bikou
          FROM syouk1
          WHERE 1=1 ${filterSql}
          UNION ALL
          SELECT
            'syouk2' AS src, uno, syono, kainame, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, sdate, edate, NULL, syokind, bikou
          FROM syouk2
          WHERE 1=1 ${filterSql}
          UNION ALL
          SELECT
            'syoukj3' AS src, uno, syono, kainame, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, sdate, edate, NULL, syokind, bikou
          FROM syoukj3
          WHERE 1=1 ${filterSql}
        ) u
        ORDER BY u.uno ASC
        LIMIT ${pageSize} OFFSET ${offset}
      `,
    );

    res.json({
      data: {
        rows: rawList.map(mapRaw),
        meta: {
          total,
          page,
          pageSize,
          totalPages: Math.max(1, Math.ceil(total / pageSize)),
        },
      },
    });
  } catch (e) {
    res.status(500).json({ error: e instanceof Error ? e.message : "Internal error" });
  }
});

export { irradiationResultsRouter };
