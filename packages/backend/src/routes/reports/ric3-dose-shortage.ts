import { Router } from "express";
import { z } from "zod";

import { prisma } from "../../lib/prisma.js";
import { requireAuth } from "../../middleware/auth.js";

const ric3DoseShortageRouter = Router();

const LIST_CAP = 500;

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

/** Excel ROUNDUP(x,0) に相当（正の数のみ想定） */
function roundUp0(x: number): number {
  return Math.ceil(x - 1e-12);
}

function right4(s: string | null | undefined): string {
  if (s == null) return "";
  const t = String(s).trim();
  return t.length <= 4 ? t : t.slice(-4);
}

const querySchema = z.object({
  uno: z.string().min(1).optional(),
  senkNo: z.string().min(1).optional(),
});

type Ric3DoseShortagePlanRow = {
  slot: number;
  ukeNo: string | null;
  ricNoS: string | null;
  ricNoE: string | null;
  siteiSn: number | null;
  kagen: number | null;
  joug: number | null;
  jituP: number | null;
  kaiCd: string | null;
  kaiName: string | null;
};

type Ric3DoseShortageAdditionalRow = {
  slot: number;
  ukeNo: string | null;
  ricNoS: string | null;
  ricNoE: string | null;
  shortage: number | null;
  shortageDup: number | null;
  marginFromUpper: number | null;
  /** 一般版: ROUNDUP((下限−測定)/ (測定線量/実パス), 0) */
  extraPasses: number | null;
  kaiName: string | null;
};

/** MVP: Syoukj3 列が最小のため、3〜5 スロットは空。一般版の追加パス数は単純 ROUNDUP。 */
ric3DoseShortageRouter.get("/", requireAuth, async (req, res) => {
  const parsed = querySchema.safeParse({
    uno: firstQueryString(req.query.uno),
    senkNo: firstQueryString(req.query.senkNo),
  });
  if (!parsed.success) {
    res.status(400).json({ error: "Invalid query" });
    return;
  }

  const qUno = parsed.data.uno?.trim();
  const qSenk = parsed.data.senkNo?.trim();

  try {
    if (!qUno && !qSenk) {
      const [rows, shortageCount] = await Promise.all([
        prisma.syoukj3.findMany({
          take: LIST_CAP,
          orderBy: { uno: "desc" },
          select: { uno: true, syono: true, kainame: true },
        }),
        prisma.syoukj3.count(),
      ]);

      res.json({
        data: {
          kind: "list" as const,
          shortageCount,
          candidates: rows.map((r) => ({
            uno: r.uno,
            syono: r.syono ?? null,
            kainame: r.kainame ?? null,
          })),
          truncated: rows.length >= LIST_CAP,
        },
      });
      return;
    }

    const sykj =
      qUno != null
        ? await prisma.syoukj3.findUnique({ where: { uno: qUno } })
        : await prisma.syoukj3.findFirst({
            where: {
              OR: [{ syono: qSenk }, { uno: qSenk ?? "" }],
            },
          });

    if (!sykj) {
      res.status(404).json({ error: "照射計画（syoukj3）が見つかりません" });
      return;
    }

    const uno = sykj.uno;
    const senkNo = sykj.syono?.trim() || uno;

    const [zk, sk1, j2, employees] = await Promise.all([
      prisma.zaiko.findUnique({ where: { uno } }),
      prisma.syouk1.findUnique({ where: { uno } }),
      prisma.syouj2.findFirst({ where: { uno } }),
      prisma.shainmst.findMany({
        take: 40,
        orderBy: { shano: "asc" },
        select: { shano: true, shaname: true },
      }),
    ]);

    const sokutSn = parseNum(sk1?.senritu != null ? String(sk1.senritu) : null);
    const siteiSn = parseNum(sk1?.siteisn);
    const jituP = sk1?.syosuu != null ? sk1.syosuu : null;

    const plans: Ric3DoseShortagePlanRow[] = [];
    for (let slot = 1; slot <= 5; slot += 1) {
      if (slot === 1) {
        plans.push({
          slot: 1,
          ukeNo: uno,
          ricNoS: sk1?.syono ?? null,
          ricNoE: sk1?.slotno ?? null,
          siteiSn,
          kagen: null,
          joug: null,
          jituP,
          kaiCd: zk?.kaisyacd ?? null,
          kaiName: zk?.kainame ?? sykj.kainame ?? sk1?.kainame ?? null,
        });
      } else {
        plans.push({
          slot,
          ukeNo: null,
          ricNoS: null,
          ricNoE: null,
          siteiSn: null,
          kagen: null,
          joug: null,
          jituP: null,
          kaiCd: null,
          kaiName: null,
        });
      }
    }

    const measurement = {
      tuika: null as string | null,
      sokDate: sk1?.sdate ?? j2?.status ?? null,
      soktCd: null as string | null,
      saiSoku: null as string | null,
      sokutSn,
      tani: zk?.syouso ?? null,
      pass: zk?.pass != null ? parseNum(zk.pass) : null,
      sikiKodo: sk1?.syokind ?? null,
      sosi: null as string | null,
      atusa: null as string | null,
      syokind: sykj.syokind ?? sk1?.syokind ?? null,
      bikou: sykj.bikou ?? sk1?.bikou ?? null,
      j2Meisai: j2?.meisai ?? null,
      j2Status: j2?.status ?? null,
    };

    const additional: Ric3DoseShortageAdditionalRow[] = [];

    for (let i = 0; i < 5; i += 1) {
      const p = plans[i];
      const hasUke = Boolean(p.ukeNo?.trim());
      let shortage: number | null = null;
      let marginFromUpper: number | null = null;
      let extraPasses: number | null = null;

      if (hasUke && p.kagen != null && sokutSn != null) {
        shortage = p.kagen - sokutSn;
      }
      if (hasUke && p.joug != null && sokutSn != null) {
        marginFromUpper = p.joug - sokutSn;
      }
      if (
        hasUke &&
        shortage != null &&
        sokutSn != null &&
        sokutSn !== 0 &&
        p.jituP != null &&
        p.jituP !== 0
      ) {
        extraPasses = roundUp0(shortage / (sokutSn / p.jituP));
      }

      additional.push({
        slot: i + 1,
        ukeNo: p.ukeNo,
        ricNoS: p.ricNoS,
        ricNoE: p.ricNoE,
        shortage,
        shortageDup: shortage,
        marginFromUpper,
        extraPasses,
        kaiName: p.kaiName,
      });
    }

    const numericExtras = additional.map((r) => r.extraPasses).filter((x): x is number => x != null);
    const warnings = {
      tooManyShortage: false,
      negativeExtraPasses: additional.some((r) => r.extraPasses != null && r.extraPasses < 0),
      extraPassesMismatch:
        numericExtras.length > 1 &&
        Math.max(...numericExtras) !== Math.min(...numericExtras),
    };

    res.json({
      data: {
        kind: "report" as const,
        meta: {
          formNo: "G3-08",
          title: "３号機線量不足報告書（工程確認・追加照射指示書）",
          variant: "general" as const,
        },
        uno,
        senkNo,
        doseMeterLast4: right4(senkNo),
        syoukj3: {
          syono: sykj.syono ?? null,
          kainame: sykj.kainame ?? null,
          syokind: sykj.syokind ?? null,
          sdate: sykj.sdate ?? null,
          edate: sykj.edate ?? null,
          bikou: sykj.bikou ?? null,
        },
        measurement,
        plans,
        additional,
        doseMeter: {
          senKind: sk1?.syokind ?? sykj.syokind ?? null,
          atusaT: null as string | null,
          keisask: sk1?.syokind ?? null,
        },
        warnings,
        employees,
      },
    });
  } catch (e) {
    res.status(500).json({ error: e instanceof Error ? e.message : "Server error" });
  }
});

export { ric3DoseShortageRouter };
