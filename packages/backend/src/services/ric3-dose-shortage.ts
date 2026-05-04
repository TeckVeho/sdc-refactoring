import type { PrismaClient } from "@prisma/client";

/** 通常版 Excel / 住電版 Excel のレイアウト・数式差分フラグ */
export type Ric3DoseShortageVariant = "standard" | "sumiden";

export type Ric3LayoutMeta = {
  variant: Ric3DoseShortageVariant;
  reportTitle: string;
  formCode: string;
  /** Excel の SyainTB 想定行数（通常 36、住電 39） */
  employeeTableRowCap: number;
  /** 追加パス数（AA 列）の計算モード */
  additionalPassRounding: "simple_roundup" | "sumiden_margin_lt2";
};

export type Ric3MeasurementRow = {
  tuikaN: string;
  sokDate: string;
  sokutCd: string;
  saiSoku: string;
  sokutSn: number | null;
  tani: string;
  pass: number | null;
  sikiKodo: string;
  sosi: string;
  atusa: number | null;
};

export type Ric3PlanRow = {
  ukeNo: string;
  ricNoS: string;
  ricNoE: string;
  siteiSn: number | null;
  kagen: number | null;
  jyoug: number | null;
  jituP: number | null;
  kaiCd: string;
  kaiName: string;
};

export type Ric3DoseShortagePayload = {
  layout: Ric3LayoutMeta;
  /** 照射管理番号（SYKNO）候補 — RIC 系カラムが Prisma に揃い次第まとめる */
  senkNoOptions: string[];
  sebkNoCount: number;
  employees: { shano: string; shaname: string }[];
  senkindTypes: { code: string; label: string }[];
  /** 仕様 mokuteki（製品目的）テーブル相当 */
  mokutekiTable: { code: string; description: string }[];
  detail: {
    measurement: Ric3MeasurementRow | null;
    planRows: Ric3PlanRow[];
  } | null;
};

function layoutMeta(variant: Ric3DoseShortageVariant): Ric3LayoutMeta {
  if (variant === "sumiden") {
    return {
      variant: "sumiden",
      reportTitle: "３号機線量不足報告書　住友電工専用（工程確認・追加照射指示書）",
      formCode: "G03-08",
      employeeTableRowCap: 39,
      additionalPassRounding: "sumiden_margin_lt2",
    };
  }
  return {
    variant: "standard",
    reportTitle: "３号機線量不足報告書（工程確認・追加照射指示書）",
    formCode: "G03-08",
    employeeTableRowCap: 36,
    additionalPassRounding: "simple_roundup",
  };
}

/**
 * Excel AA 列（追加パス数）相当。
 * - standard: ROUNDUP(O18/($S$7/AA9),0)
 * - sumiden: 上記に加え、ROUNDUP結果の線量余裕が 2 未満なら +1
 */
export function computeRic3AdditionalPasses(
  rounding: Ric3LayoutMeta["additionalPassRounding"],
  shortage: number,
  measuredDose: number,
  actualPasses: number,
): number | null {
  if (actualPasses <= 0 || measuredDose <= 0) return null;
  const dosePerPass = measuredDose / actualPasses;
  if (dosePerPass <= 0) return null;
  const base = Math.ceil(shortage / dosePerPass - 1e-9);
  if (rounding === "simple_roundup") {
    return base;
  }
  const roundPortion = base * dosePerPass;
  if (roundPortion - shortage < 2) {
    return base + 1;
  }
  return base;
}

/**
 * RIC3 線量不足報告書（通常／住電）共通ペイロード。
 * 明細は Oracle 移行カラム整備まで `detail: null`（または今後 senkNo で拡張）。
 */
export async function buildRic3DoseShortagePayload(
  prisma: PrismaClient,
  variant: Ric3DoseShortageVariant,
  _senkNo?: string,
): Promise<Ric3DoseShortagePayload> {
  const layout = layoutMeta(variant);

  const employeesRaw = await prisma.shainmst.findMany({
    where: { hshika: "1" },
    select: { shano: true, shaname: true },
    orderBy: { shano: "asc" },
    take: layout.employeeTableRowCap,
  });

  const senkinds = await prisma.senkind.findMany({
    orderBy: { kindcd: "asc" },
    select: { kindcd: true, name: true },
  });

  return {
    layout,
    senkNoOptions: [],
    sebkNoCount: 0,
    employees: employeesRaw.map((e) => ({
      shano: e.shano.trim(),
      shaname: String(e.shaname ?? "").trim(),
    })),
    senkindTypes: senkinds.map((s) => ({
      code: s.kindcd.trim(),
      label: String(s.name ?? "").trim() || s.kindcd.trim(),
    })),
    mokutekiTable: [
      { code: "1", description: "一般品" },
      { code: "2", description: "医療機器" },
      { code: "3", description: "滅菌対象" },
      { code: "4", description: "その他" },
    ],
    detail: null,
  };
}
