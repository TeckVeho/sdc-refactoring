/**
 * 線量率計算v1.1 相当の純粋計算（仕様: docs/線量率計算v1.1_仕様書.md）
 * 電離箱校正定数の実数は校正Excelと同期要（MVP では代表値を格納）
 */

import { decayFactor } from "./radiation-decay.js";

export const DOSE_RATE_CALC_VERSION = "1.1" as const;

export const MASS_CO_WATER = 0.02965;
export const MASS_CS_AIR = 0.02932;
/** Cs 空気 → Co 水 換算（X11） */
export const CS_TO_CO_WATER_FACTOR = MASS_CO_WATER / MASS_CS_AIR;

export const CABLE_NUMERATOR = 0.02666; // Co 空気
export const CABLE_DENOMINATOR = 0.02965; // Co 水
export const CABLE_CORRECTION_DEFAULT = CABLE_NUMERATOR / CABLE_DENOMINATOR; // 0.8991…

const POTENTIOMETER_LABELS = ["#3302229", "#4502260"] as const;

/** 線量率 RATE: 電位計校正定数 S6, S7 */
export const RATE_POTENTIOMETER_CONST = [1.0013, 1.0] as const;
/** 積算 INTEGR: R6, R7 */
export const INTEGR_POTENTIOMETER_CONST = [0.9996, 0.9997] as const;

/** 電位計不確かさ Y32（%）JQA 基準、ANTM 計算でも同参照 */
export const POTENTIOMETER_UNCERTAINTY_PCT = 0.6;

export type PotentiometerOption = {
  id: number;
  label: (typeof POTENTIOMETER_LABELS)[number];
  rateConst: number;
  integrConst: number;
};

export const POTENTIOMETER_OPTIONS: readonly PotentiometerOption[] = [
  {
    id: 1,
    label: "#3302229",
    rateConst: RATE_POTENTIOMETER_CONST[0],
    integrConst: INTEGR_POTENTIOMETER_CONST[0],
  },
  {
    id: 2,
    label: "#4502260",
    rateConst: RATE_POTENTIOMETER_CONST[1],
    integrConst: INTEGR_POTENTIOMETER_CONST[1],
  },
];

/** 各電離箱 ANTM 用不確かさ（AD 列、%）— 機番順 */
export const ION_CHAMBER_UNCERTAINTY_ANTM: readonly number[] = [
  0.8, 0.75, 0.85, 0.8, 0.82, 0.88, 0.86, 0.84, 0.87,
];

/**
 * 電離箱: Co 水校正（#1253…）と Cs 空気系（T×換算）を仕様の丸めに合わせた代表値
 * index 0..4: Co 水、5..8: Cs 換算後
 */
const T_CS_RAW = [0.117, 0.1175, 0.118, 0.1185] as const;

function roundDownN(x: number, n: number): number {
  const f = 10 ** n;
  return Math.floor(x * f + 1e-12) / f;
}

function roundN(x: number, n: number): number {
  const f = 10 ** n;
  return Math.round(x * f) / f;
}

function buildCsChamberConstants(): [number, number, number, number] {
  const t17 = T_CS_RAW[0];
  const t18 = T_CS_RAW[1];
  const t19 = T_CS_RAW[2];
  const t20 = T_CS_RAW[3];
  const s17 = roundDownN(t17 * CS_TO_CO_WATER_FACTOR, 6);
  const s18 = roundDownN(t18 * CS_TO_CO_WATER_FACTOR, 7);
  const s19 = roundN(t19 * CS_TO_CO_WATER_FACTOR, 7);
  const s20 = roundDownN(t20 * CS_TO_CO_WATER_FACTOR, 8);
  return [s17, s18, s19, s20];
}

const CS_S = buildCsChamberConstants();

const RATE_ION_S12_S20: readonly number[] = [
  0.1195,
  0.119,
  0.1185,
  0.118,
  0.1175,
  CS_S[0],
  CS_S[1],
  CS_S[2],
  CS_S[3],
];

export type IonChamberEntry = {
  id: number;
  model: string;
  calConstantGyPerNc: number;
  antmUncertaintyPct: number;
  /** JQA 第2不確かさ用（仕様 T32: AD36 等 — MVP では chamber 縦揃え） */
  jqaSecondUncertaintyPct: number;
};

export const ION_CHAMBERS: readonly IonChamberEntry[] = [
  { id: 1, model: "#1253", calConstantGyPerNc: RATE_ION_S12_S20[0], antmUncertaintyPct: ION_CHAMBER_UNCERTAINTY_ANTM[0]!, jqaSecondUncertaintyPct: 0.7 },
  { id: 2, model: "#1289", calConstantGyPerNc: RATE_ION_S12_S20[1], antmUncertaintyPct: ION_CHAMBER_UNCERTAINTY_ANTM[1]!, jqaSecondUncertaintyPct: 0.68 },
  { id: 3, model: "#1144", calConstantGyPerNc: RATE_ION_S12_S20[2], antmUncertaintyPct: ION_CHAMBER_UNCERTAINTY_ANTM[2]!, jqaSecondUncertaintyPct: 0.72 },
  { id: 4, model: "#1166", calConstantGyPerNc: RATE_ION_S12_S20[3], antmUncertaintyPct: ION_CHAMBER_UNCERTAINTY_ANTM[3]!, jqaSecondUncertaintyPct: 0.71 },
  { id: 5, model: "#1188", calConstantGyPerNc: RATE_ION_S12_S20[4], antmUncertaintyPct: ION_CHAMBER_UNCERTAINTY_ANTM[4]!, jqaSecondUncertaintyPct: 0.7 },
  { id: 6, model: "#1200", calConstantGyPerNc: RATE_ION_S12_S20[5], antmUncertaintyPct: ION_CHAMBER_UNCERTAINTY_ANTM[5]!, jqaSecondUncertaintyPct: 0.74 },
  { id: 7, model: "#1202", calConstantGyPerNc: RATE_ION_S12_S20[6], antmUncertaintyPct: ION_CHAMBER_UNCERTAINTY_ANTM[6]!, jqaSecondUncertaintyPct: 0.73 },
  { id: 8, model: "#1203", calConstantGyPerNc: RATE_ION_S12_S20[7], antmUncertaintyPct: ION_CHAMBER_UNCERTAINTY_ANTM[7]!, jqaSecondUncertaintyPct: 0.72 },
  { id: 9, model: "#1205", calConstantGyPerNc: RATE_ION_S12_S20[8], antmUncertaintyPct: ION_CHAMBER_UNCERTAINTY_ANTM[8]!, jqaSecondUncertaintyPct: 0.75 },
];

const PA_TO_NC_PER_H = 3.6;

/**
 * 温度・気圧補正 κTP（4 桁切り捨て、標準 22℃ / 1013.3 hPa）
 */
export function kappaTp(temperatureC: number, pressureHpa: number): number {
  return roundDownN((273.2 + temperatureC) / (273.2 + 22) * (1013.3 / pressureHpa), 4);
}

/**
 * 線量率 RATE: Excel H37 相当（有効数字 3 桁・切り捨て）
 */
export function roundDown3Significant(x: number): number {
  if (x === 0 || !Number.isFinite(x)) {
    return x;
  }
  const sign = x < 0 ? -1 : 1;
  const ax = Math.abs(x);
  const nDecimals = 3 - Math.floor(Math.log10(ax));
  const f = 10 ** nDecimals;
  return (sign * Math.floor(ax * f + 1e-12)) / f;
}

function uncertaintyFactorAntm(potU: number, chamberU: number): number {
  const t31 = Math.sqrt(potU * potU + chamberU * chamberU);
  return roundDownN(1 - t31 / 100, 4);
}

function uncertaintyFactorJqa(potU: number, chamber2U: number): number {
  const t32 = Math.sqrt(potU * potU + chamber2U * chamber2U);
  return roundDownN(1 - t32 / 100, 4);
}

function pickUncertainty(
  org: "ANTM" | "JQA" | "none",
  potU: number,
  chamber: IonChamberEntry,
): { factor: number; combinedPct: number } {
  if (org === "none") {
    return { factor: 1, combinedPct: 0 };
  }
  if (org === "ANTM") {
    const t = Math.sqrt(potU * potU + chamber.antmUncertaintyPct * chamber.antmUncertaintyPct);
    return { factor: uncertaintyFactorAntm(potU, chamber.antmUncertaintyPct), combinedPct: t };
  }
  const t2 = Math.sqrt(potU * potU + chamber.jqaSecondUncertaintyPct * chamber.jqaSecondUncertaintyPct);
  return { factor: uncertaintyFactorJqa(potU, chamber.jqaSecondUncertaintyPct), combinedPct: t2 };
}

export type DoseRateCalcInput = {
  mode: "RATE" | "INTEG";
  potentiometerId: 1 | 2;
  ionChamberId: number;
  readValue: number;
  /** 温度 (℃) — tp ありのとき必須 */
  temperatureC: number;
  /** 気圧 (hPa) */
  pressureHpa: number;
  tpEnabled: boolean;
  uncertaintyOrg: "ANTM" | "JQA" | "none";
  cableOn: boolean;
  /** 設定線量 (Gy) — 指定時に照射時間を返す */
  targetDoseGy?: number;
  /** 減衰: 半減期(日) と 経過日数 — 最終の線量率/積算に乗算 */
  decay?: { halfLifeDays: number; elapsedDays: number } | null;
};

export type DoseRateCalcResult = {
  kappaTp: number;
  cableFactor: number;
  potentiometerConst: number;
  ionChamberConst: number;
  uncertaintyFactor: number;
  uncertaintyCombinedPct: number;
  /** RATE: pA 経路、INTEG: nC 経路 */
  rawProduct: number;
  /** 減衰前 */
  doseRateGyPerH: number | null;
  integratedDoseGy: number | null;
  /** 有効数字 3 桁切り捨て後 */
  displayDoseRateGyPerH: number | null;
  displayIntegratedDoseGy: number | null;
  decayFactor: number;
  displayDoseRateAfterDecay: number | null;
  displayIntegratedDoseAfterDecay: number | null;
  /** 照射時間 (h) — 設定線量/表示線量率、小数第 2 位切り上げ */
  exposureTimeH: number | null;
};

function getChamberById(ionChamberId: number): IonChamberEntry {
  const ch = ION_CHAMBERS.find((c) => c.id === ionChamberId);
  if (!ch) {
    throw new RangeError(`ionChamberId out of range: ${ionChamberId}`);
  }
  return ch;
}

function getPotById(potentiometerId: 1 | 2): PotentiometerOption {
  return POTENTIOMETER_OPTIONS.find((p) => p.id === potentiometerId) ?? POTENTIOMETER_OPTIONS[0]!;
}

export function computeDoseRate(input: DoseRateCalcInput): DoseRateCalcResult {
  const pot = getPotById(input.potentiometerId);
  const chamber = getChamberById(input.ionChamberId);
  const kT = input.tpEnabled ? kappaTp(input.temperatureC, input.pressureHpa) : 1;
  const cable = input.cableOn ? CABLE_CORRECTION_DEFAULT : 1;
  const pConst = input.mode === "RATE" ? pot.rateConst : pot.integrConst;
  const { factor: uFact, combinedPct } = pickUncertainty(
    input.uncertaintyOrg,
    POTENTIOMETER_UNCERTAINTY_PCT,
    chamber,
  );
  const kIon = chamber.calConstantGyPerNc;

  let raw: number;
  if (input.mode === "RATE") {
    raw = pConst * kIon * kT * input.readValue * uFact * cable * PA_TO_NC_PER_H;
  } else {
    raw = pConst * kIon * kT * input.readValue * uFact * cable;
  }

  const dRate = input.mode === "RATE" ? raw : null;
  const integ = input.mode === "INTEG" ? raw : null;

  const displayRate = dRate != null ? roundDown3Significant(dRate) : null;
  const displayInteg = integ != null ? roundDown3Significant(integ) : null;

  const decayF =
    input.decay && input.decay.halfLifeDays > 0
      ? decayFactor(input.decay.elapsedDays, input.decay.halfLifeDays)
      : 1;
  const drDecay = displayRate != null ? displayRate * decayF : null;
  const inDecay = displayInteg != null ? displayInteg * decayF : null;

  let exposureTimeH: number | null = null;
  if (input.mode === "RATE" && input.targetDoseGy != null && input.targetDoseGy > 0) {
    const base = drDecay ?? displayRate;
    if (base != null && base > 0) {
      exposureTimeH = Math.ceil((input.targetDoseGy / base) * 100) / 100;
    }
  }

  return {
    kappaTp: kT,
    cableFactor: cable,
    potentiometerConst: pConst,
    ionChamberConst: kIon,
    uncertaintyFactor: uFact,
    uncertaintyCombinedPct: combinedPct,
    rawProduct: raw,
    doseRateGyPerH: dRate,
    integratedDoseGy: integ,
    displayDoseRateGyPerH: displayRate,
    displayIntegratedDoseGy: displayInteg,
    decayFactor: decayF,
    displayDoseRateAfterDecay: drDecay,
    displayIntegratedDoseAfterDecay: inDecay,
    exposureTimeH,
  };
}

export function getCoefficientsPayload() {
  return {
    version: DOSE_RATE_CALC_VERSION,
    specRef: "docs/線量率計算v1.1_仕様書.md",
    constants: {
      massAbsorptionCoWater: MASS_CO_WATER,
      massAbsorptionCsAir: MASS_CS_AIR,
      csToCoWater: CS_TO_CO_WATER_FACTOR,
      cableNumerator: CABLE_NUMERATOR,
      cableDenominator: CABLE_DENOMINATOR,
      cableFactorWhenEnabled: CABLE_CORRECTION_DEFAULT,
      potentiometerUncertaintyPct: POTENTIOMETER_UNCERTAINTY_PCT,
      paToNcPerH: PA_TO_NC_PER_H,
    },
    potentiometerOptions: POTENTIOMETER_OPTIONS,
    ionChambers: ION_CHAMBERS,
  };
}
