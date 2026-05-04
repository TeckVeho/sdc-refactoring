/**
 * Exガンマ照射課実績集計 — 集計シートの受付番号・完了月日範囲（Excel 数式相当）
 */

export function unoRangeFromYearMonth(year: number, month: number): { unoMin: string; unoMax: string } {
  const first = new Date(year, month - 1, 1);
  const minus300d = new Date(first.getTime() - 300 * 24 * 60 * 60 * 1000);
  const y1 = minus300d.getFullYear();
  const m1 = minus300d.getMonth() + 1;
  const unoMin = `${y1}${String(m1).padStart(2, "0")}0000`;

  const plus31d = new Date(first.getTime() + 31 * 24 * 60 * 60 * 1000);
  const y2 = plus31d.getFullYear();
  const m2 = plus31d.getMonth() + 1;
  const unoMax = `${y2}${String(m2).padStart(2, "0")}9999`;

  return { unoMin, unoMax };
}

/** 集計!O5:P5 相当（完了月日フィルタ用文字列レンジ — MVP は ctimer との単純比較用） */
export function kkmdRangeFromMonth(month: number): { kkmd1: string; kkmd2: string } {
  const kkmd1 = `${String(month).padStart(2, "0")}00`;
  const kkmd2 = `${String(month).padStart(2, "0")}32`;
  return { kkmd1, kkmd2 };
}

export function ctimerInKkmdRange(ctimer: string | null | undefined, kkmd1: string, kkmd2: string): boolean {
  if (ctimer == null || ctimer.trim() === "") return false;
  const t = ctimer.trim();
  return t >= kkmd1 && t <= kkmd2;
}

export function parseSiteisnKg(siteisn: string | null | undefined): number | null {
  if (siteisn == null || siteisn.trim() === "") return null;
  const n = Number(String(siteisn).replace(/,/g, "").trim());
  return Number.isFinite(n) ? n : null;
}
