/**
 * ガンマ勤務表: 期間は「当月11日〜翌月10日」（仕様: S_Day は必ず11日）
 */

export type ShiftCycleInfo = {
  year: number;
  month: number;
  /** 期間開始（YYYY-MM-DD） */
  startIso: string;
  /** 期間終了（YYYY-MM-DD） */
  endIso: string;
  /** カレンダー列（連続日付） */
  dates: string[];
};

function pad2(n: number): string {
  return String(n).padStart(2, "0");
}

export function toIsoDateLocal(d: Date): string {
  return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
}

/** 年月から勤務サイクル（11日始まり）を算出 */
export function shiftCycleFromYearMonth(year: number, month: number): ShiftCycleInfo {
  const start = new Date(year, month - 1, 11);
  const end = new Date(year, month, 10);
  const dates: string[] = [];
  for (let t = start.getTime(); t <= end.getTime(); t += 24 * 60 * 60 * 1000) {
    dates.push(toIsoDateLocal(new Date(t)));
  }
  return {
    year,
    month,
    startIso: toIsoDateLocal(start),
    endIso: toIsoDateLocal(end),
    dates,
  };
}
