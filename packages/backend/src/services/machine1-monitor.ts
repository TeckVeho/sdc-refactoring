import { prisma } from "../lib/prisma.js";
import { trimStr } from "../lib/trim-str.js";

/** Ex装置 §2.3 / 線源 state codes (SENGNR1.EVENT) */
const SOURCE_EVENT_LABEL: Record<string, string> = {
  "0": "昇降中",
  "1": "照射中",
  "2": "貯蔵中",
  "3": "移動照射中",
};

/** syouk1.SYOSTAT: 1:開始/2:中断/3:完了/4:取消/5:再開/6:修正（Ex1号機照射情報 §2.2） */
// 照会モニタでは未完了扱いの行のみ表示: VBA 照射データのうち、完了(3)・取消(4)を除きつつ実運用上は 1,2,5,6 を対象にする
const IN_PROGRESS_SYOSTAT = ["1", "2", "5", "6"] as const;

const POSITION_CODE_LABEL: Record<string, string> = {
  "11": "南ｺﾝ",
  "12": "南固",
  "21": "北ｺﾝ",
  "22": "北固",
  "31": "特１",
  "32": "特２",
};

export type Machine1MonitorSourceState = {
  label: string;
  sdate?: string;
  stime?: string;
  event?: string;
  timer?: string;
};

export type Machine1MonitorRow = {
  no: number;
  receiptNo: string;
  company: string;
  dose: string;
  qty: number;
  position: string;
  remaining: string;
  eta: string;
  shipDate: string;
  note: string;
};

export type Machine1MonitorData = {
  updatedAt: string;
  sourceState: Machine1MonitorSourceState;
  rows: Machine1MonitorRow[];
};

function formatYmd(ymd: string | null | undefined): string {
  const y = trimStr(ymd);
  if (!y || y.length < 8) return y;
  const d = y.replace(/\D/g, "");
  if (d.length < 8) return y;
  return `${d.slice(0, 4)}/${d.slice(4, 6)}/${d.slice(6, 8)}`;
}

function formatHhmmss(t: string | null | undefined): string {
  const raw = trimStr(t);
  if (!raw) return "";
  const d = raw.replace(/\D/g, "");
  if (d.length < 4) return raw;
  if (d.length >= 6) {
    return `${d.slice(0, 2)}:${d.slice(2, 4)}:${d.slice(4, 6)}`;
  }
  return `${d.slice(0, 2)}:${d.slice(2, 4)}`;
}

function parseNum(s: string | null | undefined): number | null {
  if (s == null) return null;
  const t = trimStr(s);
  if (t === "") return null;
  const n = parseFloat(t.replace(/,/g, ""));
  return Number.isFinite(n) ? n : null;
}

/**
 * 完了までの目安（分）: 完了側タイマー − 現在タイマー（HTIMER/KTIMER と CTIMER を数値化して試行）
 */
function remainingMinutesHminusC(row: {
  htimer: string | null;
  ktimer: string | null;
  ctimer: string | null;
}): number | null {
  const end = parseNum(row.htimer) ?? parseNum(row.ktimer);
  const cur = parseNum(row.ctimer);
  if (end === null || cur === null) return null;
  return end - cur;
}

/**
 * ざっくり人間可読の残り時間。VBA の線源補正・1.2倍等は未実装（簡易版）
 */
function formatRemainingMinutes(m: number | null): string {
  if (m === null) return "—";
  if (m < 0) return "0 分";
  if (m < 1) return `${(m * 60).toFixed(0)} 秒`;
  if (m < 60) return `${m.toFixed(1)} 分`;
  const h = Math.floor(m / 60);
  const min = m - h * 60;
  return `${h} 時間 ${min.toFixed(0)} 分`;
}

function formatEtaFromMinutes(m: number | null): string {
  if (m === null) return "—";
  const now = new Date();
  if (m < 0) {
    return now.toLocaleString("ja-JP", { timeZone: "Asia/Tokyo" });
  }
  const end = new Date(now.getTime() + m * 60_000);
  return end.toLocaleString("ja-JP", { timeZone: "Asia/Tokyo" });
}

type KeikakuByUno = Map<string, { shipDate: string; note: string }>;

function readKeikakuJson(
  j: import("@prisma/client").Prisma.JsonValue,
  sink: KeikakuByUno,
): void {
  if (j === null || typeof j !== "object" || Array.isArray(j)) return;
  const o = j as Record<string, unknown>;
  const souti = String(o.souti ?? o.SOUTI ?? "").trim();
  if (souti !== "1" && souti !== "1.0") return;
  const rawUno = o.uno ?? o.UNO;
  if (rawUno === undefined || rawUno === null) return;
  const uno = String(rawUno).trim();
  const ship = String(o.syukkabi ?? o.SYUKKABI ?? "")
    .trim()
    .replace(/\s+/g, " ");
  const note = String(o.bikou1 ?? o.BIKOU1 ?? "")
    .trim()
    .replace(/\s+/g, " ");
  sink.set(uno, { shipDate: ship, note });
}

/**
 * 1号機出荷予定: ExKeikakuX の JSON から受付番号 (uno) ベースのマッピングを作成。
 * 列ベース正規化 DB 移行前提で json 形は揺れうるため、小文字キーを優先して走査する。
 */
async function buildKeikakuUnoMap(): Promise<KeikakuByUno> {
  const rows = await prisma.exKeikakuX.findMany({
    select: { json: true },
  });
  const map: KeikakuByUno = new Map();
  for (const r of rows) {
    const j = r.json;
    if (j === null) continue;
    if (typeof j === "object" && j !== null && !Array.isArray(j)) {
      readKeikakuJson(j, map);
    } else {
      try {
        const s = JSON.stringify(j);
        const p = JSON.parse(s) as unknown;
        if (p && typeof p === "object" && !Array.isArray(p)) {
          readKeikakuJson(p as import("@prisma/client").Prisma.JsonValue, map);
        }
      } catch {
        // ignore
      }
    }
  }
  return map;
}

function positionLabel(syoichi: string | null | undefined): string {
  if (syoichi == null) return "??";
  const c = String(syoichi).trim();
  return POSITION_CODE_LABEL[c] ?? "??";
}

/**
 * モニタ用スナップショット。HTTP / SSE 共通
 */
export async function getMachine1MonitorData(): Promise<Machine1MonitorData> {
  const [latestSeng, syoukRows, keikaku] = await Promise.all([
    prisma.sengnr1.findFirst({
      orderBy: { id: "desc" },
    }),
    prisma.syouk1.findMany({
      where: {
        // syokind <> '2'（VBA）: NULL は従来 SQL と同様に対象に含めたい
        OR: [{ syokind: null }, { syokind: { not: "2" } }],
        syostat: { in: [...IN_PROGRESS_SYOSTAT] },
      },
    }),
    buildKeikakuUnoMap(),
  ]);

  const eventCode = trimStr(latestSeng?.event);
  const sourceState: Machine1MonitorSourceState = {
    label: SOURCE_EVENT_LABEL[eventCode] ?? "不明",
    sdate: latestSeng ? formatYmd(latestSeng.sdate) : undefined,
    stime: latestSeng ? formatHhmmss(latestSeng.stime) : undefined,
    event: eventCode || undefined,
    timer: latestSeng?.timer != null ? String(latestSeng.timer) : undefined,
  };

  const sengEvent = eventCode;
  const rowsPrelim: Array<{
    receiptNo: string;
    company: string;
    dose: string;
    qty: number;
    position: string;
    remaining: string;
    eta: string;
    shipDate: string;
    note: string;
    _sort: number;
  }> = [];

  for (const s of syoukRows) {
    const uno = trimStr(s.uno);
    const stat = trimStr(s.syostat);
    const k = keikaku.get(uno) ?? { shipDate: "", note: "" };

    let remaining: string;
    let eta: string;
    if (stat === "2") {
      remaining = "-----";
      eta = "中断中";
    } else if (sengEvent === "3" && trimStr(s.syokind) === "1") {
      remaining = "------";
      eta = "担当者に確認";
    } else {
      const mins = remainingMinutesHminusC(s);
      remaining = formatRemainingMinutes(mins);
      eta = formatEtaFromMinutes(mins);
    }

    const sortMins = (() => {
      if (stat === "2") return 1e9;
      if (sengEvent === "3" && trimStr(s.syokind) === "1") return 1e9 - 1;
      const m = remainingMinutesHminusC(s);
      return m ?? 1e9 - 0.5;
    })();

    rowsPrelim.push({
      receiptNo: uno,
      company: trimStr(s.kainame),
      dose: trimStr(s.siteisn),
      qty: s.syosuu ?? 0,
      position: positionLabel(s.syoichi),
      remaining,
      eta,
      shipDate: k.shipDate,
      note: k.note,
      _sort: sortMins,
    });
  }

  rowsPrelim.sort((a, b) => a._sort - b._sort);

  const rows: Machine1MonitorRow[] = rowsPrelim.map((r, i) => ({
    no: i + 1,
    receiptNo: r.receiptNo,
    company: r.company,
    dose: r.dose,
    qty: r.qty,
    position: r.position,
    remaining: r.remaining,
    eta: r.eta,
    shipDate: r.shipDate,
    note: r.note,
  }));

  return {
    updatedAt: new Date().toISOString(),
    sourceState,
    rows,
  };
}
