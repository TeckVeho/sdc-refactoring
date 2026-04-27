import { prisma } from "../lib/prisma.js";

/**
 * Ex装置運転状況（docs/Ex装置運転状況_仕様書.md）
 * §2.1 G_Data: 1号 0/0.5/1, 2号 1.2/2.2, 3号 2.4/3.4
 * §6.3 Ric1Kadou / Ric2Kadou / Ric3Kadou: 本実装は Strategy 簡易版（累積稼働時間の厳密再現は行わない）
 */

export type MachineStatusId = 1 | 2 | 3;

export type MachineStatusPoint = { t: string; value: number; machine: MachineStatusId };

export type MachineStatusSnapshot = {
  updatedAt: string;
  machines: Array<{
    id: MachineStatusId;
    timer?: string;
    statusKey: string;
    statusLabel: string;
    detail?: string;
  }>;
  chart: { points: MachineStatusPoint[] };
};

const CHART_24H_MS = 24 * 60 * 60 * 1000;
/** グラフ1系列あたりの上限（直近N点・DB負荷抑制） */
const CHART_MAX_POINTS_PER_MACHINE = 2000;

/** 1号機: 1号機集計表 イベント名（Ex §2.2 イベントコード表）＋G_Data 値対応 */
const SENGNR1_EVENT_LABEL: Record<string, string> = {
  "0": "貯蔵中",
  "1": "照射",
  "2": "昇降開始",
  "3": "移動照射",
  E: "ＰＣ Stop",
  e: "ＰＣ Stop",
  S: "PC Start",
  s: "PC Start",
};

function normalizeYyyymmdd(s: string | null | undefined): string {
  if (s == null) return "";
  const digits = String(s).replace(/\D/g, "");
  return digits.length >= 8 ? digits.slice(0, 8) : digits;
}

/** Ex §2.1 D列: 0=停止, 0.5=移動照射, 1=照射（R1Graph 簡易版） */
function sengnr1EventToGdata(event: string | null | undefined): number {
  const e = (event ?? "").trim();
  if (e === "1") return 1;
  if (e === "3") return 0.5;
  return 0;
}

function ymdHmsToLocalDate(ymd8: string, hms6: string): Date | null {
  if (ymd8.length !== 8) return null;
  const t = hms6.replace(/\D/g, "").padEnd(6, "0").slice(0, 6);
  if (t.length < 6) return null;
  const y = Number(ymd8.slice(0, 4));
  const mo = Number(ymd8.slice(4, 6)) - 1;
  const d = Number(ymd8.slice(6, 8));
  const hh = Number(t.slice(0, 2));
  const mm = Number(t.slice(2, 4));
  const ss = Number(t.slice(4, 6));
  return new Date(y, mo, d, hh, mm, ss);
}

function sengnr1RowTime(sdate: string, stime: string | null): Date | null {
  const ymd = normalizeYyyymmdd(sdate);
  if (ymd.length !== 8) return null;
  return ymdHmsToLocalDate(ymd, stime ? stime : "000000");
}

/** kyouj2.rectime: 「YYYYMMDD ... HHMMSS」等の揺れを想定 */
function parseKyouj2Rectime(rectime: string): Date | null {
  const d = String(rectime).replace(/\D/g, "");
  if (d.length < 8) return null;
  const ymd = d.slice(0, 8);
  const hms = d.length >= 14 ? d.slice(8, 14) : "000000";
  return ymdHmsToLocalDate(ymd, hms);
}

function sengnr3RowTime(sdate: string, sekitime: string | null): Date | null {
  const head = sdate.replace(/\D/g, "");
  if (head.length < 8) return null;
  const ymd = head.slice(0, 8);
  const st = (sekitime ?? "").replace(/\D/g, "");
  const hms = st.length >= 6 ? st.slice(0, 6) : "000000";
  return ymdHmsToLocalDate(ymd, hms);
}

/**
 * 3号機: sengnr3.event を 16 進・10 進いずれかで数値化し下位 16bit を参照（完全な ExHenkan 再現ではない・コメント参照）
 * G_Data: 2.4=停止系, 3.4=照射系 — 下位 bit の一例として bit1 OR bit4 で「照射目安」
 */
function decodeSengnr3Event(
  event: string | null | undefined,
): { statusKey: string; statusLabel: string; detail: string; gValue: 2.4 | 3.4 } {
  if (event == null || String(event).trim() === "") {
    return {
      statusKey: "unknown",
      statusLabel: "不明",
      detail: "event なし",
      gValue: 2.4,
    };
  }
  const raw = String(event).trim();
  let n = 0;
  if (/^0x[0-9a-fA-F]+$/.test(raw)) {
    n = Number.parseInt(raw, 16);
  } else {
    const onlyNum = raw.replace(/[^\d-]/g, "");
    n = onlyNum ? Number.parseInt(onlyNum, 10) : 0;
  }
  if (!Number.isFinite(n)) n = 0;
  const w = n & 0xffff;
  // 例示ビット（実VLAと異なる可能性あり。仕様§6.3 Ric3Kadou / R3Graph 簡易版）
  const irradiationHint = (w & 0x2) !== 0 || (w & 0x10) !== 0;
  const g: 2.4 | 3.4 = irradiationHint ? 3.4 : 2.4;
  return {
    statusKey: irradiationHint ? "irradiation_like" : "idle_like",
    statusLabel: irradiationHint ? "照射（下位16bit 簡易）" : "停止（下位16bit 簡易）",
    detail: `raw=${raw} lower16=0x${w.toString(16)} (16bit 風・厳密でない)`,
    gValue: g,
  };
}

function m2GdataFromRowAndPrev(
  current: Date,
  prev: Date | null,
): { g: 1.2 | 2.2; key: "active" | "idle" } {
  if (prev == null) return { g: 1.2, key: "idle" };
  const min = Math.abs(current.getTime() - prev.getTime()) / 60_000;
  if (min <= 5) return { g: 2.2, key: "active" };
  return { g: 1.2, key: "idle" };
}

export async function getMachineStatusSnapshot(): Promise<MachineStatusSnapshot> {
  const now = new Date();
  const since = new Date(now.getTime() - CHART_24H_MS);

  const [latest1, twoK2, latest3, s1rows, k2rows, s3rows] = await Promise.all([
    prisma.sengnr1.findFirst({ orderBy: { id: "desc" } }),
    prisma.kyouj2.findMany({ orderBy: { id: "desc" }, take: 2 }),
    prisma.sengnr3.findFirst({ orderBy: { id: "desc" } }),
    prisma.sengnr1.findMany({ orderBy: { id: "desc" }, take: CHART_MAX_POINTS_PER_MACHINE * 2 }),
    prisma.kyouj2.findMany({ orderBy: { id: "desc" }, take: CHART_MAX_POINTS_PER_MACHINE * 2 }),
    prisma.sengnr3.findMany({ orderBy: { id: "desc" }, take: CHART_MAX_POINTS_PER_MACHINE * 2 }),
  ]);

  // --- 1号機 現在状態 ---
  const e1 = latest1?.event?.trim() ?? "";
  const label1 = SENGNR1_EVENT_LABEL[e1] ?? (e1 ? `イベント(${e1})` : "データなし");
  const machine1: MachineStatusSnapshot["machines"][0] = {
    id: 1,
    timer: latest1?.timer != null ? String(latest1.timer) : undefined,
    statusKey: e1 || "empty",
    statusLabel: label1,
    detail: latest1
      ? `${normalizeYyyymmdd(latest1.sdate)} ${(latest1.stime ?? "").trim()}`.trim()
      : undefined,
  };

  // --- 2号機: 直近2件の rectime 差分が5分以内なら「稼働」---
  let machine2: MachineStatusSnapshot["machines"][1];
  if (twoK2.length < 2) {
    machine2 = {
      id: 2,
      timer: twoK2[0]?.realspd != null ? String(twoK2[0]!.realspd) : undefined,
      statusKey: "insufficient",
      statusLabel: "判定不能（2件未満）",
    };
  } else {
    const tA = parseKyouj2Rectime(twoK2[0]!.rectime);
    const tB = parseKyouj2Rectime(twoK2[1]!.rectime);
    if (!tA || !tB) {
      machine2 = {
        id: 2,
        statusKey: "unknown",
        statusLabel: "時刻解釈失敗",
        detail: `rectime=${twoK2[0]!.rectime} / ${twoK2[1]!.rectime}`,
      };
    } else {
      const diffMin = Math.abs(tA.getTime() - tB.getTime()) / 60_000;
      const active = diffMin <= 5;
      machine2 = {
        id: 2,
        timer: twoK2[0]?.realspd != null ? String(twoK2[0]!.realspd) : undefined,
        statusKey: active ? "active" : "idle",
        statusLabel: active ? "稼働" : "停止",
        detail: `直近2件の間隔 ${diffMin.toFixed(1)} 分（≤5分で稼働）`,
      };
    }
  }

  // --- 3号機 ---
  const d3 = decodeSengnr3Event(latest3?.event);
  const machine3: MachineStatusSnapshot["machines"][2] = {
    id: 3,
    timer: latest3?.sekitime != null ? String(latest3.sekitime) : undefined,
    statusKey: d3.statusKey,
    statusLabel: d3.statusLabel,
    detail: d3.detail,
  };

  const points: MachineStatusPoint[] = [];

  let m1c = 0;
  for (const r of s1rows) {
    if (m1c >= CHART_MAX_POINTS_PER_MACHINE) break;
    const dt = sengnr1RowTime(r.sdate, r.stime);
    if (!dt || dt < since) continue;
    m1c += 1;
    points.push({
      t: dt.toISOString(),
      value: sengnr1EventToGdata(r.event),
      machine: 1,
    });
  }

  const k2asc = [...k2rows]
    .map((r) => ({ r, dt: parseKyouj2Rectime(r.rectime) }))
    .filter((x): x is { r: (typeof k2rows)[0]; dt: Date } => x.dt != null)
    .filter((x) => x.dt >= since)
    .sort((a, b) => a.dt.getTime() - b.dt.getTime());
  let m2c = 0;
  for (let i = 0; i < k2asc.length; i++) {
    if (m2c >= CHART_MAX_POINTS_PER_MACHINE) break;
    const prev = i > 0 ? k2asc[i - 1]!.dt : null;
    const { g } = m2GdataFromRowAndPrev(k2asc[i]!.dt, prev);
    m2c += 1;
    points.push({ t: k2asc[i]!.dt.toISOString(), value: g, machine: 2 });
  }

  let m3c = 0;
  for (const r of s3rows) {
    if (m3c >= CHART_MAX_POINTS_PER_MACHINE) break;
    const dt = sengnr3RowTime(r.sdate, r.sekitime);
    if (!dt || dt < since) continue;
    m3c += 1;
    const dec = decodeSengnr3Event(r.event);
    points.push({ t: dt.toISOString(), value: dec.gValue, machine: 3 });
  }

  points.sort((a, b) => a.t.localeCompare(b.t));

  return {
    updatedAt: now.toISOString(),
    machines: [machine1, machine2, machine3],
    chart: { points },
  };
}
