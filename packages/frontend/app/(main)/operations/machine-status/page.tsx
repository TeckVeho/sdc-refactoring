"use client";

import { useMemo } from "react";
import { useQuery } from "@tanstack/react-query";

import { MachineStatusBadge, type MachineStatus } from "@/components/shared/machine-status-badge";
import { TrendChart } from "@/components/shared/trend-chart";
import { apiFetch } from "@/lib/api";

type MachineStatusId = 1 | 2 | 3;

type StatusPoint = { t: string; value: number; machine: MachineStatusId };

type MachineRow = {
  id: MachineStatusId;
  timer?: string;
  statusKey: string;
  statusLabel: string;
  detail?: string;
};

type MachineStatusResponse = {
  data: {
    updatedAt: string;
    machines: MachineRow[];
    chart: { points: StatusPoint[] };
  };
};

function rowToBadgeStatus(m: MachineRow): MachineStatus {
  if (m.id === 1) {
    const e = m.statusKey;
    if (e === "1") return "running";
    if (e === "0" || e === "2") return "stopped";
    return "unknown";
  }
  if (m.id === 2) {
    if (m.statusKey === "active") return "running";
    if (m.statusKey === "idle") return "stopped";
    return "unknown";
  }
  if (m.id === 3) {
    if (m.statusKey === "irradiation_like") return "running";
    if (m.statusKey === "idle_like") return "stopped";
    return "unknown";
  }
  return "unknown";
}

function nameForId(id: MachineStatusId): string {
  if (id === 1) return "1号機";
  if (id === 2) return "2号機";
  return "3号機";
}

function pivotChartPoints(points: StatusPoint[]) {
  const byIso = new Map<
    string,
    { time: string; 号機1?: number; 号機2?: number; 号機3?: number }
  >();
  for (const p of points) {
    if (!byIso.has(p.t)) {
      byIso.set(p.t, {
        time: new Date(p.t).toLocaleString("ja-JP", {
          timeZone: "Asia/Tokyo",
          month: "numeric",
          day: "numeric",
          hour: "2-digit",
          minute: "2-digit",
        }),
      });
    }
    const r = byIso.get(p.t)!;
    if (p.machine === 1) r.号機1 = p.value;
    if (p.machine === 2) r.号機2 = p.value;
    if (p.machine === 3) r.号機3 = p.value;
  }
  return Array.from(byIso.entries())
    .sort((a, b) => a[0].localeCompare(b[0]))
    .map(([, v]) => v);
}

const CHART_KEYS = ["号機1", "号機2", "号機3"] as const;

export default function IrradiationMachineStatusPage() {
  const { data, isLoading, isError, error } = useQuery({
    queryKey: ["irradiation", "machine-status"],
    queryFn: () => apiFetch<MachineStatusResponse>("/api/irradiation/machine-status"),
    refetchInterval: 60_000,
  });

  const body = data?.data;
  const machines = useMemo(
    () =>
      body?.machines ?? [
        { id: 1 as const, statusKey: "", statusLabel: "…" },
        { id: 2 as const, statusKey: "", statusLabel: "…" },
        { id: 3 as const, statusKey: "", statusLabel: "…" },
      ],
    [body?.machines],
  );
  const chartData = useMemo(
    () => (body?.chart.points.length ? pivotChartPoints(body.chart.points) : []),
    [body?.chart.points],
  );

  return (
    <div className="flex flex-col gap-6 p-4 md:p-6">
      <div>
        <h1 className="text-xl font-semibold text-ink">装置運転状況</h1>
        <p className="mt-1 text-sm text-ink-muted">60 秒間隔で自動更新（G_Data 風: 1号 0/0.5/1, 2号 1.2/2.2, 3号 2.4/3.4）</p>
        {body?.updatedAt && (
          <p className="mt-0.5 text-xs text-ink-muted">
            最終取得: {new Date(body.updatedAt).toLocaleString("ja-JP", { timeZone: "Asia/Tokyo" })}
          </p>
        )}
      </div>

      {isError && (
        <p className="rounded-xl border border-red-200 bg-red-50 px-3 py-2 text-sm text-red-800">
          {error instanceof Error ? error.message : "取得に失敗しました"}
        </p>
      )}

      {!isError && (
        <div className="flex flex-col gap-3 rounded-xl border border-cream-300 bg-cream-50/80 p-4">
          <p className="text-sm font-medium text-ink">各号機ステータス</p>
          <div className="flex flex-wrap gap-2">
            {machines.map((m) => (
              <MachineStatusBadge
                key={m.id}
                status={!body || isLoading ? "unknown" : rowToBadgeStatus(m)}
                machineName={nameForId(m.id)}
              />
            ))}
          </div>
          {body && (
            <ul className="space-y-2 text-sm text-ink-muted">
              {body.machines.map((m) => (
                <li key={m.id}>
                  <span className="text-ink">{nameForId(m.id)}</span> — {m.statusLabel}
                  {m.timer != null && m.timer !== "" && (
                    <span className="ml-1 text-ink-muted">（タイマー/指標: {m.timer}）</span>
                  )}
                  {m.detail && (
                    <span className="mt-0.5 block text-xs text-ink-muted/90">{m.detail}</span>
                  )}
                </li>
              ))}
            </ul>
          )}
        </div>
      )}

      <div className="rounded-xl border border-cream-300 bg-cream-50/80 p-4">
        <h2 className="mb-2 text-sm font-medium text-ink">直近 24 時間トレンド</h2>
        {isLoading && !body ? (
          <p className="text-sm text-ink-muted">読み込み中…</p>
        ) : chartData.length === 0 ? (
          <p className="text-sm text-ink-muted">該当データがありません</p>
        ) : (
          <TrendChart
            data={chartData}
            dataKeys={[...CHART_KEYS]}
            xAxisKey="time"
            height={320}
          />
        )}
      </div>
    </div>
  );
}
