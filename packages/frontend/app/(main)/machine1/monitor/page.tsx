"use client";

import { useQuery } from "@tanstack/react-query";
import type { ColumnDef } from "@tanstack/react-table";

import { DataTable } from "@/components/shared/data-table";
import { MachineStatusBadge, type MachineStatus } from "@/components/shared/machine-status-badge";
import { apiFetch } from "@/lib/api";

type MonitorRow = {
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

type Machine1MonitorResponse = {
  data: {
    updatedAt: string;
    sourceState: {
      label: string;
      sdate?: string;
      stime?: string;
      event?: string;
      timer?: string;
    };
    rows: MonitorRow[];
  };
};

/**
 * SENGNR1.EVENT: 0 昇降中 / 1 照射中 / 2 貯蔵中 / 3 移動照射中（Ex1号機照射情報 §2.3）
 * MachineStatusBadge の4状態に寄せる: 1=running, 2=stopped, その他=unknown
 */
function eventToMachineStatus(event: string | number | undefined): MachineStatus {
  const e = String(event ?? "").trim();
  if (e === "1") return "running";
  if (e === "2") return "stopped";
  return "unknown";
}

const columns: ColumnDef<MonitorRow>[] = [
  { accessorKey: "no", header: "No" },
  { accessorKey: "receiptNo", header: "受付番号" },
  { accessorKey: "company", header: "会社名" },
  { accessorKey: "dose", header: "指定線量" },
  { accessorKey: "qty", header: "数量" },
  { accessorKey: "position", header: "照射位置" },
  { accessorKey: "remaining", header: "完了までの時間" },
  { accessorKey: "eta", header: "完了予想日時" },
  { accessorKey: "shipDate", header: "出荷日" },
  { accessorKey: "note", header: "備考" },
];

export default function IrradiationMachine1MonitorPage() {
  const { data, isLoading, isError, error } = useQuery({
    queryKey: ["irradiation", "machine1", "monitor"],
    queryFn: () => apiFetch<Machine1MonitorResponse>("/api/irradiation/machine1/monitor"),
    refetchInterval: 30_000,
  });

  const monitor = data?.data;
  const src = monitor?.sourceState;

  return (
    <div className="flex flex-col gap-6 p-4 md:p-6">
      <div>
        <h1 className="text-xl font-semibold text-ink">1号機固定照射情報</h1>
        <p className="mt-1 text-sm text-ink-muted">
          直近30秒で自動更新（手動再取得は次回ポーリングまで待機）
        </p>
      </div>

      {src && (
        <div className="flex flex-col gap-3 rounded-xl border border-cream-300 bg-cream-50/80 p-4">
          <div className="flex flex-wrap items-center gap-3">
            <span className="text-sm font-medium text-ink">線源状態</span>
            <MachineStatusBadge status={eventToMachineStatus(src.event)} machineName="1号機" />
            <span className="text-sm text-ink">{src.label}</span>
          </div>
          <div className="grid gap-1 text-sm text-ink-muted sm:grid-cols-2">
            {src.sdate && (
              <p>
                <span className="text-ink-muted">日付: </span>
                {src.sdate}
              </p>
            )}
            {src.stime && (
              <p>
                <span className="text-ink-muted">時刻: </span>
                {src.stime}
              </p>
            )}
            {src.timer && (
              <p>
                <span className="text-ink-muted">タイマー: </span>
                {src.timer}
              </p>
            )}
            {monitor?.updatedAt && (
              <p>
                <span className="text-ink-muted">更新: </span>
                {new Date(monitor.updatedAt).toLocaleString("ja-JP", { timeZone: "Asia/Tokyo" })}
              </p>
            )}
          </div>
        </div>
      )}

      {isError && (
        <p className="rounded-xl border border-red-200 bg-red-50 px-3 py-2 text-sm text-red-800">
          {error instanceof Error ? error.message : "取得に失敗しました"}
        </p>
      )}

      <DataTable columns={columns} data={monitor?.rows ?? []} loading={isLoading} pageSize={20} />
    </div>
  );
}
