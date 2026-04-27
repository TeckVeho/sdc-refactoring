"use client";

import type { ColumnDef } from "@tanstack/react-table";
import { useCallback, useEffect, useMemo, useState } from "react";

import { DataTable } from "@/components/shared/data-table";
import { ExcelExportButton } from "@/components/shared/excel-export-button";
import { apiFetch } from "@/lib/api";
import { cn } from "@/lib/utils";
import { Button } from "@/components/ui/button";

type View = "arrival" | "shipment" | "uptime";

type ArrivalSeries = {
  machines: Array<{
    machine: "1" | "2" | "3";
    label: string;
    rows: Array<{ kainame: string; nyukasu: number }>;
    totalNyukasu: number;
  }>;
};

type ShipmentSeries = {
  rows: Array<{ kaisyacd: string; companyName: string; syukasu: number }>;
  totalSyukasu: number;
};

type UptimeSeries = {
  periodDays: number;
  sengnr1: { count: number; timerSum: number; last: { sdate: string; stime: string | null } | null };
  kyouj2: { count: number; lastRectime: string | null };
  sengnr3: { count: number; last: { sdate: string; sekitime: string | null } | null };
  simpleUptimeRate: number;
};

type ShipmentSummaryPayload =
  | { view: "arrival"; asOf: string; date: string; series: ArrivalSeries }
  | { view: "shipment"; asOf: string; date: string; series: ShipmentSeries }
  | { view: "uptime"; asOf: string; from: string; to: string; series: UptimeSeries };

function isoDateToYyyymmdd(iso: string): string {
  return iso.replaceAll("-", "");
}

function yyyymmddToIso(ymd: string): string {
  if (ymd.length !== 8) return "";
  return `${ymd.slice(0, 4)}-${ymd.slice(4, 6)}-${ymd.slice(6, 8)}`;
}

export default function InventoryShipmentSummaryPage() {
  const [view, setView] = useState<View>("arrival");
  const [date, setDate] = useState(() => {
    const t = new Date();
    return `${t.getFullYear()}-${String(t.getMonth() + 1).padStart(2, "0")}-${String(t.getDate()).padStart(2, "0")}`;
  });
  const [uptimeFrom, setUptimeFrom] = useState<string | null>(null);
  const [uptimeTo, setUptimeTo] = useState<string | null>(null);
  const [data, setData] = useState<ShipmentSummaryPayload | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const ymd = useMemo(() => isoDateToYyyymmdd(date), [date]);
  const uptimeFromYmd = uptimeFrom ? isoDateToYyyymmdd(uptimeFrom) : null;
  const uptimeToYmd = uptimeTo ? isoDateToYyyymmdd(uptimeTo) : null;

  const load = useCallback(async () => {
    setLoading(true);
    setError(null);
    try {
      if (view === "uptime") {
        const from = uptimeFromYmd ?? ymd;
        const to = uptimeToYmd ?? ymd;
        const res = await apiFetch<{ data: ShipmentSummaryPayload }>(
          `/api/inventory/shipment-summary?${new URLSearchParams({ view, from, to })}`,
        );
        setData(res.data);
      } else {
        const res = await apiFetch<{ data: ShipmentSummaryPayload }>(
          `/api/inventory/shipment-summary?${new URLSearchParams({ view, date: ymd })}`,
        );
        setData(res.data);
      }
    } catch (e) {
      setData(null);
      setError(e instanceof Error ? e.message : "取得に失敗しました");
    } finally {
      setLoading(false);
    }
  }, [view, ymd, uptimeFromYmd, uptimeToYmd]);

  useEffect(() => {
    void load();
  }, [load]);

  const tableRows: Record<string, unknown>[] = useMemo(() => {
    if (!data) return [];
    if (data.view === "arrival") {
      const s = data.series;
      const out: Record<string, unknown>[] = [];
      for (const m of s.machines) {
        for (const r of m.rows) {
          out.push({ 号機: m.label, 会社名: r.kainame, 入荷数: r.nyukasu });
        }
        out.push({ 号機: m.label, 会社名: "【合計】", 入荷数: m.totalNyukasu });
      }
      return out;
    }
    if (data.view === "shipment") {
      const s = data.series;
      const out = s.rows.map((r) => ({
        会社コード: r.kaisyacd,
        社名: r.companyName,
        出荷数: r.syukasu,
      }));
      out.push({ 会社コード: "", 社名: "合計", 出荷数: s.totalSyukasu });
      return out;
    }
    const s = data.series;
    return [
      { 項目: "期間日数", 値: s.periodDays },
      { 項目: "1号機(sengnr1) 記録件数", 値: s.sengnr1.count },
      { 項目: "1号機 timer 合計", 値: s.sengnr1.timerSum },
      { 項目: "1号機 最終 sdate+stime", 値: s.sengnr1.last ? `${s.sengnr1.last.sdate} ${s.sengnr1.last.stime ?? ""}`.trim() : "—" },
      { 項目: "2号機(kyouj2) 記録件数", 値: s.kyouj2.count },
      { 項目: "2号機 最終 rectime", 値: s.kyouj2.lastRectime ?? "—" },
      { 項目: "3号機(sengnr3) 記録件数", 値: s.sengnr3.count },
      { 項目: "3号機 最終行", 値: s.sengnr3.last ? `${s.sengnr3.last.sdate} ${s.sengnr3.last.sekitime ?? ""}`.trim() : "—" },
      { 項目: "簡易稼働率（API定義）", 値: s.simpleUptimeRate.toFixed(4) },
    ];
  }, [data]);

  const rowKeys = useMemo(() => {
    if (tableRows[0]) return Object.keys(tableRows[0]);
    if (view === "uptime") return ["項目", "値"];
    if (view === "shipment") return ["会社コード", "社名", "出荷数"];
    return ["号機", "会社名", "入荷数"];
  }, [tableRows, view]);

  const columns: ColumnDef<Record<string, unknown>>[] = useMemo(
    () =>
      rowKeys.map((k) => ({
        accessorKey: k,
        header: k,
        cell: (info) => {
          const v = info.getValue();
          if (v === null || v === undefined) return "—";
          return String(v);
        },
      })),
    [rowKeys],
  );

  const exportName = useMemo(() => {
    if (!data) return "入出荷集計";
    if (data.view === "uptime")
      return `入出荷集計_稼働_${data.from}_${data.to}`;
    return `入出荷集計_${data.view === "arrival" ? "入荷" : "出荷"}_${data.date}`;
  }, [data]);

  return (
    <div className="mx-auto max-w-5xl space-y-6 p-4 md:p-6">
      <div>
        <h1 className="text-2xl font-semibold text-ink">入出荷集計</h1>
        <p className="mt-1 text-sm text-ink-muted">
          入荷状況（Zaiko・装置別）・出荷実績（Syukar）・装置稼働指標（概算）を表示します。
        </p>
      </div>

      <div className="flex flex-wrap items-center gap-2 rounded-2xl border border-cream-300 bg-cream-50 p-1">
        {(
          [
            ["arrival", "入荷状況"],
            ["shipment", "出荷実績"],
            ["uptime", "稼働時間"],
          ] as const
        ).map(([v, label]) => (
          <Button
            key={v}
            type="button"
            variant="ghost"
            className={cn(
              "rounded-xl px-4 py-2 text-sm",
              view === v
                ? "bg-white text-ink shadow-sm ring-1 ring-cream-300"
                : "text-ink-muted hover:text-ink",
            )}
            onClick={() => setView(v)}
          >
            {label}
          </Button>
        ))}
      </div>

      {view === "uptime" ? (
        <div className="flex flex-wrap items-end gap-4">
          <label className="flex flex-col gap-1 text-sm text-ink-muted">
            期間（開始）
            <input
              type="date"
              value={uptimeFrom ?? date}
              onChange={(e) => setUptimeFrom(e.target.value)}
              className="rounded-xl border border-cream-300 bg-white px-3 py-2 text-ink"
            />
          </label>
          <label className="flex flex-col gap-1 text-sm text-ink-muted">
            期間（終了）
            <input
              type="date"
              value={uptimeTo ?? date}
              onChange={(e) => setUptimeTo(e.target.value)}
              className="rounded-xl border border-cream-300 bg-white px-3 py-2 text-ink"
            />
          </label>
          <Button type="button" onClick={() => void load()}>
            再読み込み
          </Button>
        </div>
      ) : (
        <div className="flex flex-wrap items-end gap-4">
          <label className="flex flex-col gap-1 text-sm text-ink-muted">
            基準日
            <input
              type="date"
              value={date}
              onChange={(e) => setDate(e.target.value)}
              className="rounded-xl border border-cream-300 bg-white px-3 py-2 text-ink"
            />
          </label>
          <Button type="button" onClick={() => void load()}>
            再読み込み
          </Button>
        </div>
      )}

      {error && <p className="text-sm text-red-700">{error}</p>}

      <div className="flex flex-wrap items-center justify-between gap-3">
        <p className="text-xs text-ink-muted">
          {data && (
            <>
              照会時刻: {new Date(data.asOf).toLocaleString("ja-JP")}{" "}
              {data.view === "uptime" ? `（期間: ${data.from}〜${data.to}）` : `（日付: ${yyyymmddToIso(data.date).replaceAll("-", "/")}）`}
            </>
          )}
        </p>
        <ExcelExportButton
          data={tableRows as Record<string, unknown>[]}
          fileName={exportName}
          sheetName="集計"
        >
          Excel出力
        </ExcelExportButton>
      </div>

      <DataTable
        columns={columns}
        data={tableRows as Record<string, unknown>[]}
        loading={loading}
        pageSize={20}
        globalFilter={view === "arrival" || view === "shipment"}
      />
    </div>
  );
}
