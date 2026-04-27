"use client";

import type { ColumnDef } from "@tanstack/react-table";
import { useCallback, useEffect, useMemo, useState } from "react";

import { DataTable } from "@/components/shared/data-table";
import { ExcelExportButton } from "@/components/shared/excel-export-button";
import { apiFetch } from "@/lib/api";
import { Button } from "@/components/ui/button";

type InvBucket = "1" | "2" | "3" | "EB" | "LOCA" | "other";

type MonthEndLine = {
  uno: string;
  kaisyacd: string;
  sehncd: string;
  syouso: string | null;
  kainame: string | null;
  nyukabi: string | null;
  nyukasu: number | null;
  pass: string | null;
  incnt: number | null;
  nouki: string | null;
  source: "zaiko" | "zaikor";
  bucket: InvBucket;
  bucketLabel: string;
};

type MonthEndSummaryRow = {
  bucket: InvBucket;
  label: string;
  rowCount: number;
  stockProxyNyukasu: number;
  wipAmount: number;
};

type MonthEndPayload = {
  asOf: string;
  closingDate: string;
  yearMonth: string;
  generatedAt: string;
  mvpNote: string;
  lines: MonthEndLine[];
  summary: MonthEndSummaryRow[];
  totals: { rowCount: number; totalNyukasu: number };
};

function isoDateToYyyymmdd(iso: string): string {
  return iso.replaceAll("-", "");
}

function yyyymmddToIso(ymd: string): string {
  if (ymd.length !== 8) return "";
  return `${ymd.slice(0, 4)}-${ymd.slice(4, 6)}-${ymd.slice(6, 8)}`;
}

export default function InventoryMonthEndPage() {
  const [date, setDate] = useState(() => {
    const t = new Date();
    return `${t.getFullYear()}-${String(t.getMonth() + 1).padStart(2, "0")}-${String(t.getDate()).padStart(2, "0")}`;
  });
  const [data, setData] = useState<MonthEndPayload | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const asOfYmd = useMemo(() => isoDateToYyyymmdd(date), [date]);

  const load = useCallback(async () => {
    setLoading(true);
    setError(null);
    try {
      const res = await apiFetch<{ data: MonthEndPayload }>(
        `/api/inventory/month-end?${new URLSearchParams({ asOf: asOfYmd })}`,
      );
      setData(res.data);
    } catch (e) {
      setData(null);
      setError(e instanceof Error ? e.message : "取得に失敗しました");
    } finally {
      setLoading(false);
    }
  }, [asOfYmd]);

  useEffect(() => {
    void load();
  }, [load]);

  const tableRows: Record<string, unknown>[] = useMemo(() => {
    if (!data) return [];
    const detail = data.lines.map((l) => ({
      行種別: "明細",
      件数: "",
      装置区分: l.bucketLabel,
      受付番号: l.uno,
      会社コード: l.kaisyacd,
      得意先名: l.kainame ?? "",
      製品コード: l.sehncd,
      装置: l.syouso ?? "",
      入荷日: l.nyukabi ?? "",
      入荷数: l.nyukasu ?? 0,
      パス: l.pass ?? "",
      入数: l.incnt ?? "",
      納期: l.nouki ?? "",
      参照元: l.source,
    }));
    const summaryRows = data.summary.map((s) => ({
      行種別: "装置別集計",
      件数: s.rowCount,
      装置区分: s.label,
      受付番号: "",
      会社コード: "",
      得意先名: "",
      製品コード: "",
      装置: "",
      入荷日: "",
      入荷数: s.stockProxyNyukasu,
      パス: "",
      入数: "",
      納期: "",
      参照元: `仕掛(MVP)=${s.wipAmount}`,
    }));
    return [...detail, ...summaryRows];
  }, [data]);

  const columns: ColumnDef<Record<string, unknown>>[] = useMemo(
    () =>
      [
        "行種別",
        "件数",
        "装置区分",
        "受付番号",
        "会社コード",
        "得意先名",
        "製品コード",
        "装置",
        "入荷日",
        "入荷数",
        "パス",
        "入数",
        "納期",
        "参照元",
      ].map((k) => ({
        accessorKey: k,
        header: k,
        cell: (info) => {
          const v = info.getValue();
          if (v === null || v === undefined) return "—";
          return String(v);
        },
      })),
    [],
  );

  const exportName = useMemo(() => {
    if (!data) return "月末在庫集計";
    return `月末在庫集計_${data.asOf}`;
  }, [data]);

  return (
    <div className="mx-auto max-w-6xl space-y-6 p-4 md:p-6">
      <div>
        <h1 className="text-2xl font-semibold text-ink">月末在庫集計</h1>
        <p className="mt-1 text-sm text-ink-muted">
          締日を指定して zaiko / zaikor から月末スナップショット風の一覧・装置別集計を表示します（MVP・簡易SQL 相当）。
        </p>
      </div>

      <div className="flex flex-wrap items-end gap-4">
        <label className="flex flex-col gap-1 text-sm text-ink-muted">
          締日（asOf）
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

      {data && (
        <p className="text-xs text-ink-muted leading-relaxed">
          {data.mvpNote}
        </p>
      )}

      {error && <p className="text-sm text-red-700">{error}</p>}

      <div className="flex flex-wrap items-center justify-between gap-3">
        <p className="text-xs text-ink-muted">
          {data && (
            <>
              照会時刻: {new Date(data.generatedAt).toLocaleString("ja-JP")}（締日:{" "}
              {yyyymmddToIso(data.asOf).replaceAll("-", "/")} / 合計 {data.totals.rowCount} 行 / 入荷数合計{" "}
              {data.totals.totalNyukasu}）
            </>
          )}
        </p>
        <ExcelExportButton data={tableRows} fileName={exportName} sheetName="月末在庫">
          Excel出力
        </ExcelExportButton>
      </div>

      <DataTable columns={columns} data={tableRows} loading={loading} pageSize={20} globalFilter />
    </div>
  );
}
