"use client";

import type { ColumnDef } from "@tanstack/react-table";
import { useCallback, useEffect, useMemo, useState } from "react";

import { DataTable } from "@/components/shared/data-table";
import { ExcelExportButton } from "@/components/shared/excel-export-button";
import { apiFetch } from "@/lib/api";
import { Button } from "@/components/ui/button";

type CompanyRow = { kaisyacd: string; companyName: string; nyukasu: number };
type ProductRow = {
  kaisyacd: string;
  sehncd: string;
  productName: string;
  nyukasu: number;
};

type CompanyArrivalByCompany = {
  mode: "byCompany";
  date: string;
  rows: CompanyRow[];
  totals?: { nyukasu: number; companyCount: number };
};

type CompanyArrivalByProduct = {
  mode: "byProduct";
  date: string;
  kaisyacd: string;
  companyName: string;
  rows: ProductRow[];
  totals?: { nyukasu: number; productCount: number };
};

type CompanyArrivalPayload = CompanyArrivalByCompany | CompanyArrivalByProduct;

function isoDateToYyyymmdd(iso: string): string {
  return iso.replaceAll("-", "");
}

function yyyymmddToSlash(ymd: string): string {
  if (ymd.length !== 8) return ymd;
  return `${ymd.slice(0, 4)}/${ymd.slice(4, 6)}/${ymd.slice(6, 8)}`;
}

export default function InventoryCompanyArrivalPage() {
  const [date, setDate] = useState(() => {
    const t = new Date();
    return `${t.getFullYear()}-${String(t.getMonth() + 1).padStart(2, "0")}-${String(t.getDate()).padStart(2, "0")}`;
  });
  const [kaisyacd, setKaisyacd] = useState("");
  const [payload, setPayload] = useState<CompanyArrivalPayload | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const ymd = useMemo(() => isoDateToYyyymmdd(date), [date]);

  const load = useCallback(async () => {
    setLoading(true);
    setError(null);
    try {
      const q = new URLSearchParams({ date: ymd });
      const trimmed = kaisyacd.trim();
      if (trimmed) q.set("kaisyacd", trimmed);
      const res = await apiFetch<{ data: CompanyArrivalPayload }>(
        `/api/inventory/company-arrival?${q}`,
      );
      setPayload(res.data);
    } catch (e) {
      setPayload(null);
      setError(e instanceof Error ? e.message : "取得に失敗しました");
    } finally {
      setLoading(false);
    }
  }, [ymd, kaisyacd]);

  useEffect(() => {
    void load();
  }, [load]);

  const tableRows: Record<string, unknown>[] = useMemo(() => {
    if (!payload) return [];
    if (payload.mode === "byCompany") {
      return payload.rows.map((r) => ({
        会社コード: r.kaisyacd,
        社名: r.companyName,
        入荷数: r.nyukasu,
      }));
    }
    return payload.rows.map((r) => ({
      会社コード: r.kaisyacd,
      製品コード: r.sehncd,
      製品名: r.productName,
      入荷数: r.nyukasu,
    }));
  }, [payload]);

  const columns: ColumnDef<Record<string, unknown>>[] = useMemo(() => {
    if (!payload) return [];
    if (payload.mode === "byCompany") {
      return (["会社コード", "社名", "入荷数"] as const).map((k) => ({
        accessorKey: k,
        header: k,
        cell: (info) => {
          const v = info.getValue();
          if (v === null || v === undefined) return "—";
          return String(v);
        },
      }));
    }
    return (["会社コード", "製品コード", "製品名", "入荷数"] as const).map((k) => ({
      accessorKey: k,
      header: k,
      cell: (info) => {
        const v = info.getValue();
        if (v === null || v === undefined) return "—";
        return String(v);
      },
    }));
  }, [payload]);

  const exportName = useMemo(() => {
    if (!payload) return "会社別入荷集計";
    const suffix = payload.mode === "byProduct" ? `_${payload.kaisyacd}` : "";
    return `会社別入荷集計_${payload.date}${suffix}`;
  }, [payload]);

  const summaryLine = useMemo(() => {
    if (!payload?.totals) return null;
    if (payload.mode === "byCompany") {
      return `全社合計: ${payload.totals.nyukasu}（会社数 ${payload.totals.companyCount}）`;
    }
    return `合計: ${payload.totals.nyukasu}（製品行 ${payload.totals.productCount}）`;
  }, [payload]);

  return (
    <div className="mx-auto max-w-5xl space-y-6 p-4 md:p-6">
      <div>
        <h1 className="text-2xl font-semibold text-ink">会社別入荷集計</h1>
        <p className="mt-1 text-sm text-ink-muted">
          履歴在庫（Zaikor）から指定日の入荷数を集計します。会社コードを空にすると会社別、指定すると当該会社の製品別です。
        </p>
      </div>

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
        <label className="flex flex-col gap-1 text-sm text-ink-muted">
          会社コード（任意）
          <input
            type="text"
            value={kaisyacd}
            onChange={(e) => setKaisyacd(e.target.value)}
            placeholder="例: 0001（空=全社・会社別）"
            className="w-56 rounded-xl border border-cream-300 bg-white px-3 py-2 text-ink"
          />
        </label>
        <Button type="button" onClick={() => void load()}>
          再読み込み
        </Button>
      </div>

      {error && <p className="text-sm text-red-700">{error}</p>}

      <div className="flex flex-wrap items-center justify-between gap-3">
        <div className="text-xs text-ink-muted">
          {payload && (
            <p>
              集計日: {yyyymmddToSlash(payload.date)}
              {payload.mode === "byProduct" && (
                <>
                  {" "}
                  · {payload.companyName}（{payload.kaisyacd}）
                </>
              )}
            </p>
          )}
          {summaryLine && <p className="mt-1">{summaryLine}</p>}
        </div>
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
        globalFilter
      />
    </div>
  );
}
