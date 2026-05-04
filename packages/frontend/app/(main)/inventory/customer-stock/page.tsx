"use client";

import type { CellContext, ColumnDef } from "@tanstack/react-table";
import { useCallback, useMemo, useState } from "react";

import { DataTable } from "@/components/shared/data-table";
import { ExcelExportButton } from "@/components/shared/excel-export-button";
import { Button } from "@/components/ui/button";
import { apiFetch } from "@/lib/api";

type CustomerStockRow = {
  no: number;
  uno: string;
  nyukabi: string | null;
  sehncd: string;
  productName: string | null;
  kainame: string | null;
  pass: string | null;
  syouso: string | null;
  nyukasu: number | null;
  computedStock: number | null;
};

type CustomerStockPayload = {
  asOf: string;
  basisDate: string;
  kaisyacd: string;
  companyName: string | null;
  note: string;
  rows: CustomerStockRow[];
  totals: {
    rowCount: number;
    nyukasu: number;
    computedStock: number;
  };
  shipmentAfterAsOfSum: number;
};

function isoDateToYyyymmdd(iso: string): string {
  return iso.replaceAll("-", "");
}

export default function InventoryCustomerStockPage() {
  const [kaisyacd, setKaisyacd] = useState("");
  const [date, setDate] = useState(() => {
    const t = new Date();
    return `${t.getFullYear()}-${String(t.getMonth() + 1).padStart(2, "0")}-${String(t.getDate()).padStart(2, "0")}`;
  });
  const [data, setData] = useState<CustomerStockPayload | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const asOf = useMemo(() => isoDateToYyyymmdd(date), [date]);

  const load = useCallback(async () => {
    const cd = kaisyacd.trim();
    if (!cd) {
      setError("会社コードを入力してください");
      setData(null);
      return;
    }
    setLoading(true);
    setError(null);
    try {
      const res = await apiFetch<{ data: CustomerStockPayload }>(
        `/api/inventory/customer-stock?${new URLSearchParams({ kaisyacd: cd, asOf })}`,
      );
      setData(res.data);
    } catch (e) {
      setData(null);
      setError(e instanceof Error ? e.message : "取得に失敗しました");
    } finally {
      setLoading(false);
    }
  }, [kaisyacd, asOf]);

  const tableRows: Record<string, unknown>[] = useMemo(() => {
    if (!data) return [];
    const body: Record<string, unknown>[] = data.rows.map((r) => ({
      No: r.no,
      受付番号: r.uno,
      入荷日: r.nyukabi ?? "—",
      製品コード: r.sehncd,
      製品名: r.productName ?? "—",
      得意先名: r.kainame ?? "—",
      号機: r.syouso ?? "—",
      入荷数: r.nyukasu ?? 0,
      預り在庫目安: r.computedStock ?? 0,
    }));
    body.push({
      No: "—",
      受付番号: "—",
      入荷日: "—",
      製品コード: "—",
      製品名: "—",
      得意先名: "【合計】",
      号機: "—",
      入荷数: data.totals.nyukasu,
      預り在庫目安: data.totals.computedStock,
    });
    return body;
  }, [data]);

  const columns: ColumnDef<Record<string, unknown>>[] = useMemo(
    () =>
      ["No", "受付番号", "入荷日", "製品コード", "製品名", "得意先名", "号機", "入荷数", "預り在庫目安"].map(
        (k) => ({
          accessorKey: k,
          header: k,
          cell: (info: CellContext<Record<string, unknown>, unknown>) => {
            const v = info.getValue();
            if (v === null || v === undefined) return "—";
            return String(v);
          },
        }),
      ),
    [],
  );

  const exportName = useMemo(() => {
    if (!data) return "顧客在庫報告";
    return `顧客在庫報告_${data.kaisyacd}_${data.basisDate}`;
  }, [data]);

  return (
    <div className="mx-auto max-w-6xl space-y-6 p-4 md:p-6">
      <div>
        <h1 className="text-2xl font-semibold text-ink">顧客在庫報告</h1>
        <p className="mt-1 text-sm text-ink-muted">
          会社コードと在庫基準日を指定し、基準日までに入荷した在庫（zaiko）を一覧します。Excel
          原票の履歴・出荷加算は DB スキーマ拡張後に追補可能です。
        </p>
      </div>

      <div className="flex flex-wrap items-end gap-4">
        <label className="flex flex-col gap-1 text-sm text-ink-muted">
          会社コード
          <input
            type="text"
            value={kaisyacd}
            onChange={(e) => setKaisyacd(e.target.value)}
            placeholder="例: 0226"
            className="w-40 rounded-xl border border-cream-300 bg-white px-3 py-2 text-ink"
          />
        </label>
        <label className="flex flex-col gap-1 text-sm text-ink-muted">
          在庫基準日
          <input
            type="date"
            value={date}
            onChange={(e) => setDate(e.target.value)}
            className="rounded-xl border border-cream-300 bg-white px-3 py-2 text-ink"
          />
        </label>
        <Button type="button" onClick={() => void load()}>
          集計
        </Button>
      </div>

      {error && <p className="text-sm text-red-700">{error}</p>}

      {data && (
        <div className="space-y-2 rounded-2xl border border-cream-300 bg-cream-50 p-4 text-sm text-ink">
          <p>
            <span className="text-ink-muted">顧客名: </span>
            {data.companyName ?? "—"}
            <span className="ml-4 text-ink-muted">会社コード: </span>
            {data.kaisyacd}
            <span className="ml-4 text-ink-muted">基準日: </span>
            {data.basisDate.replace(/(\d{4})(\d{2})(\d{2})/, "$1/$2/$3")}
          </p>
          <p>
            <span className="text-ink-muted">基準日以降の出荷数合計（会社単位・参考）: </span>
            {data.shipmentAfterAsOfSum}
          </p>
          <p className="text-xs text-ink-muted">{data.note}</p>
        </div>
      )}

      <div className="flex flex-wrap items-center justify-between gap-3">
        <p className="text-xs text-ink-muted">
          {data && <>照会時刻: {new Date(data.asOf).toLocaleString("ja-JP")}</>}
        </p>
        <ExcelExportButton
          data={tableRows as Record<string, unknown>[]}
          fileName={exportName}
          sheetName="在庫"
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
