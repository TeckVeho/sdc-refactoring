"use client";

import { useQuery } from "@tanstack/react-query";
import type { ColumnDef } from "@tanstack/react-table";
import { useCallback, useMemo, useState } from "react";

import { DataTable } from "@/components/shared/data-table";
import { Button } from "@/components/ui/button";
import { apiFetch } from "@/lib/api";

type SyousoFilter = "" | "1" | "2" | "3" | "EB";

type ProductionListRow = {
  uno: string;
  kaisyacd: string;
  sehncd: string;
  syouso: string | null;
  kainame: string | null;
  nyukabi: string | null;
  nouki: string | null;
  pass: string | null;
  nyukasu: number | null;
  incnt: number | null;
  siteisn: string | null;
  misyousu: number | null;
  syosuu: number | null;
  senritu: string | null;
  syokind: string | null;
  syostat: string | null;
  syukkabi: string | null;
  bikou1: string | null;
  kakunin: string | null;
  syuhouhou: string | null;
  yoyakubi: string | null;
  yoyakuno: string | null;
  yoyakuBikou: string | null;
};

type ProductionListResponse = {
  data: {
    rows: ProductionListRow[];
    meta: { total: number };
  };
};

const DEFAULT_LIMIT = 200;

function cellString(v: unknown): string {
  if (v === null || v === undefined) return "—";
  if (typeof v === "number" && !Number.isFinite(v)) return "—";
  return String(v);
}

const columns: ColumnDef<ProductionListRow>[] = [
  { accessorKey: "uno", header: "受付番号", cell: (ctx) => cellString(ctx.getValue()) },
  { accessorKey: "kainame", header: "会社名", cell: (ctx) => cellString(ctx.getValue()) },
  { accessorKey: "kaisyacd", header: "会社コード", cell: (ctx) => cellString(ctx.getValue()) },
  { accessorKey: "sehncd", header: "製品コード", cell: (ctx) => cellString(ctx.getValue()) },
  { accessorKey: "syouso", header: "装置", cell: (ctx) => cellString(ctx.getValue()) },
  { accessorKey: "nyukabi", header: "入荷日", cell: (ctx) => cellString(ctx.getValue()) },
  { accessorKey: "nouki", header: "納期", cell: (ctx) => cellString(ctx.getValue()) },
  { accessorKey: "pass", header: "パス", cell: (ctx) => cellString(ctx.getValue()) },
  { accessorKey: "incnt", header: "入数", cell: (ctx) => cellString(ctx.getValue()) },
  { accessorKey: "nyukasu", header: "入荷数", cell: (ctx) => cellString(ctx.getValue()) },
  { accessorKey: "siteisn", header: "指定線量", cell: (ctx) => cellString(ctx.getValue()) },
  { accessorKey: "syosuu", header: "照射数量", cell: (ctx) => cellString(ctx.getValue()) },
  { accessorKey: "misyousu", header: "未照射数", cell: (ctx) => cellString(ctx.getValue()) },
  { accessorKey: "senritu", header: "線量率", cell: (ctx) => cellString(ctx.getValue()) },
  { accessorKey: "syokind", header: "照射種別", cell: (ctx) => cellString(ctx.getValue()) },
  { accessorKey: "syostat", header: "照射状態", cell: (ctx) => cellString(ctx.getValue()) },
  { accessorKey: "syukkabi", header: "出荷日(計画)", cell: (ctx) => cellString(ctx.getValue()) },
  { accessorKey: "syuhouhou", header: "出荷方法", cell: (ctx) => cellString(ctx.getValue()) },
  { accessorKey: "kakunin", header: "確認", cell: (ctx) => cellString(ctx.getValue()) },
  { accessorKey: "bikou1", header: "備考1(計画)", cell: (ctx) => cellString(ctx.getValue()) },
  { accessorKey: "yoyakubi", header: "予約日", cell: (ctx) => cellString(ctx.getValue()) },
  { accessorKey: "yoyakuno", header: "予約番号", cell: (ctx) => cellString(ctx.getValue()) },
  { accessorKey: "yoyakuBikou", header: "予約備考", cell: (ctx) => cellString(ctx.getValue()) },
];

export default function IrradiationProductionPage() {
  const [syouso, setSyouso] = useState<SyousoFilter>("");
  const [limit, setLimit] = useState(DEFAULT_LIMIT);

  const buildUrl = useCallback(() => {
    const p = new URLSearchParams();
    p.set("limit", String(limit));
    if (syouso) p.set("syouso", syouso);
    return `/api/irradiation/production?${p.toString()}`;
  }, [limit, syouso]);

  const { data, isLoading, isError, error, refetch, isFetching } = useQuery({
    queryKey: ["irradiation", "production", { syouso, limit }],
    queryFn: () => apiFetch<ProductionListResponse>(buildUrl()),
  });

  const rows = useMemo(() => data?.data.rows ?? [], [data]);
  const total = data?.data.meta?.total;

  return (
    <div className="flex flex-col gap-6 p-4 md:p-6">
      <div>
        <h1 className="text-xl font-semibold text-ink">生産情報一覧（MVP）</h1>
        <p className="mt-1 text-sm text-ink-muted">
          在庫（Zaiko）を主に、照射（Syouk1）と計画 JSON（ExKeikakuX）・予約 JSON（ExYoyakuX）を受付番号で付加しています。DB
          に無い列は API で省略されます。
        </p>
      </div>

      <div className="flex flex-wrap items-end gap-4">
        <label className="flex flex-col gap-1 text-sm text-ink-muted">
          装置
          <select
            value={syouso}
            onChange={(e) => setSyouso(e.target.value as SyousoFilter)}
            className="min-w-[10rem] rounded-xl border border-cream-300 bg-white px-3 py-2 text-ink"
          >
            <option value="">すべて</option>
            <option value="1">1号機</option>
            <option value="2">2号機</option>
            <option value="3">3号機</option>
            <option value="EB">EB</option>
          </select>
        </label>
        <label className="flex flex-col gap-1 text-sm text-ink-muted">
          件数上限
          <input
            type="number"
            min={1}
            max={2000}
            value={limit}
            onChange={(e) => setLimit(Number(e.target.value) || DEFAULT_LIMIT)}
            className="w-28 rounded-xl border border-cream-300 bg-white px-3 py-2 text-ink"
          />
        </label>
        <Button
          type="button"
          variant="default"
          onClick={() => void refetch()}
          disabled={isFetching}
        >
          再取得
        </Button>
        {total != null && (
          <p className="text-sm text-ink-muted">
            全 {total} 件中 {rows.length} 件表示
          </p>
        )}
      </div>

      {isError && (
        <p className="rounded-xl border border-red-200 bg-red-50 px-3 py-2 text-sm text-red-800">
          {error instanceof Error ? error.message : "取得に失敗しました"}
        </p>
      )}

      <DataTable columns={columns} data={rows} loading={isLoading} pageSize={20} />
    </div>
  );
}
