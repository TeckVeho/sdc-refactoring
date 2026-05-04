"use client";

import { useQuery } from "@tanstack/react-query";
import type { ColumnDef } from "@tanstack/react-table";
import { useCallback, useMemo, useState } from "react";

import { DataTable } from "@/components/shared/data-table";
import { Button } from "@/components/ui/button";
import { apiFetch } from "@/lib/api";

type IrradiationSource = "all" | "syouk1" | "syouk2" | "syoukj3";

type IrradiationResultRow = {
  source: "syouk1" | "syouk2" | "syoukj3";
  uno: string;
  syono: string | null;
  kainame: string | null;
  siteisn: string | null;
  syosuu: number | null;
  syoichi: string | null;
  syotime: string | null;
  hansuu: number | null;
  stimer: string | null;
  ktimer: string | null;
  senritu: string | null;
  syostat: string | null;
  ctimer: string | null;
  zhansuu: number | null;
  htimer: string | null;
  slotno: string | null;
  sdate: string | null;
  edate: string | null;
  updfflg: string | null;
  syokind: string | null;
  bikou: string | null;
};

type ListResponse = {
  data: {
    rows: IrradiationResultRow[];
    meta: {
      total: number;
      page: number;
      pageSize: number;
      totalPages: number;
    };
  };
};

const DEFAULT_PAGE_SIZE = 50;

function cellString(v: unknown): string {
  if (v === null || v === undefined) return "—";
  if (typeof v === "number" && !Number.isFinite(v)) return "—";
  return String(v);
}

function sourceLabel(s: string): string {
  if (s === "syouk1") return "照射1 (syouk1)";
  if (s === "syouk2") return "照射2 (syouk2)";
  if (s === "syoukj3") return "照射3 (syoukj3)";
  return s;
}

const columns: ColumnDef<IrradiationResultRow>[] = [
  {
    accessorKey: "source",
    header: "ソース",
    cell: (ctx) => sourceLabel(cellString(ctx.getValue())),
  },
  { accessorKey: "uno", header: "受付番号", cell: (ctx) => cellString(ctx.getValue()) },
  { accessorKey: "kainame", header: "会社名", cell: (ctx) => cellString(ctx.getValue()) },
  { accessorKey: "syono", header: "照射番号", cell: (ctx) => cellString(ctx.getValue()) },
  { accessorKey: "siteisn", header: "指定線量", cell: (ctx) => cellString(ctx.getValue()) },
  { accessorKey: "syosuu", header: "照射数量", cell: (ctx) => cellString(ctx.getValue()) },
  { accessorKey: "syoichi", header: "照射位置", cell: (ctx) => cellString(ctx.getValue()) },
  { accessorKey: "syotime", header: "照射時刻", cell: (ctx) => cellString(ctx.getValue()) },
  { accessorKey: "senritu", header: "線量率", cell: (ctx) => cellString(ctx.getValue()) },
  { accessorKey: "syostat", header: "照射状態", cell: (ctx) => cellString(ctx.getValue()) },
  { accessorKey: "zhansuu", header: "残数", cell: (ctx) => cellString(ctx.getValue()) },
  { accessorKey: "syokind", header: "照射種別", cell: (ctx) => cellString(ctx.getValue()) },
  { accessorKey: "sdate", header: "開始日", cell: (ctx) => cellString(ctx.getValue()) },
  { accessorKey: "edate", header: "終了日", cell: (ctx) => cellString(ctx.getValue()) },
  { accessorKey: "bikou", header: "備考", cell: (ctx) => cellString(ctx.getValue()) },
];

export default function DosimetryIrradiationResultsPage() {
  const [source, setSource] = useState<IrradiationSource>("all");
  const [unoInput, setUnoInput] = useState("");
  const [kainameInput, setKainameInput] = useState("");
  const [searchUno, setSearchUno] = useState("");
  const [searchKainame, setSearchKainame] = useState("");
  const [page, setPage] = useState(1);
  const [pageSize, setPageSize] = useState(DEFAULT_PAGE_SIZE);

  const buildUrl = useCallback(() => {
    const p = new URLSearchParams();
    p.set("source", source);
    p.set("page", String(page));
    p.set("pageSize", String(pageSize));
    const u = searchUno.trim();
    const k = searchKainame.trim();
    if (u) p.set("uno", u);
    if (k) p.set("kainame", k);
    return `/api/dosimetry/irradiation-results?${p.toString()}`;
  }, [source, page, pageSize, searchUno, searchKainame]);

  const { data, isLoading, isError, error, refetch, isFetching } = useQuery({
    queryKey: [
      "dosimetry",
      "irradiation-results",
      { source, page, pageSize, searchUno, searchKainame },
    ],
    queryFn: () => apiFetch<ListResponse>(buildUrl()),
  });

  const rows = useMemo(() => data?.data.rows ?? [], [data]);
  const meta = data?.data.meta;
  const total = meta?.total;
  const totalPages = meta?.totalPages ?? 1;

  const handleSearch = () => {
    setSearchUno(unoInput);
    setSearchKainame(kainameInput);
    setPage(1);
  };

  const tablePageSize = Math.max(rows.length, 1);

  return (
    <div className="flex flex-col gap-6 p-4 md:p-6">
      <div>
        <h1 className="text-xl font-semibold text-ink">照射実績表示</h1>
        <p className="mt-1 text-sm text-ink-muted">
          照射マスタ（Syouk1 / Syouk2 / Syoukj3）を条件検索し、一覧表示します。データソースを絞り込むか「すべて」で
          3 テーブルを結合した結果を取得します。
        </p>
      </div>

      <div className="flex flex-wrap items-end gap-4">
        <label className="flex flex-col gap-1 text-sm text-ink-muted">
          データソース
          <select
            value={source}
            onChange={(e) => {
              setSource(e.target.value as IrradiationSource);
              setPage(1);
            }}
            className="min-w-[12rem] rounded-xl border border-cream-300 bg-white px-3 py-2 text-ink"
          >
            <option value="all">すべて（3テーブル結合）</option>
            <option value="syouk1">syouk1 のみ</option>
            <option value="syouk2">syouk2 のみ</option>
            <option value="syoukj3">syoukj3 のみ</option>
          </select>
        </label>
        <label className="flex flex-col gap-1 text-sm text-ink-muted">
          受付番号（部分一致）
          <input
            type="text"
            value={unoInput}
            onChange={(e) => setUnoInput(e.target.value)}
            className="min-w-[10rem] rounded-xl border border-cream-300 bg-white px-3 py-2 text-ink"
            placeholder="例: 12345"
          />
        </label>
        <label className="flex flex-col gap-1 text-sm text-ink-muted">
          会社名（部分一致）
          <input
            type="text"
            value={kainameInput}
            onChange={(e) => setKainameInput(e.target.value)}
            className="min-w-[12rem] rounded-xl border border-cream-300 bg-white px-3 py-2 text-ink"
            placeholder="会社名の一部"
          />
        </label>
        <label className="flex flex-col gap-1 text-sm text-ink-muted">
          1ページあたり
          <select
            value={pageSize}
            onChange={(e) => {
              setPageSize(Number(e.target.value));
              setPage(1);
            }}
            className="rounded-xl border border-cream-300 bg-white px-3 py-2 text-ink"
          >
            {[25, 50, 100, 200].map((n) => (
              <option key={n} value={n}>
                {n} 件
              </option>
            ))}
          </select>
        </label>
        <Button type="button" variant="default" onClick={() => void handleSearch()} disabled={isFetching}>
          検索
        </Button>
        <Button type="button" variant="outline" onClick={() => void refetch()} disabled={isFetching}>
          再取得
        </Button>
      </div>

      <div className="flex flex-wrap items-center gap-3 text-sm text-ink-muted">
        {total != null && (
          <span>
            全 {total} 件 · ページ {page} / {totalPages}
          </span>
        )}
        <div className="flex items-center gap-2">
          <Button
            type="button"
            variant="outline"
            size="sm"
            onClick={() => setPage((p) => Math.max(1, p - 1))}
            disabled={isFetching || page <= 1}
          >
            前のページ
          </Button>
          <Button
            type="button"
            variant="outline"
            size="sm"
            onClick={() => setPage((p) => Math.min(totalPages, p + 1))}
            disabled={isFetching || page >= totalPages}
          >
            次のページ
          </Button>
        </div>
      </div>

      {isError && (
        <p className="rounded-xl border border-red-200 bg-red-50 px-3 py-2 text-sm text-red-800">
          {error instanceof Error ? error.message : "取得に失敗しました"}
        </p>
      )}

      <DataTable columns={columns} data={rows} loading={isLoading} pageSize={tablePageSize} />
    </div>
  );
}
