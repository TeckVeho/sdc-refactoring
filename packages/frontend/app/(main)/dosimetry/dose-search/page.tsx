"use client";

import type { ColumnDef } from "@tanstack/react-table";
import { useCallback, useMemo, useState } from "react";

import { DataTable } from "@/components/shared/data-table";
import { DateRangePicker } from "@/components/shared/date-range-picker";
import { Button } from "@/components/ui/button";
import { apiFetch } from "@/lib/api";

type DoseSearchRow = {
  uno: string;
  kaisyacd: string;
  sehncd: string;
  syouso: string | null;
  kainame: string | null;
  nyukabi: string | null;
  nouki: string | null;
  pass: string | null;
  source: "zaiko" | "zaikor";
  siteisn: string | null;
  senritu: string | null;
  syokind: string | null;
  syostat: string | null;
  syosuu: number | null;
  zhansuu: number | null;
  jno: string | null;
  j2meisai: string | null;
  j2status: string | null;
};

type KansokuMaster = { kid: string; sokutei: string | null; bikou: string | null };

type SearchPayload = {
  data: {
    rows: DoseSearchRow[];
    truncated: boolean;
    kansokuMasters: KansokuMaster[];
  };
};

function isoDateToYyyymmdd(d: Date): string {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}${m}${day}`;
}

function countConditions(
  hasDate: boolean,
  hasUno: boolean,
  hasKai: boolean,
  hasSeh: boolean,
): number {
  let n = 0;
  if (hasDate) n += 1;
  if (hasUno) n += 1;
  if (hasKai) n += 1;
  if (hasSeh) n += 1;
  return n;
}

export default function DosimetryDoseSearchPage() {
  const [range, setRange] = useState(() => {
    const to = new Date();
    const from = new Date();
    from.setDate(from.getDate() - 30);
    return { from, to };
  });
  const [useDate, setUseDate] = useState(true);
  const [uno, setUno] = useState("");
  const [kaisyacd, setKaisyacd] = useState("");
  const [sehncd, setSehncd] = useState("");
  const [source, setSource] = useState<"zaiko" | "zaikor">("zaiko");

  const [rows, setRows] = useState<DoseSearchRow[]>([]);
  const [kansokuMasters, setKansokuMasters] = useState<KansokuMaster[]>([]);
  const [truncated, setTruncated] = useState(false);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [clientHint, setClientHint] = useState<string | null>(null);

  const search = useCallback(async () => {
    const fromY = useDate ? isoDateToYyyymmdd(range.from) : "";
    const toY = useDate ? isoDateToYyyymmdd(range.to) : "";
    const u = uno.trim();
    const k = kaisyacd.trim();
    const s = sehncd.trim();
    const hasDate = useDate && fromY.length === 8 && toY.length === 8;
    const hasUno = u.length > 0;
    const hasKai = k.length > 0;
    const hasSeh = s.length > 0;
    if (countConditions(hasDate, hasUno, hasKai, hasSeh) < 2) {
      setClientHint("日付範囲・受付番号・会社コード・製品コードのうち2つ以上を指定してください。");
      return;
    }
    setClientHint(null);
    setLoading(true);
    setError(null);
    try {
      const params = new URLSearchParams();
      params.set("source", source);
      if (hasDate) {
        params.set("dateFrom", fromY);
        params.set("dateTo", toY);
      }
      if (hasUno) params.set("uno", u);
      if (hasKai) params.set("kaisyacd", k);
      if (hasSeh) params.set("sehncd", s);
      const res = await apiFetch<SearchPayload>(`/api/dosimetry/dose-search?${params.toString()}`);
      setRows(res.data.rows);
      setTruncated(res.data.truncated);
      setKansokuMasters(res.data.kansokuMasters);
    } catch (e) {
      setRows([]);
      setTruncated(false);
      setKansokuMasters([]);
      setError(e instanceof Error ? e.message : "検索に失敗しました");
    } finally {
      setLoading(false);
    }
  }, [range.from, range.to, useDate, uno, kaisyacd, sehncd, source]);

  const columns: ColumnDef<DoseSearchRow>[] = useMemo(
    () => [
      {
        accessorKey: "source",
        header: "種別",
        cell: ({ row }) => (row.original.source === "zaikor" ? "履歴" : "在庫"),
      },
      { accessorKey: "uno", header: "受付番号" },
      { accessorKey: "kaisyacd", header: "会社" },
      { accessorKey: "sehncd", header: "製品" },
      { accessorKey: "nyukabi", header: "入荷日" },
      { accessorKey: "nouki", header: "納期" },
      { accessorKey: "syouso", header: "装置" },
      { accessorKey: "siteisn", header: "指定線量" },
      { accessorKey: "senritu", header: "線量率" },
      { accessorKey: "zhansuu", header: "残数" },
      { accessorKey: "j2meisai", header: "明細(syouj2)" },
    ],
    [],
  );

  return (
    <div className="mx-auto max-w-6xl space-y-6 p-4 md:p-6">
      <div>
        <h1 className="text-2xl font-semibold text-ink">線量検索</h1>
        <p className="mt-1 text-sm text-ink-muted">
          在庫（zaiko）／履歴（zaikor）のヘッダを検索し、Syouk1（指定線量等）と Syouj2 先頭1件を結合して表示します。MVP
          では日付範囲・受付番号・会社コード・製品コードのうち2条件以上が必要です。
        </p>
      </div>

      <div className="space-y-4 rounded-2xl border border-cream-300 bg-cream-50 p-4">
        <div className="flex flex-wrap items-center gap-4">
          <label className="flex items-center gap-2 text-sm text-ink">
            <input
              type="checkbox"
              checked={useDate}
              onChange={(e) => setUseDate(e.target.checked)}
              className="rounded border-cream-300"
            />
            入荷日（nyukabi）で絞る
          </label>
        </div>
        {useDate ? (
          <DateRangePicker label="日付範囲" value={range} onChange={setRange} />
        ) : (
          <p className="text-sm text-ink-muted">日付条件は使用しません。他の条件を2つ以上指定してください。</p>
        )}

        <div className="flex flex-wrap items-end gap-4">
          <label className="flex flex-col gap-1 text-sm text-ink-muted">
            受付番号（10桁 または 開始-終了）
            <input
              value={uno}
              onChange={(e) => setUno(e.target.value)}
              placeholder="例: 2024010001 または 2024010000-2024010099"
              className="min-w-[240px] rounded-xl border border-cream-300 bg-white px-3 py-2 text-ink focus:outline-none focus:ring-2 focus:ring-terracotta/40"
            />
          </label>
          <label className="flex flex-col gap-1 text-sm text-ink-muted">
            会社コード
            <input
              value={kaisyacd}
              onChange={(e) => setKaisyacd(e.target.value)}
              className="w-28 rounded-xl border border-cream-300 bg-white px-3 py-2 text-ink focus:outline-none focus:ring-2 focus:ring-terracotta/40"
            />
          </label>
          <label className="flex flex-col gap-1 text-sm text-ink-muted">
            製品コード
            <input
              value={sehncd}
              onChange={(e) => setSehncd(e.target.value)}
              className="w-28 rounded-xl border border-cream-300 bg-white px-3 py-2 text-ink focus:outline-none focus:ring-2 focus:ring-terracotta/40"
            />
          </label>
          <label className="flex flex-col gap-1 text-sm text-ink-muted">
            対象
            <select
              value={source}
              onChange={(e) => setSource(e.target.value as "zaiko" | "zaikor")}
              className="rounded-xl border border-cream-300 bg-white px-3 py-2 text-ink focus:outline-none focus:ring-2 focus:ring-terracotta/40"
            >
              <option value="zaiko">在庫（zaiko）</option>
              <option value="zaikor">履歴（zaikor）</option>
            </select>
          </label>
          <Button type="button" onClick={() => void search()} disabled={loading}>
            {loading ? "検索中…" : "検索"}
          </Button>
        </div>
      </div>

      {clientHint ? <p className="text-sm text-amber-800">{clientHint}</p> : null}
      {error ? <p className="text-sm text-red-700">{error}</p> : null}
      {truncated ? (
        <p className="text-sm text-amber-800">件数上限に達したため、以降の行は切り捨てられています。</p>
      ) : null}
      {kansokuMasters.length > 0 ? (
        <p className="text-xs text-ink-muted">観測マスタ（Kansoku）: {kansokuMasters.length} 件を参照用に同梱</p>
      ) : null}

      <DataTable columns={columns} data={rows} loading={loading} pageSize={20} />
    </div>
  );
}
