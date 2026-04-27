"use client";

import type { ColumnDef } from "@tanstack/react-table";
import { useCallback, useMemo, useState } from "react";

import { DataTable } from "@/components/shared/data-table";
import { apiFetch } from "@/lib/api";
import { Button } from "@/components/ui/button";
import { cn } from "@/lib/utils";

type UnissuedRow = {
  rowkey: string;
  uno: string;
  kaisyacd: string;
  kainame: string;
  sehncd: string;
  syinji: string;
  houfax: string;
  hyouji: string;
  rotno2: string;
  syukkabi: string;
  json: Record<string, unknown>;
};

type SearchPayload = {
  data: {
    rows: UnissuedRow[];
    truncated: boolean;
  };
};

function isoDateToYyyymmdd(iso: string): string {
  return iso.replaceAll("-", "");
}

export default function ReportsUnissuedSearchPage() {
  const [minUno, setMinUno] = useState("");
  const [kaisyacd, setKaisyacd] = useState("");
  const [maxShipDate, setMaxShipDate] = useState(() => {
    const t = new Date();
    return `${t.getFullYear()}-${String(t.getMonth() + 1).padStart(2, "0")}-${String(t.getDate()).padStart(2, "0")}`;
  });

  const [rows, setRows] = useState<UnissuedRow[]>([]);
  const [truncated, setTruncated] = useState(false);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [selected, setSelected] = useState<UnissuedRow | null>(null);

  const [editHoufax, setEditHoufax] = useState("");
  const [editHyouji, setEditHyouji] = useState("");
  const [editRotno2, setEditRotno2] = useState("");
  const [saving, setSaving] = useState(false);
  const [saveMsg, setSaveMsg] = useState<string | null>(null);

  const search = useCallback(async () => {
    setLoading(true);
    setError(null);
    setSaveMsg(null);
    try {
      const params = new URLSearchParams();
      if (minUno.trim()) params.set("minUno", minUno.trim());
      if (kaisyacd.trim()) params.set("kaisyacd", kaisyacd.trim());
      if (maxShipDate) params.set("maxShipDate", isoDateToYyyymmdd(maxShipDate));

      const res = await apiFetch<SearchPayload>(`/api/reports/unissued-search?${params}`);
      setRows(res.data.rows);
      setTruncated(res.data.truncated);
      setSelected(null);
    } catch (e) {
      setRows([]);
      setTruncated(false);
      setError(e instanceof Error ? e.message : "検索に失敗しました");
    } finally {
      setLoading(false);
    }
  }, [minUno, kaisyacd, maxShipDate]);

  const openEdit = useCallback((r: UnissuedRow) => {
    setSelected(r);
    setEditHoufax(r.houfax || "");
    setEditHyouji(r.hyouji || "");
    setEditRotno2(r.rotno2 || "");
    setSaveMsg(null);
  }, []);

  const savePatch = useCallback(async () => {
    if (!selected) return;
    setSaving(true);
    setSaveMsg(null);
    try {
      const jsonPatch: Record<string, string> = {};
      if (editHoufax !== (selected.houfax || "")) jsonPatch.houfax = editHoufax;
      if (editHyouji !== (selected.hyouji || "")) jsonPatch.hyouji = editHyouji;
      if (editRotno2 !== (selected.rotno2 || "")) jsonPatch.rotno2 = editRotno2;

      if (Object.keys(jsonPatch).length === 0) {
        setSaveMsg("変更がありません");
        setSaving(false);
        return;
      }

      await apiFetch<{ ok: boolean }>(`/api/reports/unissued-search/rows/${encodeURIComponent(selected.rowkey)}`, {
        method: "PATCH",
        body: JSON.stringify({ jsonPatch }),
      });
      setSaveMsg("保存しました");
      await search();
    } catch (e) {
      setSaveMsg(e instanceof Error ? e.message : "保存に失敗しました");
    } finally {
      setSaving(false);
    }
  }, [selected, editHoufax, editHyouji, editRotno2, search]);

  const columns: ColumnDef<UnissuedRow>[] = useMemo(
    () => [
      {
        id: "select",
        header: "",
        cell: ({ row }) => (
          <Button type="button" variant="outline" size="sm" className="h-7 text-xs" onClick={() => openEdit(row.original)}>
            選択
          </Button>
        ),
        enableSorting: false,
      },
      { accessorKey: "uno", header: "受付番号" },
      { accessorKey: "kaisyacd", header: "会社コード" },
      { accessorKey: "kainame", header: "顧客名" },
      { accessorKey: "sehncd", header: "製品コード" },
      { accessorKey: "syinji", header: "発行フラグ" },
      { accessorKey: "houfax", header: "Fax" },
      { accessorKey: "hyouji", header: "表示" },
      { accessorKey: "rotno2", header: "ロット" },
      { accessorKey: "syukkabi", header: "出荷日" },
    ],
    [openEdit],
  );

  return (
    <div className="mx-auto max-w-6xl space-y-6 p-4 md:p-6">
      <div>
        <h1 className="text-2xl font-semibold text-ink">報告書未発行検索</h1>
        <p className="mt-1 text-sm text-ink-muted">
          ExKeikakuX の JSON を対象に、未発行・Fax 未送信相当の行を検索します。行を選択して Fax・表示・ロットを更新できます。
        </p>
      </div>

      <div className="flex flex-wrap items-end gap-4 rounded-2xl border border-cream-300 bg-cream-50 p-4">
        <label className="flex flex-col gap-1 text-sm text-ink-muted">
          検索開始受付番号（minUno）
          <input
            value={minUno}
            onChange={(e) => setMinUno(e.target.value)}
            placeholder="例: 2024010000"
            className="rounded-xl border border-cream-300 bg-white px-3 py-2 text-ink focus:outline-none focus:ring-2 focus:ring-terracotta/40"
          />
        </label>
        <label className="flex flex-col gap-1 text-sm text-ink-muted">
          会社コード
          <input
            value={kaisyacd}
            onChange={(e) => setKaisyacd(e.target.value)}
            className="w-36 rounded-xl border border-cream-300 bg-white px-3 py-2 text-ink focus:outline-none focus:ring-2 focus:ring-terracotta/40"
          />
        </label>
        <label className="flex flex-col gap-1 text-sm text-ink-muted">
          最終出荷日まで（YYYYMMDD）
          <input
            type="date"
            value={maxShipDate}
            onChange={(e) => setMaxShipDate(e.target.value)}
            className="rounded-xl border border-cream-300 bg-white px-3 py-2 text-ink focus:outline-none focus:ring-2 focus:ring-terracotta/40"
          />
        </label>
        <Button type="button" onClick={() => void search()} disabled={loading}>
          {loading ? "検索中…" : "検索"}
        </Button>
      </div>

      {error ? <p className="text-sm text-red-700">{error}</p> : null}
      {truncated ? (
        <p className="text-sm text-amber-800">件数上限に達したため結果が切り詰められている可能性があります。</p>
      ) : null}

      <DataTable columns={columns} data={rows} loading={loading} pageSize={20} />

      <div
        className={cn(
          "rounded-2xl border border-cream-300 bg-cream-50 p-4 transition-opacity",
          !selected && "opacity-60",
        )}
      >
        <h2 className="text-sm font-semibold text-ink">行の更新（PATCH）</h2>
        {!selected ? (
          <p className="mt-2 text-sm text-ink-muted">表の「選択」から行を選んでください。</p>
        ) : (
          <div className="mt-3 flex flex-col gap-3 sm:flex-row sm:flex-wrap sm:items-end">
            <p className="w-full text-xs text-ink-muted">rowkey: {selected.rowkey}</p>
            <label className="flex flex-col gap-1 text-sm text-ink-muted">
              Fax（houfax）
              <select
                value={editHoufax}
                onChange={(e) => setEditHoufax(e.target.value)}
                className="rounded-xl border border-cream-300 bg-white px-3 py-2 text-ink"
              >
                <option value="">（空）</option>
                <option value="未">未</option>
                <option value="済">済</option>
                <option value="不">不</option>
              </select>
            </label>
            <label className="flex flex-col gap-1 text-sm text-ink-muted">
              表示（hyouji）
              <select
                value={editHyouji}
                onChange={(e) => setEditHyouji(e.target.value)}
                className="rounded-xl border border-cream-300 bg-white px-3 py-2 text-ink"
              >
                <option value="">（表示）</option>
                <option value="非">非</option>
              </select>
            </label>
            <label className="flex flex-col gap-1 text-sm text-ink-muted">
              ロット（rotno2）
              <input
                value={editRotno2}
                onChange={(e) => setEditRotno2(e.target.value)}
                className="rounded-xl border border-cream-300 bg-white px-3 py-2 text-ink"
              />
            </label>
            <Button type="button" onClick={() => void savePatch()} disabled={saving}>
              {saving ? "保存中…" : "変更を保存"}
            </Button>
          </div>
        )}
        {saveMsg ? <p className="mt-2 text-sm text-ink-muted">{saveMsg}</p> : null}
      </div>
    </div>
  );
}
