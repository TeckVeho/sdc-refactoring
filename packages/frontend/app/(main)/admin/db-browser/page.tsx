"use client";

import { useCallback, useEffect, useMemo, useState } from "react";

import { ExcelExportButton } from "@/components/shared/excel-export-button";
import { Button } from "@/components/ui/button";
import { apiFetch } from "@/lib/api";

type Meta = {
  data: { tables: { id: string; label: string }[]; maxRows: number };
};

type QueryPayload = {
  data: {
    table: string;
    take: number;
    skip: number;
    rowCount: number;
    rows: Record<string, unknown>[];
  };
};

export default function AdminDbBrowserPage() {
  const [meta, setMeta] = useState<Meta["data"] | null>(null);
  const [table, setTable] = useState("");
  const [q, setQ] = useState("");
  const [take, setTake] = useState("100");
  const [skip, setSkip] = useState("0");
  const [rows, setRows] = useState<Record<string, unknown>[]>([]);
  const [error, setError] = useState<string | null>(null);
  const [loading, setLoading] = useState(false);

  useEffect(() => {
    let cancelled = false;
    void apiFetch<Meta>("/api/admin/db-browser/meta")
      .then((r) => {
        if (cancelled) return;
        setMeta(r.data);
        setTable((prev) => prev || r.data.tables[0]?.id || "");
      })
      .catch(() => {
        if (!cancelled) setMeta(null);
      });
    return () => {
      cancelled = true;
    };
  }, []);

  const run = useCallback(async () => {
    const t = take.trim();
    const s = skip.trim();
    setLoading(true);
    setError(null);
    try {
      const p = new URLSearchParams({ table, take: t || "100", skip: s || "0" });
      if (q.trim()) p.set("q", q.trim());
      const res = await apiFetch<QueryPayload>(`/api/admin/db-browser?${p}`);
      setRows(res.data.rows);
    } catch (e) {
      setRows([]);
      setError(e instanceof Error ? e.message : "検索に失敗しました");
    } finally {
      setLoading(false);
    }
  }, [q, skip, table, take]);

  const exportRows = useMemo(() => rows as Record<string, unknown>[], [rows]);

  return (
    <div className="space-y-4 p-4">
      <div>
        <h1 className="text-lg font-semibold tracking-tight">DBブラウザ</h1>
        <p className="text-muted-foreground mt-1 text-sm">
          許可リスト内のテーブルのみ参照できます（最大 500 件）。結果は Excel で保存できます。
        </p>
      </div>

      <div className="flex flex-wrap items-end gap-3">
        <label className="flex flex-col gap-1 text-xs text-ink-muted">
          テーブル
          <select
            value={table}
            onChange={(e) => setTable(e.target.value)}
            className="min-w-[240px] rounded-xl border border-cream-300 bg-white px-3 py-2 text-sm text-ink"
          >
            {(meta?.tables ?? []).map((t) => (
              <option key={t.id} value={t.id}>
                {t.label}
              </option>
            ))}
          </select>
        </label>
        <label className="flex flex-col gap-1 text-xs text-ink-muted">
          検索文字列（contains）
          <input
            value={q}
            onChange={(e) => setQ(e.target.value)}
            className="w-48 rounded-xl border border-cream-300 bg-white px-3 py-2 text-sm text-ink"
          />
        </label>
        <label className="flex flex-col gap-1 text-xs text-ink-muted">
          take
          <input
            value={take}
            onChange={(e) => setTake(e.target.value)}
            className="w-20 rounded-xl border border-cream-300 bg-white px-3 py-2 font-mono text-sm text-ink"
          />
        </label>
        <label className="flex flex-col gap-1 text-xs text-ink-muted">
          skip
          <input
            value={skip}
            onChange={(e) => setSkip(e.target.value)}
            className="w-20 rounded-xl border border-cream-300 bg-white px-3 py-2 font-mono text-sm text-ink"
          />
        </label>
        <Button type="button" size="sm" onClick={() => void run()} disabled={loading || !table}>
          抽出
        </Button>
        {rows.length > 0 ? (
          <ExcelExportButton
            data={exportRows}
            fileName={`db-${table}-${Date.now()}`}
            sheetName={table || "data"}
          />
        ) : null}
      </div>

      {error ? <p className="text-destructive text-sm">{error}</p> : null}

      <div className="max-h-[calc(100vh-260px)] overflow-auto rounded-xl border border-cream-300 bg-white">
        <table className="w-full border-collapse text-xs">
          <thead className="sticky top-0 bg-cream-100">
            <tr>
              {rows[0]
                ? Object.keys(rows[0]).map((k) => (
                    <th key={k} className="border border-cream-300 px-2 py-1 text-left font-medium">
                      {k}
                    </th>
                  ))
                : null}
            </tr>
          </thead>
          <tbody>
            {rows.map((r, i) => (
              <tr key={i} className="odd:bg-white even:bg-cream-50/80">
                {Object.values(r).map((v, j) => (
                  <td key={j} className="border border-cream-200 px-2 py-1 font-mono">
                    {v === null || v === undefined ? "" : String(v)}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <p className="text-xs text-ink-muted">{rows.length} 件表示</p>
    </div>
  );
}
