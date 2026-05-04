"use client";

import { type ChangeEvent, useCallback, useEffect, useState } from "react";

import { Button } from "@/components/ui/button";
import { apiFetch } from "@/lib/api";

type Row = { gammacd: string; ebcd: string; memo: string | null };

export default function AdminCompanyCodeConvertPage() {
  const [rows, setRows] = useState<Row[]>([]);
  const [qInput, setQInput] = useState("");
  const [appliedQ, setAppliedQ] = useState("");
  const [gammacd, setGammacd] = useState("");
  const [ebcd, setEbcd] = useState("");
  const [memo, setMemo] = useState("");
  const [error, setError] = useState<string | null>(null);
  const [loading, setLoading] = useState(false);

  const load = useCallback(async () => {
    setLoading(true);
    setError(null);
    try {
      const p = appliedQ ? `?${new URLSearchParams({ q: appliedQ })}` : "";
      const res = await apiFetch<{ data: { rows: Row[] } }>(`/api/admin/company-code-convert${p}`);
      setRows(res.data.rows);
    } catch (e) {
      setError(e instanceof Error ? e.message : "読込に失敗しました");
    } finally {
      setLoading(false);
    }
  }, [appliedQ]);

  useEffect(() => {
    void load();
  }, [load]);

  const upsert = useCallback(async () => {
    setLoading(true);
    setError(null);
    try {
      await apiFetch("/api/admin/company-code-convert", {
        method: "POST",
        body: JSON.stringify({
          gammacd: gammacd.trim(),
          ebcd: ebcd.trim(),
          memo: memo.trim() || undefined,
        }),
      });
      setGammacd("");
      setEbcd("");
      setMemo("");
      await load();
    } catch (e) {
      setError(e instanceof Error ? e.message : "登録に失敗しました");
    } finally {
      setLoading(false);
    }
  }, [ebcd, gammacd, load, memo]);

  const del = useCallback(
    async (g: string, e: string) => {
      if (!confirm(`${g} / ${e} を削除しますか？`)) return;
      setLoading(true);
      setError(null);
      try {
        await apiFetch(
          `/api/admin/company-code-convert?${new URLSearchParams({ gammacd: g, ebcd: e })}`,
          { method: "DELETE" },
        );
        await load();
      } catch (err) {
        setError(err instanceof Error ? err.message : "削除に失敗しました");
      } finally {
        setLoading(false);
      }
    },
    [load],
  );

  return (
    <div className="space-y-4 p-4">
      <div>
        <h1 className="text-lg font-semibold tracking-tight">GM / EB 会社コード変換</h1>
        <p className="text-muted-foreground mt-1 text-sm">kcdcnvmst テーブルを編集します。</p>
      </div>

      <div className="flex flex-wrap items-end gap-3">
        <label className="flex flex-col gap-1 text-xs text-ink-muted">
          フィルタ
          <input
            value={qInput}
            onChange={(e: ChangeEvent<HTMLInputElement>) => setQInput(e.target.value)}
            className="w-48 rounded-xl border border-cream-300 bg-white px-3 py-2 text-sm text-ink"
          />
        </label>
        <Button
          type="button"
          variant="outline"
          size="sm"
          onClick={() => setAppliedQ(qInput.trim())}
          disabled={loading}
        >
          絞込
        </Button>
        <Button type="button" variant="outline" size="sm" onClick={() => void load()} disabled={loading}>
          再読込
        </Button>
      </div>

      <div className="grid gap-2 rounded-2xl border border-cream-300 bg-cream-50 p-4 sm:grid-cols-4">
        <label className="flex flex-col gap-1 text-xs text-ink-muted sm:col-span-1">
          ガンマコード
          <input
            value={gammacd}
            onChange={(e) => setGammacd(e.target.value)}
            className="rounded-xl border border-cream-300 bg-white px-3 py-2 font-mono text-sm text-ink"
          />
        </label>
        <label className="flex flex-col gap-1 text-xs text-ink-muted sm:col-span-1">
          EBコード
          <input
            value={ebcd}
            onChange={(e) => setEbcd(e.target.value)}
            className="rounded-xl border border-cream-300 bg-white px-3 py-2 font-mono text-sm text-ink"
          />
        </label>
        <label className="flex flex-col gap-1 text-xs text-ink-muted sm:col-span-2">
          メモ
          <input
            value={memo}
            onChange={(e) => setMemo(e.target.value)}
            className="rounded-xl border border-cream-300 bg-white px-3 py-2 text-sm text-ink"
          />
        </label>
        <div className="flex items-end sm:col-span-4">
          <Button type="button" size="sm" onClick={() => void upsert()} disabled={loading || !gammacd.trim() || !ebcd.trim()}>
            登録／更新
          </Button>
        </div>
      </div>

      {error ? <p className="text-destructive text-sm">{error}</p> : null}

      <div className="max-h-[calc(100vh-380px)] overflow-auto rounded-xl border border-cream-300 bg-white">
        <table className="w-full border-collapse text-sm">
          <thead className="sticky top-0 bg-cream-100">
            <tr>
              <th className="border border-cream-300 px-2 py-2 text-left">ガンマ</th>
              <th className="border border-cream-300 px-2 py-2 text-left">EB</th>
              <th className="border border-cream-300 px-2 py-2 text-left">メモ</th>
              <th className="border border-cream-300 px-2 py-2 w-24" />
            </tr>
          </thead>
          <tbody>
            {rows.map((r) => (
              <tr key={`${r.gammacd}-${r.ebcd}`} className="odd:bg-white even:bg-cream-50/80">
                <td className="border border-cream-200 px-2 py-1 font-mono">{r.gammacd}</td>
                <td className="border border-cream-200 px-2 py-1 font-mono">{r.ebcd}</td>
                <td className="border border-cream-200 px-2 py-1">{r.memo ?? ""}</td>
                <td className="border border-cream-200 px-2 py-1">
                  <Button type="button" variant="ghost" size="sm" className="text-destructive h-8" onClick={() => void del(r.gammacd, r.ebcd)}>
                    削除
                  </Button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}
