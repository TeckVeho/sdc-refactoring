"use client";

import { useCallback, useEffect, useMemo, useState } from "react";

import { apiFetch } from "@/lib/api";
import { cn } from "@/lib/utils";
import { Button } from "@/components/ui/button";

type Ric3Item = {
  id: string;
  lastRepackDate: string;
  irradiationStatus: string;
  irradiationCode: string;
  doseMeterNo: string;
  receiptNo: string;
  companyName: string;
  dueDate: string;
  shipDate: string;
  passCount: string;
  note: string;
  checked: boolean;
};

type Ric3Product = {
  kaisyacd: string;
  sehncd: string;
  composite?: string;
  companyName: string;
  productName: string;
  tumikae: string;
  originalTumikae: string;
};

type Ric3Payload = {
  version: 1;
  items: Ric3Item[];
  products: Ric3Product[];
  updatedAt: string;
};

type Tab = "items" | "products";

export default function InventoryRic3RepackPage() {
  const [tab, setTab] = useState<Tab>("items");
  const [data, setData] = useState<Ric3Payload | null>(null);
  const [loading, setLoading] = useState(false);
  const [busy, setBusy] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [productDraft, setProductDraft] = useState<Ric3Product[]>([]);
  const [productFilter, setProductFilter] = useState("");

  const load = useCallback(async () => {
    setLoading(true);
    setError(null);
    try {
      const res = await apiFetch<{ data: Ric3Payload }>("/api/inventory/ric3-repack");
      setData(res.data);
      setProductDraft(res.data.products.map((p) => ({ ...p })));
    } catch (e) {
      setData(null);
      setError(e instanceof Error ? e.message : "取得に失敗しました");
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    void load();
  }, [load]);

  const refreshItems = useCallback(async () => {
    setBusy("items");
    setError(null);
    try {
      const res = await apiFetch<{ data: Ric3Payload }>("/api/inventory/ric3-repack/actions/refresh-items", {
        method: "POST",
        body: JSON.stringify({}),
      });
      setData(res.data);
      setProductDraft(res.data.products.map((p) => ({ ...p })));
    } catch (e) {
      setError(e instanceof Error ? e.message : "一覧の更新に失敗しました");
    } finally {
      setBusy(null);
    }
  }, []);

  const refreshProducts = useCallback(async () => {
    setBusy("products");
    setError(null);
    try {
      const res = await apiFetch<{ data: Ric3Payload }>("/api/inventory/ric3-repack/actions/refresh-products", {
        method: "POST",
        body: JSON.stringify({}),
      });
      setData(res.data);
      setProductDraft(res.data.products.map((p) => ({ ...p })));
    } catch (e) {
      setError(e instanceof Error ? e.message : "製品テーブルの更新に失敗しました");
    } finally {
      setBusy(null);
    }
  }, []);

  const saveProducts = useCallback(async () => {
    const rows = productDraft
      .filter((p) => p.tumikae !== p.originalTumikae)
      .map((p) => ({ kaisyacd: p.kaisyacd, sehncd: p.sehncd, tumikae: p.tumikae }));
    if (rows.length === 0) {
      setError("変更された詰替要否がありません");
      return;
    }
    setBusy("save-products");
    setError(null);
    try {
      const res = await apiFetch<{ data: Ric3Payload }>("/api/inventory/ric3-repack/products", {
        method: "PATCH",
        body: JSON.stringify({ rows }),
      });
      setData(res.data);
      setProductDraft(res.data.products.map((p) => ({ ...p })));
    } catch (e) {
      setError(e instanceof Error ? e.message : "登録の更新に失敗しました");
    } finally {
      setBusy(null);
    }
  }, [productDraft]);

  const filteredProducts = useMemo(() => {
    const q = productFilter.trim().toLowerCase();
    if (!q) return productDraft;
    return productDraft.filter((p) => {
      const hay = `${p.kaisyacd} ${p.sehncd} ${p.companyName} ${p.productName}`.toLowerCase();
      return hay.includes(q);
    });
  }, [productDraft, productFilter]);

  const toggleItemCheck = useCallback(
    async (id: string, checked: boolean) => {
      if (!data) return;
      setBusy(`item-${id}`);
      setError(null);
      try {
        const res = await apiFetch<{ data: Ric3Payload }>("/api/inventory/ric3-repack/items", {
          method: "PATCH",
          body: JSON.stringify({ items: [{ id, checked }] }),
        });
        setData(res.data);
      } catch (e) {
        setError(e instanceof Error ? e.message : "チェックの保存に失敗しました");
      } finally {
        setBusy(null);
      }
    },
    [data],
  );

  const items = data?.items ?? [];

  return (
    <div className="mx-auto max-w-6xl space-y-6 p-4 md:p-6">
      <div>
        <h1 className="text-2xl font-semibold text-ink">Ric3 詰替作業</h1>
        <p className="mt-1 text-sm text-ink-muted">
          積替え対象品の照射状況一覧と、RIC-3 向け詰替要否（ExR3Stg JSON）の編集を行います。一覧再取得は DB
          スナップショットに基づきます。
        </p>
      </div>

      <div className="flex flex-wrap items-center gap-2 rounded-2xl border border-cream-300 bg-cream-50 p-1">
        {(
          [
            ["items", "積替品一覧"],
            ["products", "積替製品テーブル"],
          ] as const
        ).map(([v, label]) => (
          <Button
            key={v}
            type="button"
            variant="ghost"
            className={cn(
              "rounded-xl px-4 py-2 text-sm",
              tab === v ? "bg-white text-ink shadow-sm ring-1 ring-cream-300" : "text-ink-muted hover:text-ink",
            )}
            onClick={() => setTab(v)}
          >
            {label}
          </Button>
        ))}
      </div>

      {tab === "items" ? (
        <div className="flex flex-wrap items-center gap-2">
          <Button type="button" onClick={() => void refreshItems()} disabled={busy !== null}>
            {busy === "items" ? "取得中…" : "現状の積替品の照射状況（再取得）"}
          </Button>
          <Button type="button" variant="outline" onClick={() => void load()} disabled={loading || busy !== null}>
            Stg から再読込
          </Button>
        </div>
      ) : (
        <div className="flex flex-wrap items-center gap-2">
          <label className="flex flex-col gap-1 text-sm text-ink-muted">
            絞り込み
            <input
              type="search"
              value={productFilter}
              onChange={(e) => setProductFilter(e.target.value)}
              placeholder="会社・製品コード・名称"
              className="min-w-[200px] rounded-xl border border-cream-300 bg-white px-3 py-2 text-ink"
            />
          </label>
          <Button type="button" onClick={() => void refreshProducts()} disabled={busy !== null}>
            {busy === "products" ? "取得中…" : "製品マスタ表示（再取得）"}
          </Button>
          <Button type="button" onClick={() => void saveProducts()} disabled={busy !== null}>
            {busy === "save-products" ? "保存中…" : "登録内容更新（Stg）"}
          </Button>
          <Button type="button" variant="outline" onClick={() => void load()} disabled={loading || busy !== null}>
            Stg から再読込
          </Button>
        </div>
      )}

      {error && <p className="text-sm text-red-700">{error}</p>}

      <p className="text-xs text-ink-muted">
        {loading ? "読み込み中…" : data && <>更新: {new Date(data.updatedAt).toLocaleString("ja-JP")}</>}
      </p>

      {tab === "items" ? (
        <div className="space-y-3">
          <p className="text-sm text-ink-muted">
            チェックは行ごとに API へ保存します。照射状況は syoukj3 / 計画 JSON（3号機）からのベストエフォート表示です。
          </p>
          <div className="max-h-[480px] overflow-auto rounded-2xl border border-cream-200">
            <table className="w-full min-w-[880px] border-collapse text-left text-sm">
              <thead className="sticky top-0 bg-cream-100 text-ink">
                <tr>
                  <th className="border-b border-cream-300 px-2 py-2">操作</th>
                  <th className="border-b border-cream-300 px-2 py-2">受付(下4)</th>
                  <th className="border-b border-cream-300 px-2 py-2">最終積替日</th>
                  <th className="border-b border-cream-300 px-2 py-2">照射状況</th>
                  <th className="border-b border-cream-300 px-2 py-2">コード</th>
                  <th className="border-b border-cream-300 px-2 py-2">線量計(下4)</th>
                  <th className="border-b border-cream-300 px-2 py-2">会社名</th>
                  <th className="border-b border-cream-300 px-2 py-2">納期</th>
                  <th className="border-b border-cream-300 px-2 py-2">出荷日</th>
                  <th className="border-b border-cream-300 px-2 py-2">パス数</th>
                  <th className="border-b border-cream-300 px-2 py-2">備考</th>
                  <th className="border-b border-cream-300 px-2 py-2">チェック</th>
                </tr>
              </thead>
              <tbody>
                {items.map((row) => (
                  <tr key={row.id} className="border-b border-cream-200 odd:bg-white even:bg-cream-50/60">
                    <td className="px-2 py-1">
                      <Button
                        type="button"
                        variant="ghost"
                        size="sm"
                        className="h-8 rounded-lg text-xs"
                        disabled={busy !== null}
                        onClick={() => void toggleItemCheck(row.id, !row.checked)}
                      >
                        {row.checked ? "チェック解除" : "チェック"}
                      </Button>
                    </td>
                    <td className="px-2 py-1">{row.receiptNo || "—"}</td>
                    <td className="px-2 py-1">{row.lastRepackDate || "—"}</td>
                    <td className="px-2 py-1">{row.irradiationStatus}</td>
                    <td className="px-2 py-1">{row.irradiationCode || "—"}</td>
                    <td className="px-2 py-1">{row.doseMeterNo || "—"}</td>
                    <td className="px-2 py-1">{row.companyName || "—"}</td>
                    <td className="px-2 py-1">{row.dueDate || "—"}</td>
                    <td className="px-2 py-1">{row.shipDate || "—"}</td>
                    <td className="px-2 py-1">{row.passCount || "—"}</td>
                    <td className="px-2 py-1 max-w-[200px] truncate" title={row.note}>
                      {row.note || "—"}
                    </td>
                    <td className="px-2 py-1">{row.checked ? "☑" : "□"}</td>
                  </tr>
                ))}
              </tbody>
            </table>
            {items.length === 0 && !loading && (
              <p className="p-4 text-sm text-ink-muted">データがありません。「再取得」で DB から読み込んでください。</p>
            )}
          </div>
        </div>
      ) : (
        <div className="max-h-[520px] overflow-auto rounded-2xl border border-cream-200">
          <table className="w-full min-w-[720px] border-collapse text-left text-sm">
            <thead className="sticky top-0 bg-cream-100 text-ink">
              <tr>
                <th className="border-b border-cream-300 px-2 py-2">会社コード</th>
                <th className="border-b border-cream-300 px-2 py-2">製品コード</th>
                <th className="border-b border-cream-300 px-2 py-2">会社名</th>
                <th className="border-b border-cream-300 px-2 py-2">製品名</th>
                <th className="border-b border-cream-300 px-2 py-2">詰替要否</th>
              </tr>
            </thead>
            <tbody>
              {filteredProducts.map((p) => (
                <tr
                  key={`${p.kaisyacd}-${p.sehncd}`}
                  className={cn(
                    "border-b border-cream-200 odd:bg-white even:bg-cream-50/60",
                    p.tumikae !== p.originalTumikae && "bg-amber-50/80",
                  )}
                >
                  <td className="px-2 py-1 font-mono text-xs">{p.kaisyacd}</td>
                  <td className="px-2 py-1 font-mono text-xs">{p.sehncd}</td>
                  <td className="px-2 py-1">{p.companyName || "—"}</td>
                  <td className="px-2 py-1">{p.productName || "—"}</td>
                  <td className="px-2 py-1">
                    <select
                      className="rounded-lg border border-cream-300 bg-white px-2 py-1 text-sm text-ink"
                      value={p.tumikae}
                      disabled={busy !== null}
                      aria-label={`詰替要否 ${p.kaisyacd}-${p.sehncd}`}
                      onChange={(e) => {
                        const v = e.target.value;
                        setProductDraft((prev) =>
                          prev.map((x) => (x.kaisyacd === p.kaisyacd && x.sehncd === p.sehncd ? { ...x, tumikae: v } : x)),
                        );
                      }}
                    >
                      <option value="0">不要 (0)</option>
                      <option value="1">要 (1)</option>
                    </select>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
          {filteredProducts.length === 0 && !loading && (
            <p className="p-4 text-sm text-ink-muted">
              行がありません。製品テーブルを再取得するか、絞り込みを解除してください。
            </p>
          )}
        </div>
      )}
    </div>
  );
}
