"use client";

import { type ChangeEvent, useCallback, useState } from "react";

import { Button } from "@/components/ui/button";
import { apiFetch } from "@/lib/api";

type PricePayload = {
  data: {
    uno: string | null;
    kaisyacd: string;
    sehncd: string;
    companyName: string | null;
    productName: string | null;
    priceRowkey: string;
    price: Record<string, string>;
    note: string;
  };
};

export default function AdminPriceSearchPage() {
  const [uno, setUno] = useState("");
  const [kaisyacd, setKaisyacd] = useState("");
  const [sehncd, setSehncd] = useState("");
  const [data, setData] = useState<PricePayload["data"] | null>(null);
  const [fields, setFields] = useState<Record<string, string>>({});
  const [error, setError] = useState<string | null>(null);
  const [loading, setLoading] = useState(false);

  const search = useCallback(async () => {
    setLoading(true);
    setError(null);
    try {
      const p = new URLSearchParams();
      const u = uno.trim();
      if (u) p.set("uno", u);
      else {
        p.set("kaisyacd", kaisyacd.trim());
        p.set("sehncd", sehncd.trim());
      }
      const res = await apiFetch<PricePayload>(`/api/admin/price-search?${p}`);
      setData(res.data);
      setFields({ ...res.data.price });
    } catch (e) {
      setData(null);
      setFields({});
      setError(e instanceof Error ? e.message : "検索に失敗しました");
    } finally {
      setLoading(false);
    }
  }, [kaisyacd, sehncd, uno]);

  const savePrice = useCallback(async () => {
    if (!data) return;
    setLoading(true);
    setError(null);
    try {
      await apiFetch("/api/admin/price-search", {
        method: "PUT",
        body: JSON.stringify({
          kaisyacd: data.kaisyacd,
          sehncd: data.sehncd,
          fields,
        }),
      });
      await search();
    } catch (e) {
      setError(e instanceof Error ? e.message : "保存に失敗しました");
    } finally {
      setLoading(false);
    }
  }, [data, fields, search]);

  const onField =
    (k: string) =>
    (e: ChangeEvent<HTMLInputElement>) =>
      setFields((prev) => ({ ...prev, [k]: e.target.value }));

  const primaryKeys = ["tanka", "tanni", "tourokubi", "dose", "souti", "pass", "folder", "file"];

  return (
    <div className="space-y-4 p-4">
      <div>
        <h1 className="text-lg font-semibold tracking-tight">単価検索</h1>
        <p className="text-muted-foreground mt-1 text-sm">
          受付番号、または会社コード＋製品コードで検索します。単価は ExSeihinz（price:コード）に保存します。
        </p>
      </div>

      <div className="flex flex-wrap items-end gap-3">
        <label className="flex flex-col gap-1 text-xs text-ink-muted">
          受付番号（uno）
          <input
            value={uno}
            onChange={(e) => setUno(e.target.value)}
            className="w-44 rounded-xl border border-cream-300 bg-white px-3 py-2 font-mono text-sm text-ink"
          />
        </label>
        <span className="pb-2 text-xs text-ink-muted">または</span>
        <label className="flex flex-col gap-1 text-xs text-ink-muted">
          会社コード
          <input
            value={kaisyacd}
            onChange={(e) => setKaisyacd(e.target.value)}
            className="w-28 rounded-xl border border-cream-300 bg-white px-3 py-2 font-mono text-sm text-ink"
          />
        </label>
        <label className="flex flex-col gap-1 text-xs text-ink-muted">
          製品コード
          <input
            value={sehncd}
            onChange={(e) => setSehncd(e.target.value)}
            className="w-28 rounded-xl border border-cream-300 bg-white px-3 py-2 font-mono text-sm text-ink"
          />
        </label>
        <Button type="button" size="sm" onClick={() => void search()} disabled={loading}>
          検索
        </Button>
        <Button type="button" size="sm" variant="outline" onClick={() => void savePrice()} disabled={loading || !data}>
          単価フィールドを保存
        </Button>
      </div>

      {error ? <p className="text-destructive text-sm">{error}</p> : null}

      {data ? (
        <div className="space-y-3 rounded-2xl border border-cream-300 bg-cream-50 p-4 text-sm">
          <div>
            <span className="text-ink-muted">rowkey: </span>
            <span className="font-mono text-xs">{data.priceRowkey}</span>
          </div>
          <div>
            {data.companyName ?? "—"} / {data.productName ?? "—"}
          </div>
          <p className="text-xs text-ink-muted">{data.note}</p>

          <div className="grid gap-2 sm:grid-cols-2">
            {primaryKeys.map((k) => (
              <label key={k} className="flex flex-col gap-1 text-xs capitalize text-ink-muted">
                {k}
                <input
                  value={fields[k] ?? ""}
                  onChange={onField(k)}
                  className="rounded-xl border border-cream-300 bg-white px-3 py-2 text-sm text-ink"
                />
              </label>
            ))}
          </div>

          <details className="text-xs">
            <summary className="cursor-pointer text-ink-muted">その他のキー</summary>
            <div className="mt-2 grid gap-2 sm:grid-cols-2">
              {Object.keys(fields)
                .filter((k) => !primaryKeys.includes(k))
                .map((k) => (
                  <label key={k} className="flex flex-col gap-1 text-xs text-ink-muted">
                    {k}
                    <input
                      value={fields[k] ?? ""}
                      onChange={onField(k)}
                      className="rounded-xl border border-cream-300 bg-white px-3 py-2 text-sm text-ink"
                    />
                  </label>
                ))}
            </div>
          </details>
        </div>
      ) : null}
    </div>
  );
}
