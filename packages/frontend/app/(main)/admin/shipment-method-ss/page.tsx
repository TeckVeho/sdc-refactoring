"use client";

import { useCallback, useEffect, useState } from "react";

import { apiFetch } from "@/lib/api";

type Row = {
  kaisyacd: string;
  kairname: string;
  coname: string;
  hikitoriCurrent: string;
  housyubeCurrent: string;
  needsReport: boolean;
};

type Payload = {
  data: {
    rows: Row[];
    truncated: boolean;
    view: "ss";
    ref: string;
  };
};

export default function ShipmentMethodSsPage() {
  const [data, setData] = useState<Payload["data"] | null>(null);
  const [error, setError] = useState<string | null>(null);

  const load = useCallback(async () => {
    setError(null);
    try {
      const res = await apiFetch<Payload>("/api/admin/shipment-method-ss");
      setData(res.data);
    } catch (e) {
      setData(null);
      setError(e instanceof Error ? e.message : "読込に失敗しました");
    }
  }, []);

  useEffect(() => {
    void load();
  }, [load]);

  return (
    <div className="space-y-4 p-4">
      <div className="rounded-xl border border-sky-300 bg-sky-50 px-4 py-3 text-sm text-sky-950">
        <strong>SS 検証ビュー</strong>（読取専用）— データソースは本番と同一です。
      </div>
      <div>
        <h1 className="text-lg font-semibold tracking-tight">出荷方法（SS検証）</h1>
        <p className="text-muted-foreground mt-1 text-xs">{data?.ref}</p>
      </div>

      {error ? <p className="text-destructive text-sm">{error}</p> : null}

      {data ? (
        <div className="max-h-[calc(100vh-220px)] overflow-auto rounded-xl border border-cream-300 bg-white">
          <table className="w-full text-left text-xs">
            <thead className="bg-cream-100 sticky top-0">
              <tr>
                <th className="border-b px-2 py-2">コード</th>
                <th className="border-b px-2 py-2">略称</th>
                <th className="border-b px-2 py-2">会社名</th>
                <th className="border-b px-2 py-2">出荷方法</th>
                <th className="border-b px-2 py-2">報告書種別</th>
                <th className="border-b px-2 py-2">要報告</th>
              </tr>
            </thead>
            <tbody>
              {data.rows.map((r) => (
                <tr key={r.kaisyacd} className="odd:bg-white even:bg-cream-50/60">
                  <td className="border-b px-2 py-1 font-mono">{r.kaisyacd}</td>
                  <td className="border-b px-2 py-1">{r.kairname}</td>
                  <td className="border-b px-2 py-1">{r.coname}</td>
                  <td className="border-b px-2 py-1">{r.hikitoriCurrent}</td>
                  <td className="border-b px-2 py-1">{r.housyubeCurrent}</td>
                  <td className="border-b px-2 py-1">{r.needsReport ? "要" : "不要"}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      ) : null}
    </div>
  );
}
