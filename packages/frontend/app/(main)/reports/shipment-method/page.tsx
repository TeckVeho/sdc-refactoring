"use client";

import { useCallback, useEffect, useMemo, useState } from "react";

import { Button } from "@/components/ui/button";
import { apiFetch } from "@/lib/api";

type Row = {
  kaisyacd: string;
  kairname: string;
  coname: string;
  hikitoriCurrent: string;
  hikitoriDb: string;
  housyubeCurrent: string;
  housyubeDb: string;
  reportFlagMax: number;
  needsReport: boolean;
};

const SHIPMENT_METHODS = [
  "引取",
  "混載便",
  "保管品",
  "チャータ便",
  "納品",
  "品証扱い",
  "",
] as const;

const REPORT_KINDS = [
  "照射後Fax送信",
  "出荷後Fax送信",
  "照射後報告書発行",
  "出荷後報告書発行",
  "不要",
  "",
] as const;

export default function ShipmentMethodReportPage() {
  const [rows, setRows] = useState<Row[]>([]);
  const [truncated, setTruncated] = useState(false);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [saving, setSaving] = useState(false);
  const [message, setMessage] = useState<string | null>(null);

  const load = useCallback(async () => {
    setLoading(true);
    setError(null);
    setMessage(null);
    try {
      const res = await apiFetch<{ data: { rows: Row[]; truncated: boolean } }>(
        "/api/reports/shipment-method",
      );
      setRows(res.data.rows);
      setTruncated(res.data.truncated);
    } catch (e) {
      setError(e instanceof Error ? e.message : "読込に失敗しました");
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    void load();
  }, [load]);

  const dirtyRows = useMemo(
    () =>
      rows.filter(
        (r) =>
          r.hikitoriCurrent.trim() !== r.hikitoriDb.trim() ||
          r.housyubeCurrent.trim() !== r.housyubeDb.trim(),
      ),
    [rows],
  );

  const updateRow = useCallback((kaisyacd: string, patch: Partial<Pick<Row, "hikitoriCurrent" | "housyubeCurrent">>) => {
    setRows((prev) =>
      prev.map((r) => (r.kaisyacd === kaisyacd ? { ...r, ...patch } : r)),
    );
  }, []);

  const save = useCallback(async () => {
    setSaving(true);
    setMessage(null);
    setError(null);
    try {
      const payload = {
        rows: dirtyRows.map((r) => ({
          kaisyacd: r.kaisyacd,
          hikitori: r.hikitoriCurrent,
          housyube: r.housyubeCurrent,
        })),
      };
      await apiFetch<{ ok: boolean; updated: number }>("/api/reports/shipment-method", {
        method: "PUT",
        body: JSON.stringify(payload),
      });
      setMessage("保存しました。一覧を再読込します。");
      await load();
    } catch (e) {
      setError(e instanceof Error ? e.message : "保存に失敗しました");
    } finally {
      setSaving(false);
    }
  }, [dirtyRows, load]);

  return (
    <div className="space-y-4 p-4">
      <div>
        <h1 className="text-lg font-semibold tracking-tight">出荷方法／報告書発行種別の登録</h1>
        <p className="text-muted-foreground mt-1 text-sm">
          TOKUMST（会社コード &lt; 2000）と ExSeihinj（shipment:コード）を組み合わせて表示します。
        </p>
      </div>

      <div className="flex flex-wrap items-center gap-2">
        <Button type="button" variant="outline" size="sm" onClick={() => void load()} disabled={loading}>
          再読込
        </Button>
        <Button type="button" size="sm" onClick={() => void save()} disabled={saving || dirtyRows.length === 0}>
          変更を保存（{dirtyRows.length} 件）
        </Button>
        {truncated ? (
          <span className="text-amber-700 text-xs">先頭 {rows.length} 件のみ表示（上限あり）</span>
        ) : null}
      </div>

      {error ? <p className="text-destructive text-sm">{error}</p> : null}
      {message ? <p className="text-sm text-emerald-800">{message}</p> : null}

      {loading ? (
        <p className="text-muted-foreground text-sm">読込中…</p>
      ) : (
        <div className="max-h-[calc(100vh-220px)] overflow-auto rounded-md border">
          <table className="w-full text-left text-xs">
            <thead className="bg-muted/80 sticky top-0 z-10">
              <tr>
                <th className="border-b px-2 py-2">コード</th>
                <th className="border-b px-2 py-2">略称</th>
                <th className="border-b px-2 py-2">会社名</th>
                <th className="border-b px-2 py-2">出荷方法</th>
                <th className="border-b px-2 py-2">DB値</th>
                <th className="border-b px-2 py-2">報告書種別</th>
                <th className="border-b px-2 py-2">DB値</th>
                <th className="border-b px-2 py-2">要報告</th>
              </tr>
            </thead>
            <tbody>
              {rows.map((r) => (
                <tr key={r.kaisyacd} className="odd:bg-background even:bg-muted/40">
                  <td className="border-b px-2 py-1 font-mono">{r.kaisyacd}</td>
                  <td className="border-b px-2 py-1">{r.kairname}</td>
                  <td className="border-b px-2 py-1">{r.coname}</td>
                  <td className="border-b px-2 py-1">
                    <select
                      className="w-full max-w-[140px] rounded border bg-background px-1 py-0.5"
                      value={r.hikitoriCurrent}
                      onChange={(e) =>
                        updateRow(r.kaisyacd, { hikitoriCurrent: e.target.value })
                      }
                    >
                      {SHIPMENT_METHODS.map((opt) => (
                        <option key={opt || "__empty"} value={opt}>
                          {opt || "（空）"}
                        </option>
                      ))}
                    </select>
                  </td>
                  <td className="border-b px-2 py-1 text-muted-foreground">{r.hikitoriDb || "—"}</td>
                  <td className="border-b px-2 py-1">
                    <select
                      className="w-full max-w-[180px] rounded border bg-background px-1 py-0.5"
                      value={r.housyubeCurrent}
                      disabled={!r.needsReport}
                      onChange={(e) =>
                        updateRow(r.kaisyacd, { housyubeCurrent: e.target.value })
                      }
                    >
                      {REPORT_KINDS.map((opt) => (
                        <option key={opt || "__empty"} value={opt}>
                          {opt || "（空）"}
                        </option>
                      ))}
                    </select>
                  </td>
                  <td className="border-b px-2 py-1 text-muted-foreground">{r.housyubeDb || "—"}</td>
                  <td className="border-b px-2 py-1">{r.needsReport ? "要" : "不要"}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}
