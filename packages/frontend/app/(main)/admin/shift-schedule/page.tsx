"use client";

import { useCallback, useEffect, useMemo, useState } from "react";

import { Button } from "@/components/ui/button";
import { apiFetch } from "@/lib/api";

type Cycle = {
  year: number;
  month: number;
  startIso: string;
  endIso: string;
  dates: string[];
};

type Payload = {
  data: {
    cycle: Cycle;
    scheduleId: string | null;
    label: string | null;
    symbols: { symbol: string; memo: string; hours: string | null }[];
    circulationList: { id: string; name: string; sortOrder: number | null }[];
    employees: { shano: string; shaname: string; inRoster: boolean }[];
    grid: Record<string, Record<string, string>>;
    rosterSize: number;
  };
};

export default function AdminShiftSchedulePage() {
  const now = new Date();
  const [year, setYear] = useState(String(now.getFullYear()));
  const [month, setMonth] = useState(String(now.getMonth() + 1));
  const [payload, setPayload] = useState<Payload["data"] | null>(null);
  const [grid, setGrid] = useState<Record<string, Record<string, string>>>({});
  const [label, setLabel] = useState("");
  const [error, setError] = useState<string | null>(null);
  const [loading, setLoading] = useState(false);
  const [saving, setSaving] = useState(false);

  const applyPayload = useCallback((d: Payload["data"]) => {
    setPayload(d);
    setGrid(structuredClone(d.grid ?? {}));
    setLabel(d.label ?? "");
  }, []);

  const load = useCallback(async () => {
    const y = Number(year);
    const m = Number(month);
    if (!Number.isFinite(y) || !Number.isFinite(m)) {
      setError("年・月が不正です");
      return;
    }
    setLoading(true);
    setError(null);
    try {
      const res = await apiFetch<Payload>(
        `/api/admin/shift-schedule?${new URLSearchParams({ year: String(y), month: String(m) })}`,
      );
      applyPayload(res.data);
    } catch (e) {
      setPayload(null);
      setError(e instanceof Error ? e.message : "読込に失敗しました");
    } finally {
      setLoading(false);
    }
  }, [applyPayload, month, year]);

  useEffect(() => {
    void load();
  }, [load]);

  const setCell = useCallback((shano: string, dateIso: string, value: string) => {
    setGrid((g) => ({
      ...g,
      [shano]: { ...(g[shano] ?? {}), [dateIso]: value },
    }));
  }, []);

  const save = useCallback(async () => {
    if (!payload) return;
    const y = Number(year);
    const m = Number(month);
    setSaving(true);
    setError(null);
    try {
      await apiFetch<{ ok: boolean }>("/api/admin/shift-schedule", {
        method: "POST",
        body: JSON.stringify({
          year: y,
          month: m,
          label: label.trim() || undefined,
          grid,
        }),
      });
      await load();
    } catch (e) {
      setError(e instanceof Error ? e.message : "保存に失敗しました");
    } finally {
      setSaving(false);
    }
  }, [grid, label, load, month, payload, year]);

  const weekdays = useMemo(() => ["日", "月", "火", "水", "木", "金", "土"], []);

  return (
    <div className="space-y-4 p-4">
      <div>
        <h1 className="text-lg font-semibold tracking-tight">勤務表（ガンマ）</h1>
        <p className="text-muted-foreground mt-1 text-sm">
          期間は当月<strong>11日</strong>〜翌月<strong>10日</strong>です（Excel S_Day 準拠）。記号はコピー＆ペーストでセルに入力できます。
        </p>
      </div>

      <div className="flex flex-wrap items-end gap-3">
        <label className="flex flex-col gap-1 text-xs text-ink-muted">
          年
          <input
            value={year}
            onChange={(e) => setYear(e.target.value)}
            className="w-24 rounded-xl border border-cream-300 bg-white px-3 py-2 text-sm text-ink"
          />
        </label>
        <label className="flex flex-col gap-1 text-xs text-ink-muted">
          月（サイクル開始の「月」）
          <input
            value={month}
            onChange={(e) => setMonth(e.target.value)}
            className="w-20 rounded-xl border border-cream-300 bg-white px-3 py-2 text-sm text-ink"
          />
        </label>
        <Button type="button" size="sm" variant="outline" onClick={() => void load()} disabled={loading}>
          読込
        </Button>
        <Button type="button" size="sm" onClick={() => void save()} disabled={saving || !payload}>
          保存
        </Button>
      </div>

      <label className="flex max-w-md flex-col gap-1 text-xs text-ink-muted">
        メモラベル
        <input
          value={label}
          onChange={(e) => setLabel(e.target.value)}
          className="rounded-xl border border-cream-300 bg-white px-3 py-2 text-sm text-ink"
        />
      </label>

      {error ? <p className="text-destructive text-sm">{error}</p> : null}

      {payload ? (
        <p className="text-xs text-ink-muted">
          {payload.cycle.startIso} 〜 {payload.cycle.endIso}（{payload.cycle.dates.length} 日） / roster{" "}
          {payload.rosterSize || "未設定（全社員先頭表示）"}
        </p>
      ) : null}

      {payload ? (
        <div className="flex flex-wrap gap-2 rounded-xl border border-cream-200 bg-cream-50 p-3 text-xs">
          <span className="font-medium text-ink">記号チップ（参照）</span>
          {payload.symbols.slice(0, 24).map((s) => (
            <button
              key={s.symbol}
              type="button"
              className="rounded-lg border border-cream-300 bg-white px-2 py-0.5 hover:bg-cream-100"
              title={s.memo || s.symbol}
              onClick={() => void navigator.clipboard.writeText(s.symbol)}
            >
              {s.symbol}
            </button>
          ))}
          <span className="text-ink-muted">クリックでクリップボードにコピー</span>
        </div>
      ) : null}

      {payload ? (
        <div className="max-h-[calc(100vh-280px)] overflow-auto rounded-xl border border-cream-300 bg-white">
          <table className="min-w-max border-collapse text-[11px]">
            <thead className="bg-cream-100 sticky top-0 z-20">
              <tr>
                <th className="sticky left-0 z-30 border border-cream-300 bg-cream-100 px-2 py-1 text-left">社員</th>
                {payload.cycle.dates.map((d) => {
                  const wd = weekdays[new Date(d + "T12:00:00").getDay()];
                  return (
                    <th key={d} className="border border-cream-300 px-1 py-1 font-normal">
                      <div>{d.slice(8)}</div>
                      <div className="text-[10px] text-ink-muted">{wd}</div>
                    </th>
                  );
                })}
              </tr>
            </thead>
            <tbody>
              {payload.employees.map((emp) => (
                <tr key={emp.shano}>
                  <td className="sticky left-0 z-10 border border-cream-300 bg-cream-50 px-2 py-0.5 whitespace-nowrap">
                    <span className="font-mono">{emp.shano}</span> {emp.shaname}
                  </td>
                  {payload.cycle.dates.map((d) => (
                    <td key={d} className="border border-cream-200 p-0">
                      <input
                        className="h-7 w-11 border-0 bg-transparent px-0.5 text-center focus:bg-amber-50 focus:outline-none"
                        maxLength={4}
                        value={(grid[emp.shano] ?? {})[d] ?? ""}
                        onChange={(e) => setCell(emp.shano, d, e.target.value)}
                      />
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      ) : loading ? (
        <p className="text-sm text-ink-muted">読込中…</p>
      ) : null}
    </div>
  );
}
