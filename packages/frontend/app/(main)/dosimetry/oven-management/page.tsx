"use client";

import { useCallback, useEffect, useMemo, useState } from "react";

import { Button } from "@/components/ui/button";
import { apiFetch } from "@/lib/api";
import { cn } from "@/lib/utils";

type OvenRowJson = {
  placeName?: string;
  startDate?: string;
  startTimer?: string;
  heatTemp?: string;
  ratioToKanri?: string;
  endDate?: string;
  endTimer?: string;
  note?: string;
};

type ExOvenRow = { ovenkey: string; json: unknown };

type Syouj3ovRow = {
  ovno: string;
  uno: string | null;
  ovenId: string | null;
  result: string | null;
  bikou: string | null;
};

type OvenManagementPayload = {
  exOvenTb: ExOvenRow[];
  syouj3ov: Syouj3ovRow[];
};

function parseOvenJson(raw: unknown): OvenRowJson {
  if (raw == null || typeof raw !== "object" || Array.isArray(raw)) return {};
  return raw as OvenRowJson;
}

function cell(s: string | undefined): string {
  const t = (s ?? "").trim();
  return t.length > 0 ? t : "—";
}

const inputClass =
  "w-full min-w-[5rem] rounded-lg border border-cream-300 bg-cream-50 px-2 py-1.5 text-sm text-ink outline-none focus:border-terracotta focus:ring-1 focus:ring-terracotta/30";

export default function DosimetryOvenManagementPage() {
  const [payload, setPayload] = useState<OvenManagementPayload | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [busyKey, setBusyKey] = useState<string | null>(null);

  const [sokutei, setSokutei] = useState("");
  const [atusa, setAtusa] = useState("");
  const [ondo, setOndo] = useState("");
  const [ratioForOven, setRatioForOven] = useState("");
  const [preview, setPreview] = useState<{
    absPerThickness: number;
    correctedKanriPoint: number;
    correctedOvenDose?: number;
  } | null>(null);

  const [newOvenkey, setNewOvenkey] = useState("");
  const [newJson, setNewJson] = useState<OvenRowJson>({});

  const load = useCallback(async () => {
    setLoading(true);
    setError(null);
    try {
      const res = await apiFetch<{ data: OvenManagementPayload }>("/api/dosimetry/oven-management");
      setPayload(res.data);
    } catch (e) {
      setPayload(null);
      setError(e instanceof Error ? e.message : "取得に失敗しました");
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    void load();
  }, [load]);

  const runPreview = useCallback(async () => {
    setError(null);
    const s = Number(sokutei);
    const a = Number(atusa);
    const t = Number(ondo);
    if (!Number.isFinite(s) || !Number.isFinite(a) || !Number.isFinite(t)) {
      setError("測定値・厚さ・温度を数値で入力してください");
      setPreview(null);
      return;
    }
    if (a === 0) {
      setError("厚さが0です");
      setPreview(null);
      return;
    }
    try {
      const q = new URLSearchParams({
        sokutei: String(s),
        atusa: String(a),
        ondo: String(t),
      });
      const res = await apiFetch<{
        data: { absPerThickness: number; correctedKanriPoint: number };
      }>(`/api/dosimetry/oven-management/correction-preview?${q.toString()}`);
      const p = Number(ratioForOven);
      const correctedOvenDose =
        Number.isFinite(p) && ratioForOven.trim() !== ""
          ? p * res.data.correctedKanriPoint
          : undefined;
      setPreview({
        absPerThickness: res.data.absPerThickness,
        correctedKanriPoint: res.data.correctedKanriPoint,
        correctedOvenDose,
      });
    } catch (e) {
      setPreview(null);
      setError(e instanceof Error ? e.message : "補正計算に失敗しました");
    }
  }, [sokutei, atusa, ondo, ratioForOven]);

  const saveNew = useCallback(async () => {
    const key = newOvenkey.trim();
    if (!key) {
      setError("ovenkey を入力してください");
      return;
    }
    setBusyKey("new");
    setError(null);
    try {
      await apiFetch("/api/dosimetry/oven-management", {
        method: "POST",
        body: JSON.stringify({ ovenkey: key, json: newJson }),
      });
      setNewOvenkey("");
      setNewJson({});
      await load();
    } catch (e) {
      setError(e instanceof Error ? e.message : "登録に失敗しました");
    } finally {
      setBusyKey(null);
    }
  }, [newOvenkey, newJson, load]);

  const patchOven = useCallback(
    async (ovenkey: string, json: OvenRowJson) => {
      setBusyKey(ovenkey);
      setError(null);
      try {
        await apiFetch(`/api/dosimetry/oven-management/${encodeURIComponent(ovenkey)}`, {
          method: "PATCH",
          body: JSON.stringify({ json }),
        });
        await load();
      } catch (e) {
        setError(e instanceof Error ? e.message : "更新に失敗しました");
      } finally {
        setBusyKey(null);
      }
    },
    [load],
  );

  const rows = payload?.exOvenTb ?? [];

  const title = useMemo(
    () => (
      <div className="mb-6">
        <h1 className="text-xl font-semibold text-ink">R3オーブン管理（温度補正）</h1>
        <p className="mt-1 text-sm text-ink-muted">
          3号機オーブンの使用状況と、管理点線量シート相当の温度補正（仕様:{" "}
          <code className="rounded bg-cream-200 px-1 text-xs">
            (-3.2478+23.386×ABS/厚さ)×(-0.006×温度+1.0558)
          </code>
          ）
        </p>
      </div>
    ),
    [],
  );

  return (
    <div className="mx-auto max-w-7xl">
      {title}

      {error ? (
        <div className="mb-4 rounded-xl border border-red-200 bg-red-50 px-4 py-3 text-sm text-red-800">{error}</div>
      ) : null}

      <section className="mb-8 rounded-xl border border-cream-300 bg-cream-50/80 p-4">
        <h2 className="mb-3 text-sm font-semibold text-ink">温度補正（管理点）</h2>
        <div className="flex flex-wrap items-end gap-3">
          <label className="block min-w-[8rem]">
            <span className="mb-1 block text-xs text-ink-muted">測定値 ABS</span>
            <input
              className={inputClass}
              value={sokutei}
              onChange={(e) => setSokutei(e.target.value)}
              inputMode="decimal"
              placeholder="H"
            />
          </label>
          <label className="block min-w-[8rem]">
            <span className="mb-1 block text-xs text-ink-muted">素子厚さ (mm)</span>
            <input
              className={inputClass}
              value={atusa}
              onChange={(e) => setAtusa(e.target.value)}
              inputMode="decimal"
              placeholder="G"
            />
          </label>
          <label className="block min-w-[8rem]">
            <span className="mb-1 block text-xs text-ink-muted">測定温度 (℃)</span>
            <input
              className={inputClass}
              value={ondo}
              onChange={(e) => setOndo(e.target.value)}
              inputMode="decimal"
              placeholder="K"
            />
          </label>
          <label className="block min-w-[8rem]">
            <span className="mb-1 block text-xs text-ink-muted">管理点との比 P（任意・AH列）</span>
            <input
              className={inputClass}
              value={ratioForOven}
              onChange={(e) => setRatioForOven(e.target.value)}
              inputMode="decimal"
              placeholder="P"
            />
          </label>
          <Button type="button" variant="outline" size="sm" onClick={() => void runPreview()}>
            補正を計算
          </Button>
        </div>
        {preview ? (
          <dl className="mt-4 grid gap-2 text-sm sm:grid-cols-2 lg:grid-cols-3">
            <div className="rounded-lg border border-cream-200 bg-cream-100/80 px-3 py-2">
              <dt className="text-xs text-ink-muted">ABS / 厚さ</dt>
              <dd className="font-mono text-ink">{preview.absPerThickness.toFixed(6)}</dd>
            </div>
            <div className="rounded-lg border border-cream-200 bg-cream-100/80 px-3 py-2">
              <dt className="text-xs text-ink-muted">補正後管理点（AG列相当）</dt>
              <dd className="font-mono text-ink">{preview.correctedKanriPoint.toFixed(6)}</dd>
            </div>
            {preview.correctedOvenDose != null ? (
              <div className="rounded-lg border border-cream-200 bg-cream-100/80 px-3 py-2">
                <dt className="text-xs text-ink-muted">補正後オーブン線量 P×AG（AH列相当）</dt>
                <dd className="font-mono text-ink">{preview.correctedOvenDose.toFixed(6)}</dd>
              </div>
            ) : null}
          </dl>
        ) : null}
      </section>

      <section className="mb-8">
        <div className="mb-2 flex items-center justify-between gap-2">
          <h2 className="text-sm font-semibold text-ink">オーブン一覧（ExOvenTb）</h2>
          <Button type="button" variant="outline" size="sm" disabled={loading} onClick={() => void load()}>
            再読込
          </Button>
        </div>
        <div className="overflow-x-auto rounded-xl border border-cream-300 bg-cream-50/50">
          <table className="w-full min-w-[960px] border-collapse text-left text-sm">
            <thead>
              <tr className="border-b border-cream-300 bg-cream-100/80 text-xs text-ink-muted">
                <th className="px-2 py-2 font-medium">ovenkey</th>
                <th className="px-2 py-2 font-medium">場所</th>
                <th className="px-2 py-2 font-medium">使用開始日</th>
                <th className="px-2 py-2 font-medium">開始ﾀｲﾏｰ</th>
                <th className="px-2 py-2 font-medium">加熱温度</th>
                <th className="px-2 py-2 font-medium">管理点比</th>
                <th className="px-2 py-2 font-medium">終了日</th>
                <th className="px-2 py-2 font-medium">終了ﾀｲﾏｰ</th>
                <th className="px-2 py-2 font-medium">備考</th>
                <th className="px-2 py-2 font-medium"> </th>
              </tr>
            </thead>
            <tbody>
              {loading && rows.length === 0 ? (
                <tr>
                  <td colSpan={10} className="px-3 py-6 text-center text-ink-muted">
                    読み込み中…
                  </td>
                </tr>
              ) : null}
              {!loading && rows.length === 0 ? (
                <tr>
                  <td colSpan={10} className="px-3 py-6 text-center text-ink-muted">
                    行がありません。下のフォームから登録できます。
                  </td>
                </tr>
              ) : null}
              {rows.map((row) => (
                <OvenRowEditor
                  key={row.ovenkey}
                  row={row}
                  busy={busyKey === row.ovenkey}
                  onSave={(json) => void patchOven(row.ovenkey, json)}
                />
              ))}
            </tbody>
          </table>
        </div>
      </section>

      <section className="mb-8 rounded-xl border border-cream-300 bg-cream-50/80 p-4">
        <h2 className="mb-3 text-sm font-semibold text-ink">オーブンを新規登録（POST）</h2>
        <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-3">
          <label className="block sm:col-span-2">
            <span className="mb-1 block text-xs text-ink-muted">ovenkey（一意・64文字以内）</span>
            <input
              className={inputClass}
              value={newOvenkey}
              onChange={(e) => setNewOvenkey(e.target.value)}
              placeholder="例: r3-oven-001"
            />
          </label>
        </div>
        <OvenFields
          value={newJson}
          onChange={setNewJson}
        />
        <div className="mt-3">
          <Button type="button" size="sm" disabled={busyKey === "new"} onClick={() => void saveNew()}>
            {busyKey === "new" ? "登録中…" : "登録"}
          </Button>
        </div>
      </section>

      <section>
        <h2 className="mb-2 text-sm font-semibold text-ink">3号機オーブン実績（syouj3ov・参照）</h2>
        <div className="overflow-x-auto rounded-xl border border-cream-300 bg-cream-50/50">
          <table className="w-full min-w-[640px] border-collapse text-left text-sm">
            <thead>
              <tr className="border-b border-cream-300 bg-cream-100/80 text-xs text-ink-muted">
                <th className="px-2 py-2 font-medium">ovno</th>
                <th className="px-2 py-2 font-medium">uno</th>
                <th className="px-2 py-2 font-medium">oven_id</th>
                <th className="px-2 py-2 font-medium">result</th>
                <th className="px-2 py-2 font-medium">bikou</th>
              </tr>
            </thead>
            <tbody>
              {(payload?.syouj3ov ?? []).slice(0, 200).map((r) => (
                <tr key={r.ovno} className="border-b border-cream-200/80">
                  <td className="px-2 py-1.5 font-mono text-xs">{cell(r.ovno)}</td>
                  <td className="px-2 py-1.5">{cell(r.uno ?? undefined)}</td>
                  <td className="px-2 py-1.5">{cell(r.ovenId ?? undefined)}</td>
                  <td className="px-2 py-1.5">{cell(r.result ?? undefined)}</td>
                  <td className="px-2 py-1.5">{cell(r.bikou ?? undefined)}</td>
                </tr>
              ))}
              {(!payload?.syouj3ov || payload.syouj3ov.length === 0) && !loading ? (
                <tr>
                  <td colSpan={5} className="px-3 py-4 text-center text-ink-muted">
                    データがありません
                  </td>
                </tr>
              ) : null}
            </tbody>
          </table>
        </div>
      </section>
    </div>
  );
}

function OvenFields({
  value,
  onChange,
}: {
  value: OvenRowJson;
  onChange: (v: OvenRowJson) => void;
}) {
  const set = (key: keyof OvenRowJson, v: string) => {
    onChange({ ...value, [key]: v });
  };
  return (
    <div className="mt-3 grid gap-3 sm:grid-cols-2 lg:grid-cols-4">
      {(
        [
          ["placeName", "オーブン場所"],
          ["startDate", "使用開始日"],
          ["startTimer", "使用開始タイマー"],
          ["heatTemp", "加熱温度(℃)"],
          ["ratioToKanri", "線量管理点との比"],
          ["endDate", "使用終了日"],
          ["endTimer", "使用終了タイマー"],
          ["note", "備考"],
        ] as const
      ).map(([key, label]) => (
        <label key={key} className="block">
          <span className="mb-1 block text-xs text-ink-muted">{label}</span>
          <input className={inputClass} value={value[key] ?? ""} onChange={(e) => set(key, e.target.value)} />
        </label>
      ))}
    </div>
  );
}

function OvenRowEditor({
  row,
  busy,
  onSave,
}: {
  row: ExOvenRow;
  busy: boolean;
  onSave: (json: OvenRowJson) => void;
}) {
  const initial = parseOvenJson(row.json);
  const [draft, setDraft] = useState<OvenRowJson>(initial);

  useEffect(() => {
    setDraft(parseOvenJson(row.json));
  }, [row.ovenkey, row.json]);

  const dirty = useMemo(() => JSON.stringify(draft) !== JSON.stringify(initial), [draft, initial]);

  return (
    <tr className="border-b border-cream-200/80">
      <td className="px-2 py-1.5 align-top font-mono text-xs text-ink">{row.ovenkey}</td>
      {(
        [
          "placeName",
          "startDate",
          "startTimer",
          "heatTemp",
          "ratioToKanri",
          "endDate",
          "endTimer",
          "note",
        ] as const
      ).map((k) => (
        <td key={k} className="px-1 py-1 align-top">
          <input
            className={cn(inputClass, "text-xs")}
            value={draft[k] ?? ""}
            onChange={(e) => setDraft((d) => ({ ...d, [k]: e.target.value }))}
          />
        </td>
      ))}
      <td className="px-1 py-1 align-top">
        <Button
          type="button"
          size="sm"
          variant="outline"
          disabled={busy || !dirty}
          onClick={() => onSave(draft)}
        >
          {busy ? "…" : "保存"}
        </Button>
      </td>
    </tr>
  );
}
