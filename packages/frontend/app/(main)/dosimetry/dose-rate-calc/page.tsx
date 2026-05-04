"use client";

import { useCallback, useEffect, useState } from "react";

import { Button } from "@/components/ui/button";
import { apiFetch, apiUrl } from "@/lib/api";

type CoefficientsResponse = {
  version: string;
  specRef: string;
  constants: {
    massAbsorptionCoWater: number;
    massAbsorptionCsAir: number;
    csToCoWater: number;
    cableFactorWhenEnabled: number;
    potentiometerUncertaintyPct: number;
    paToNcPerH: number;
  };
  potentiometerOptions: { id: number; label: string; rateConst: number; integrConst: number }[];
  ionChambers: { id: number; model: string; calConstantGyPerNc: number }[];
  kanriSources: { sikibetu: string; kousinn: string | null; ricvm: string | null; hppp: string | null }[];
};

type PostResult = {
  ok: boolean;
  result: {
    kappaTp: number;
    cableFactor: number;
    potentiometerConst: number;
    ionChamberConst: number;
    uncertaintyFactor: number;
    uncertaintyCombinedPct: number;
    rawProduct: number;
    displayDoseRateGyPerH: number | null;
    displayIntegratedDoseGy: number | null;
    decayFactor: number;
    displayDoseRateAfterDecay: number | null;
    displayIntegratedDoseAfterDecay: number | null;
    exposureTimeH: number | null;
  };
};

const defaultForm = {
  mode: "RATE" as "RATE" | "INTEG",
  potentiometerId: 1 as 1 | 2,
  ionChamberId: 1,
  readValue: 10,
  temperatureC: 22,
  pressureHpa: 1013.3,
  tpEnabled: true,
  uncertaintyOrg: "none" as "ANTM" | "JQA" | "none",
  cableOn: true,
  targetDoseGy: "" as string,
  useDecay: false,
  halfLifeDays: 30,
  elapsedDays: 0,
};

export default function DoseRateCalcPage() {
  const [coef, setCoef] = useState<CoefficientsResponse | null>(null);
  const [loadError, setLoadError] = useState<string | null>(null);
  const [form, setForm] = useState(defaultForm);
  const [out, setOut] = useState<PostResult["result"] | null>(null);
  const [apiError, setApiError] = useState<string | null>(null);
  const [pending, setPending] = useState(false);

  useEffect(() => {
    setLoadError(null);
    void apiFetch<CoefficientsResponse>("/api/dosimetry/dose-rate-calc")
      .then(setCoef)
      .catch((e: Error) => setLoadError(e.message));
  }, []);

  const runCalc = useCallback(async () => {
    setApiError(null);
    setPending(true);
    const td = form.targetDoseGy.trim() ? Number.parseFloat(form.targetDoseGy) : undefined;
    const body = {
      mode: form.mode,
      potentiometerId: form.potentiometerId,
      ionChamberId: form.ionChamberId,
      readValue: form.readValue,
      temperatureC: form.temperatureC,
      pressureHpa: form.pressureHpa,
      tpEnabled: form.tpEnabled,
      uncertaintyOrg: form.uncertaintyOrg,
      cableOn: form.cableOn,
      ...(td != null && !Number.isNaN(td) && td >= 0 ? { targetDoseGy: td } : {}),
      ...(form.useDecay
        ? { decay: { halfLifeDays: form.halfLifeDays, elapsedDays: form.elapsedDays } }
        : {}),
    };
    try {
      const r = await apiFetch<PostResult>("/api/dosimetry/dose-rate-calc", {
        method: "POST",
        body: JSON.stringify(body),
      });
      if (!r.ok || !r.result) {
        setApiError("計算に失敗しました。");
        return;
      }
      setOut(r.result);
    } catch (e) {
      setApiError(e instanceof Error ? e.message : "予期せぬエラー");
    } finally {
      setPending(false);
    }
  }, [form]);

  return (
    <section className="space-y-6 text-ink">
      <header className="space-y-1">
        <h1 className="text-2xl font-semibold tracking-tight">線量率計算</h1>
        <p className="text-sm text-ink-muted">
          仕様 v1.1 相当（MVP）— 係数は API{" "}
          <code className="rounded bg-cream-200 px-1.5 py-0.5 text-xs">GET {apiUrl("/api/dosimetry/dose-rate-calc")}</code>{" "}
          から取得し、POST で計算します。
        </p>
      </header>

      {loadError && (
        <p className="rounded-xl border border-terracotta/30 bg-terracotta/5 px-4 py-2 text-sm text-terracotta-dark">
          係数の取得に失敗しました: {loadError}（API の起動を確認してください）
        </p>
      )}

      <div className="grid gap-6 lg:grid-cols-2">
        <div className="rounded-xl border border-cream-300 bg-cream-50/80 p-5 shadow-sm">
          <h2 className="mb-4 text-sm font-semibold text-ink">入力</h2>
          <div className="space-y-3 text-sm">
            <Field label="モード">
              <select
                className="w-full rounded-lg border border-cream-300 bg-white px-3 py-2"
                value={form.mode}
                onChange={(e) => setForm((f) => ({ ...f, mode: e.target.value as "RATE" | "INTEG" }))}
              >
                <option value="RATE">線量率 (Gy/h)</option>
                <option value="INTEG">積算線量 (Gy)</option>
              </select>
            </Field>
            <Field label="電位計">
              <select
                className="w-full rounded-lg border border-cream-300 bg-white px-3 py-2"
                value={form.potentiometerId}
                onChange={(e) => setForm((f) => ({ ...f, potentiometerId: Number(e.target.value) as 1 | 2 }))}
              >
                {coef?.potentiometerOptions.map((p) => (
                  <option key={p.id} value={p.id}>
                    {p.label}
                    {form.mode === "RATE" ? ` (RATE ${p.rateConst})` : ` (INTEG ${p.integrConst})`}
                  </option>
                )) ?? (
                  <>
                    <option value={1}>#3302229</option>
                    <option value={2}>#4502260</option>
                  </>
                )}
              </select>
            </Field>
            <Field label="電離箱">
              <select
                className="w-full rounded-lg border border-cream-300 bg-white px-3 py-2"
                value={form.ionChamberId}
                onChange={(e) => setForm((f) => ({ ...f, ionChamberId: Number(e.target.value) }))}
              >
                {coef?.ionChambers.map((c) => (
                  <option key={c.id} value={c.id}>
                    {c.model}（{c.calConstantGyPerNc} Gy/nC）
                  </option>
                )) ?? <option value={1}>#1253</option>}
              </select>
            </Field>
            <Field label={form.mode === "RATE" ? "読み値 (pA)" : "読み値 (nC)"}>
              <input
                type="number"
                className="w-full rounded-lg border border-cream-300 bg-white px-3 py-2"
                value={form.readValue}
                onChange={(e) => setForm((f) => ({ ...f, readValue: Number.parseFloat(e.target.value) || 0 }))}
                step="any"
              />
            </Field>
            <div className="grid grid-cols-2 gap-3">
              <Field label="温度 (℃)">
                <input
                  type="number"
                  className="w-full rounded-lg border border-cream-300 bg-white px-3 py-2"
                  value={form.temperatureC}
                  onChange={(e) => setForm((f) => ({ ...f, temperatureC: Number.parseFloat(e.target.value) || 0 }))}
                />
              </Field>
              <Field label="気圧 (hPa)">
                <input
                  type="number"
                  className="w-full rounded-lg border border-cream-300 bg-white px-3 py-2"
                  value={form.pressureHpa}
                  onChange={(e) => setForm((f) => ({ ...f, pressureHpa: Number.parseFloat(e.target.value) || 0 }))}
                />
              </Field>
            </div>
            <Field label="温度・気圧補正">
              <label className="flex items-center gap-2">
                <input
                  type="checkbox"
                  checked={form.tpEnabled}
                  onChange={(e) => setForm((f) => ({ ...f, tpEnabled: e.target.checked }))}
                />
                <span>κTP を適用</span>
              </label>
            </Field>
            <Field label="不確かさ（機関）">
              <select
                className="w-full rounded-lg border border-cream-300 bg-white px-3 py-2"
                value={form.uncertaintyOrg}
                onChange={(e) => setForm((f) => ({ ...f, uncertaintyOrg: e.target.value as typeof form.uncertaintyOrg }))}
              >
                <option value="ANTM">ANTM</option>
                <option value="JQA">JQA</option>
                <option value="none">なし</option>
              </select>
            </Field>
            <Field label="ケーブル補正">
              <label className="flex items-center gap-2">
                <input
                  type="checkbox"
                  checked={form.cableOn}
                  onChange={(e) => setForm((f) => ({ ...f, cableOn: e.target.checked }))}
                />
                <span>あり（Co 空気/Co 水比 ≈ 0.8991）</span>
              </label>
            </Field>
            {form.mode === "RATE" && (
              <Field label="設定線量 (Gy) — 照射時間">
                <input
                  type="text"
                  inputMode="decimal"
                  className="w-full rounded-lg border border-cream-300 bg-white px-3 py-2"
                  placeholder="空欄可"
                  value={form.targetDoseGy}
                  onChange={(e) => setForm((f) => ({ ...f, targetDoseGy: e.target.value }))}
                />
              </Field>
            )}
            <div className="rounded-lg border border-cream-300 bg-cream-100/50 p-3">
              <label className="mb-2 flex items-center gap-2 text-xs font-medium">
                <input
                  type="checkbox"
                  checked={form.useDecay}
                  onChange={(e) => setForm((f) => ({ ...f, useDecay: e.target.checked }))}
                />
                減衰補正（半減期・経過日数）
              </label>
              {form.useDecay && (
                <div className="grid grid-cols-2 gap-2">
                  <Field label="半減期 (日)">
                    <input
                      type="number"
                      className="w-full rounded-lg border border-cream-300 bg-white px-2 py-1.5 text-sm"
                      value={form.halfLifeDays}
                      onChange={(e) => setForm((f) => ({ ...f, halfLifeDays: Number.parseFloat(e.target.value) || 1 }))}
                    />
                  </Field>
                  <Field label="経過 (日)">
                    <input
                      type="number"
                      className="w-full rounded-lg border border-cream-300 bg-white px-2 py-1.5 text-sm"
                      value={form.elapsedDays}
                      onChange={(e) => setForm((f) => ({ ...f, elapsedDays: Number.parseFloat(e.target.value) || 0 }))}
                    />
                  </Field>
                </div>
              )}
            </div>
          </div>
          <div className="mt-5">
            <Button
              className="w-full bg-terracotta text-white hover:bg-terracotta-dark"
              onClick={() => void runCalc()}
              disabled={pending}
            >
              {pending ? "計算中…" : "計算"}
            </Button>
          </div>
          {apiError && <p className="mt-2 text-sm text-terracotta-dark">{apiError}</p>}
        </div>

        <div className="rounded-xl border border-cream-300 bg-cream-50/80 p-5 shadow-sm">
          <h2 className="mb-4 text-sm font-semibold text-ink">結果</h2>
          {!out && <p className="text-sm text-ink-muted">「計算」で表示されます。</p>}
          {out && (
            <dl className="space-y-2 text-sm">
              {out.displayDoseRateGyPerH != null && (
                <>
                  <div className="flex justify-between gap-2 border-b border-cream-200 py-1">
                    <dt>線量率（表示）</dt>
                    <dd className="font-mono text-right">{out.displayDoseRateGyPerH} Gy/h</dd>
                  </div>
                  {out.displayDoseRateAfterDecay != null && out.decayFactor !== 1 && (
                    <div className="flex justify-between gap-2 border-b border-cream-200 py-1">
                      <dt>減衰後</dt>
                      <dd className="font-mono text-right">
                        {out.displayDoseRateAfterDecay} Gy/h（×{out.decayFactor.toFixed(4)}）
                      </dd>
                    </div>
                  )}
                </>
              )}
              {out.displayIntegratedDoseGy != null && (
                <>
                  <div className="flex justify-between gap-2 border-b border-cream-200 py-1">
                    <dt>積算線量（表示）</dt>
                    <dd className="font-mono text-right">{out.displayIntegratedDoseGy} Gy</dd>
                  </div>
                  {out.displayIntegratedDoseAfterDecay != null && out.decayFactor !== 1 && (
                    <div className="flex justify-between gap-2 border-b border-cream-200 py-1">
                      <dt>減衰後</dt>
                      <dd className="font-mono text-right">
                        {out.displayIntegratedDoseAfterDecay} Gy（×{out.decayFactor.toFixed(4)}）
                      </dd>
                    </div>
                  )}
                </>
              )}
              <div className="flex justify-between gap-2 border-b border-cream-200 py-1 text-ink-muted">
                <dt>κTP / ケーブル / 不確かさ</dt>
                <dd className="text-right text-xs">
                  {out.kappaTp} / {out.cableFactor} / {out.uncertaintyFactor}
                </dd>
              </div>
              {out.exposureTimeH != null && (
                <div className="flex justify-between gap-2 border-b border-cream-200 py-1">
                  <dt>照射時間（h・切り上げ2桁）</dt>
                  <dd className="font-mono text-right">{out.exposureTimeH}</dd>
                </div>
              )}
            </dl>
          )}
        </div>
      </div>

      {coef && (
        <div className="rounded-xl border border-cream-300 bg-cream-50/60 p-4 text-xs text-ink-muted">
          <p>
            バージョン {coef.version} — {coef.specRef} — 換算: Cs→Co 水 = {coef.constants.csToCoWater.toFixed(5)}
          </p>
          {coef.kanriSources.length > 0 && (
            <p className="mt-1">
              ex_kanri_tb: {coef.kanriSources.length} 行（表示のみ・減衰パラメータは手入力）
            </p>
          )}
        </div>
      )}
    </section>
  );
}

function Field({ label, children }: { label: string; children: React.ReactNode }) {
  return (
    <div>
      <p className="mb-1 text-xs font-medium text-ink-muted">{label}</p>
      {children}
    </div>
  );
}
