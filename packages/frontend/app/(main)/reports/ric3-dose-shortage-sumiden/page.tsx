"use client";

import { useCallback, useEffect, useMemo, useState } from "react";

import { Button } from "@/components/ui/button";
import { apiFetch } from "@/lib/api";
import { cn } from "@/lib/utils";

type Ric3LayoutMeta = {
  variant: "standard" | "sumiden";
  reportTitle: string;
  formCode: string;
  employeeTableRowCap: number;
  additionalPassRounding: "simple_roundup" | "sumiden_margin_lt2";
};

type Ric3Payload = {
  layout: Ric3LayoutMeta;
  senkNoOptions: string[];
  sebkNoCount: number;
  employees: { shano: string; shaname: string }[];
  senkindTypes: { code: string; label: string }[];
  mokutekiTable: { code: string; description: string }[];
  detail: unknown;
};

type ApiShape = { data: Ric3Payload };

const EMPTY_PLAN = (): PlanRow => ({
  ukeNo: "",
  ricNoS: "",
  ricNoE: "",
  siteiSn: "",
  kagen: "",
  jyoug: "",
  jituP: "",
  kaiCd: "",
  kaiName: "",
});

type PlanRow = {
  ukeNo: string;
  ricNoS: string;
  ricNoE: string;
  siteiSn: string;
  kagen: string;
  jyoug: string;
  jituP: string;
  kaiCd: string;
  kaiName: string;
};

function parseNum(s: string | number | null | undefined): number | null {
  const t = String(s ?? "").trim();
  if (t === "") return null;
  const n = Number(t);
  return Number.isFinite(n) ? n : null;
}

/** バックエンド `computeRic3AdditionalPasses` と同じ */
function additionalPasses(
  rounding: Ric3LayoutMeta["additionalPassRounding"],
  shortage: number,
  measuredDose: number,
  actualPasses: number,
): number | null {
  if (actualPasses <= 0 || measuredDose <= 0) return null;
  const dosePerPass = measuredDose / actualPasses;
  if (dosePerPass <= 0) return null;
  const base = Math.ceil(shortage / dosePerPass - 1e-9);
  if (rounding === "simple_roundup") return base;
  const roundPortion = base * dosePerPass;
  if (roundPortion - shortage < 2) return base + 1;
  return base;
}

export default function Ric3DoseShortageSumidenPage() {
  const [payload, setPayload] = useState<Ric3Payload | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  const [tuikaN, setTuikaN] = useState("");
  const [senkNo, setSenkNo] = useState("");
  const [sokDate, setSokDate] = useState("");
  const [sokutCd, setSokutCd] = useState("");
  const [saiSoku, setSaiSoku] = useState("");
  const [sokutSn, setSokutSn] = useState("");
  const [tani, setTani] = useState("");
  const [shiteiPass, setShiteiPass] = useState("");
  const [sikiKodo, setSikiKodo] = useState("");
  const [sosi, setSosi] = useState("");
  const [atusa, setAtusa] = useState("");
  const [senKind, setSenKind] = useState("");
  const [atusaT, setAtusaT] = useState("");
  const [keisask, setKeisask] = useState("");

  const [planRows, setPlanRows] = useState<PlanRow[]>(() =>
    Array.from({ length: 5 }, () => EMPTY_PLAN()),
  );

  const load = useCallback(async () => {
    setLoading(true);
    setError(null);
    try {
      const params = new URLSearchParams();
      if (senkNo.trim()) params.set("senkNo", senkNo.trim());
      const q = params.toString();
      const res = await apiFetch<ApiShape>(`/api/reports/ric3-dose-shortage-sumiden${q ? `?${q}` : ""}`);
      setPayload(res.data);
    } catch (e) {
      setPayload(null);
      setError(e instanceof Error ? e.message : "読み込みに失敗しました");
    } finally {
      setLoading(false);
    }
  }, [senkNo]);

  useEffect(() => {
    let cancelled = false;
    void (async () => {
      setLoading(true);
      setError(null);
      try {
        const res = await apiFetch<ApiShape>("/api/reports/ric3-dose-shortage-sumiden");
        if (!cancelled) setPayload(res.data);
      } catch (e) {
        if (!cancelled) {
          setPayload(null);
          setError(e instanceof Error ? e.message : "読み込みに失敗しました");
        }
      } finally {
        if (!cancelled) setLoading(false);
      }
    })();
    return () => {
      cancelled = true;
    };
  }, []);

  const parsedMeasured = parseNum(sokutSn);
  const headerJituP = parseNum(shiteiPass);

  type DerivedRow = { shortage: number | null; addPass: number | null; margin: number | null };

  const derived = useMemo((): { rows: DerivedRow[] } => {
    if (!payload) return { rows: [] };
    const rounding = payload.layout.additionalPassRounding;
    const rows = planRows.map((r) => {
      const uke = r.ukeNo.trim();
      if (!uke) {
        return { shortage: null, addPass: null, margin: null };
      }
      const kagen = parseNum(r.kagen);
      const jyoug = parseNum(r.jyoug);
      const jituP = parseNum(r.jituP);
      if (parsedMeasured == null || kagen == null) {
        return { shortage: null, addPass: null, margin: null };
      }
      const shortage = kagen - parsedMeasured;
      const refPasses = jituP ?? headerJituP;
      const addPass =
        refPasses != null && refPasses > 0
          ? additionalPasses(rounding, shortage, parsedMeasured, refPasses)
          : null;
      const margin =
        jyoug != null ? jyoug - parsedMeasured : null;
      return { shortage, addPass, margin };
    });
    return { rows };
  }, [payload, planRows, parsedMeasured, headerJituP]);

  const m24 =
    derived.rows.some((x) => x.addPass != null && x.addPass < 0) ? "測定データが不適切" : "";
  const passVals = derived.rows.map((x) => x.addPass).filter((x): x is number => x != null);
  const s24 =
    passVals.length > 0 && Math.max(...passVals) !== Math.min(...passVals)
      ? "追加パス数異なるため課長に報告のこと"
      : "";

  const doseMeterNo =
    senkNo.trim().length >= 4 ? senkNo.trim().slice(-4) : senkNo.trim() ? senkNo.trim() : "—";

  return (
    <div className="mx-auto max-w-5xl space-y-6 p-4 md:p-6">
      <div>
        <h1 className="text-2xl font-semibold text-ink">
          {payload?.layout.reportTitle ?? "RIC3 線量不足報告書（住電）"}
        </h1>
        <p className="mt-1 text-sm text-ink-muted">
          様式 {payload?.layout.formCode ?? "G03-08"}。印刷範囲は Excel 帳票と同様 B1:AN30 を想定。マスタは API から取得し、追加パス数は住電用の端数ルールで試算します。
        </p>
      </div>

      <div className="flex flex-wrap items-center gap-3 rounded-2xl border border-cream-300 bg-cream-50 p-4">
        <Button type="button" variant="outline" onClick={() => void load()} disabled={loading}>
          {loading ? "再読込中…" : "マスタ再読込"}
        </Button>
        {payload ? (
          <span className="text-xs text-ink-muted">
            追加パス数:{" "}
            {payload.layout.additionalPassRounding === "sumiden_margin_lt2"
              ? "住電（余裕が 2 未満なら +1 パス）"
              : "通常（切り上げのみ）"}
          </span>
        ) : null}
      </div>

      {error ? <p className="text-sm text-red-700">{error}</p> : null}

      <section className="space-y-3 rounded-2xl border border-cream-300 bg-white p-4">
        <h2 className="text-sm font-semibold text-ink">ヘッダ・測定データ</h2>
        <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-3">
          <label className="flex flex-col gap-1 text-xs text-ink-muted">
            追加照射回目（Tuika）
            <input
              value={tuikaN}
              onChange={(e) => setTuikaN(e.target.value)}
              className="rounded-xl border border-cream-300 px-3 py-2 text-sm text-ink"
            />
          </label>
          <label className="flex flex-col gap-1 text-xs text-ink-muted">
            照射管理番号（SenkNo）
            {payload && payload.senkNoOptions.length > 0 ? (
              <select
                value={senkNo}
                onChange={(e) => setSenkNo(e.target.value)}
                className="rounded-xl border border-cream-300 px-3 py-2 text-sm text-ink"
              >
                <option value="">選択してください</option>
                {payload.senkNoOptions.map((s) => (
                  <option key={s} value={s}>
                    {s}
                  </option>
                ))}
              </select>
            ) : (
              <input
                value={senkNo}
                onChange={(e) => setSenkNo(e.target.value)}
                placeholder="手入力（DB 連携後は一覧化）"
                className="rounded-xl border border-cream-300 px-3 py-2 text-sm text-ink"
              />
            )}
          </label>
          <label className="flex flex-col gap-1 text-xs text-ink-muted">
            測定日
            <input
              value={sokDate}
              onChange={(e) => setSokDate(e.target.value)}
              className="rounded-xl border border-cream-300 px-3 py-2 text-sm text-ink"
            />
          </label>
          <label className="flex flex-col gap-1 text-xs text-ink-muted">
            測定者コード
            <input
              value={sokutCd}
              onChange={(e) => setSokutCd(e.target.value)}
              className="rounded-xl border border-cream-300 px-3 py-2 text-sm text-ink"
            />
          </label>
          <label className="flex flex-col gap-1 text-xs text-ink-muted">
            再測定者
            <input
              value={saiSoku}
              onChange={(e) => setSaiSoku(e.target.value)}
              className="rounded-xl border border-cream-300 px-3 py-2 text-sm text-ink"
            />
          </label>
          <label className="flex flex-col gap-1 text-xs text-ink-muted">
            測定線量（SokutSN）
            <input
              value={sokutSn}
              onChange={(e) => setSokutSn(e.target.value)}
              className="rounded-xl border border-cream-300 px-3 py-2 text-sm text-ink"
            />
          </label>
          <label className="flex flex-col gap-1 text-xs text-ink-muted">
            製品目的（単位）
            <input
              value={tani}
              onChange={(e) => setTani(e.target.value)}
              className="rounded-xl border border-cream-300 px-3 py-2 text-sm text-ink"
            />
          </label>
          <label className="flex flex-col gap-1 text-xs text-ink-muted">
            指定パス数（Pass）
            <input
              value={shiteiPass}
              onChange={(e) => setShiteiPass(e.target.value)}
              className="rounded-xl border border-cream-300 px-3 py-2 text-sm text-ink"
            />
          </label>
          <label className="flex flex-col gap-1 text-xs text-ink-muted">
            計算式コード（測定）
            <input
              value={sikiKodo}
              onChange={(e) => setSikiKodo(e.target.value)}
              className="rounded-xl border border-cream-300 px-3 py-2 text-sm text-ink"
            />
          </label>
          <label className="flex flex-col gap-1 text-xs text-ink-muted">
            素子異常
            <input
              value={sosi}
              onChange={(e) => setSosi(e.target.value)}
              className="rounded-xl border border-cream-300 px-3 py-2 text-sm text-ink"
            />
          </label>
          <label className="flex flex-col gap-1 text-xs text-ink-muted">
            素子厚さ
            <input
              value={atusa}
              onChange={(e) => setAtusa(e.target.value)}
              className="rounded-xl border border-cream-300 px-3 py-2 text-sm text-ink"
            />
          </label>
        </div>
      </section>

      <section className="space-y-3 rounded-2xl border border-cream-300 bg-white p-4">
        <h2 className="text-sm font-semibold text-ink">線量計（G15 相当）</h2>
        <p className="text-sm text-ink-muted">
          線量計番号（下4桁）: <span className="font-mono text-ink">{doseMeterNo}</span>
        </p>
        <div className="grid gap-3 sm:grid-cols-3">
          <label className="flex flex-col gap-1 text-xs text-ink-muted">
            線量計種類
            <select
              value={senKind}
              onChange={(e) => setSenKind(e.target.value)}
              className="rounded-xl border border-cream-300 px-3 py-2 text-sm text-ink"
            >
              <option value="">選択</option>
              {(payload?.senkindTypes ?? []).map((s) => (
                <option key={s.code} value={s.code}>
                  {s.label}
                </option>
              ))}
            </select>
          </label>
          <label className="flex flex-col gap-1 text-xs text-ink-muted">
            線量計厚さ（AtusaT）
            <input
              value={atusaT}
              onChange={(e) => setAtusaT(e.target.value)}
              className="rounded-xl border border-cream-300 px-3 py-2 text-sm text-ink"
            />
          </label>
          <label className="flex flex-col gap-1 text-xs text-ink-muted">
            計算式コード（Keisask）
            <input
              value={keisask}
              onChange={(e) => setKeisask(e.target.value)}
              className="rounded-xl border border-cream-300 px-3 py-2 text-sm text-ink"
            />
          </label>
        </div>
      </section>

      <section className="space-y-3 rounded-2xl border border-cream-300 bg-white p-4">
        <h2 className="text-sm font-semibold text-ink">照射計画（最大5受付・手入力）</h2>
        <div className="overflow-x-auto">
          <table className="w-full min-w-[720px] border-collapse text-left text-xs">
            <thead>
              <tr className="border-b border-cream-300 text-ink-muted">
                <th className="py-2 pr-2">受付</th>
                <th className="py-2 pr-2">RIC開始</th>
                <th className="py-2 pr-2">RIC終了</th>
                <th className="py-2 pr-2">指定線量</th>
                <th className="py-2 pr-2">下限</th>
                <th className="py-2 pr-2">上限</th>
                <th className="py-2 pr-2">実パス</th>
                <th className="py-2 pr-2">会社CD</th>
                <th className="py-2 pr-2">会社名</th>
              </tr>
            </thead>
            <tbody>
              {planRows.map((row, i) => (
                <tr key={i} className="border-b border-cream-200">
                  {(
                    [
                      "ukeNo",
                      "ricNoS",
                      "ricNoE",
                      "siteiSn",
                      "kagen",
                      "jyoug",
                      "jituP",
                      "kaiCd",
                      "kaiName",
                    ] as const
                  ).map((key) => (
                    <td key={key} className="py-1 pr-1">
                      <input
                        value={row[key]}
                        onChange={(e) => {
                          const v = e.target.value;
                          setPlanRows((prev) => {
                            const next = [...prev];
                            next[i] = { ...next[i], [key]: v };
                            return next;
                          });
                        }}
                        className="w-full min-w-[4rem] rounded-lg border border-cream-300 px-2 py-1 text-ink"
                      />
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </section>

      <section className="space-y-3 rounded-2xl border border-cream-300 bg-cream-50 p-4">
        <h2 className="text-sm font-semibold text-ink">追加照射試算（O18 / AA18 相当）</h2>
        {(m24 || s24) && (
          <div className="space-y-1 text-sm text-amber-900">
            {m24 ? <p>{m24}</p> : null}
            {s24 ? <p>{s24}</p> : null}
          </div>
        )}
        <ul className="space-y-2 text-sm text-ink">
          {planRows.map((r, i) => {
            const d = derived.rows[i];
            if (!r.ukeNo.trim()) return null;
            return (
              <li
                key={i}
                className={cn(
                  "rounded-xl border border-cream-300 bg-white px-3 py-2",
                  d.addPass != null && d.addPass < 0 && "border-red-300",
                )}
              >
                受付 {r.ukeNo}: 線量不足量{" "}
                <span className="font-mono">{d.shortage != null ? d.shortage.toFixed(2) : "—"}</span>
                {" · "}
                上限余裕{" "}
                <span className="font-mono">{d.margin != null ? d.margin.toFixed(2) : "—"}</span>
                {" · "}
                追加パス数{" "}
                <span className="font-mono font-semibold">{d.addPass != null ? d.addPass : "—"}</span>
              </li>
            );
          })}
        </ul>
        <p className="text-xs text-ink-muted">
          実パスが行に無い場合はヘッダの指定パス数を参照します（仕様の AA9 相当）。
        </p>
      </section>

      <section className="rounded-2xl border border-cream-300 bg-white p-4">
        <h2 className="text-sm font-semibold text-ink">製品目的マスタ（mokuteki）</h2>
        <ul className="mt-2 grid gap-2 sm:grid-cols-2 text-sm text-ink">
          {(payload?.mokutekiTable ?? []).map((m) => (
            <li key={m.code} className="rounded-lg bg-cream-50 px-3 py-2">
              <span className="font-mono">{m.code}</span> — {m.description}
            </li>
          ))}
        </ul>
      </section>

      <section className="rounded-2xl border border-cream-300 bg-white p-4">
        <h2 className="text-sm font-semibold text-ink">
          測定資格者一覧（SyainTB・最大 {payload?.layout.employeeTableRowCap ?? "—"} 名）
        </h2>
        <ul className="mt-2 max-h-48 overflow-auto text-sm text-ink">
          {(payload?.employees ?? []).map((e) => (
            <li key={e.shano}>
              <span className="font-mono text-xs text-ink-muted">{e.shano}</span> {e.shaname}
            </li>
          ))}
        </ul>
      </section>

      <p className="text-xs text-ink-muted">本紙は照射計画作成後作業指図書に貼付すること（Excel 注記 B30 相当）。</p>
    </div>
  );
}
