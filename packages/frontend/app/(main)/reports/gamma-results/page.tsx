"use client";

import { type ChangeEvent, useCallback, useState } from "react";

import { Button } from "@/components/ui/button";
import { apiFetch } from "@/lib/api";

type AggregateResponse = {
  data: {
    params: {
      year: number;
      month: number;
      unoMin: string;
      unoMax: string;
      kkmd1: string;
      kkmd2: string;
    };
    goukei: {
      ric1: { inboundQty: number; simpleProcessingApprox: number };
      ric2: { inboundQty: number; simpleProcessingApprox: number };
      ric3: { inboundQty: number; simpleProcessingApprox: number };
    };
    notes: string[];
  };
};

export default function GammaResultsReportPage() {
  const now = new Date();
  const [year, setYear] = useState(String(now.getFullYear()));
  const [month, setMonth] = useState(String(now.getMonth() + 1));
  const [unoMin, setUnoMin] = useState("");
  const [unoMax, setUnoMax] = useState("");
  const [data, setData] = useState<AggregateResponse["data"] | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [loading, setLoading] = useState(false);

  const run = useCallback(async () => {
    const y = Number(year);
    const m = Number(month);
    if (!Number.isFinite(y) || !Number.isFinite(m) || m < 1 || m > 12) {
      setError("年・月が不正です");
      return;
    }
    setLoading(true);
    setError(null);
    try {
      const body: Record<string, unknown> = { year: y, month: m };
      const u1 = unoMin.trim();
      const u2 = unoMax.trim();
      if (u1) body.unoMin = u1;
      if (u2) body.unoMax = u2;

      const res = await apiFetch<AggregateResponse>("/api/reports/gamma-results/aggregate", {
        method: "POST",
        body: JSON.stringify(body),
      });
      setData(res.data);
    } catch (e) {
      setData(null);
      setError(e instanceof Error ? e.message : "集計に失敗しました");
    } finally {
      setLoading(false);
    }
  }, [month, unoMax, unoMin, year]);

  return (
    <div className="space-y-4 p-4">
      <div>
        <h1 className="text-lg font-semibold tracking-tight">ガンマ照射課実績集計</h1>
        <p className="text-muted-foreground mt-1 text-sm">
          月次パラメータで受付番号レンジを決め、入荷数・1号機処理量（近似）を集計します。
        </p>
      </div>

      <div className="flex flex-wrap items-end gap-3">
        <div className="flex flex-col gap-1">
          <label className="text-muted-foreground text-xs">年</label>
          <input
            value={year}
            onChange={(e: ChangeEvent<HTMLInputElement>) => setYear(e.target.value)}
            className="w-24 rounded-xl border border-cream-300 bg-white px-3 py-2 font-mono text-sm text-ink focus:outline-none focus:ring-2 focus:ring-terracotta/40"
          />
        </div>
        <div className="flex flex-col gap-1">
          <label className="text-muted-foreground text-xs">月</label>
          <input
            value={month}
            onChange={(e: ChangeEvent<HTMLInputElement>) => setMonth(e.target.value)}
            className="w-20 rounded-xl border border-cream-300 bg-white px-3 py-2 font-mono text-sm text-ink focus:outline-none focus:ring-2 focus:ring-terracotta/40"
          />
        </div>
        <div className="flex flex-col gap-1">
          <label className="text-muted-foreground text-xs">受付番号 From（任意）</label>
          <input
            value={unoMin}
            onChange={(e: ChangeEvent<HTMLInputElement>) => setUnoMin(e.target.value)}
            className="w-40 rounded-xl border border-cream-300 bg-white px-3 py-2 font-mono text-sm text-ink focus:outline-none focus:ring-2 focus:ring-terracotta/40"
          />
        </div>
        <div className="flex flex-col gap-1">
          <label className="text-muted-foreground text-xs">受付番号 To（任意）</label>
          <input
            value={unoMax}
            onChange={(e: ChangeEvent<HTMLInputElement>) => setUnoMax(e.target.value)}
            className="w-40 rounded-xl border border-cream-300 bg-white px-3 py-2 font-mono text-sm text-ink focus:outline-none focus:ring-2 focus:ring-terracotta/40"
          />
        </div>
        <Button type="button" size="sm" onClick={() => void run()} disabled={loading}>
          集計実行
        </Button>
      </div>

      {error ? <p className="text-destructive text-sm">{error}</p> : null}

      {data ? (
        <div className="space-y-3">
          <div className="bg-muted rounded-md p-3 text-xs font-mono leading-relaxed">
            <div>
              UNO {data.params.unoMin} … {data.params.unoMax}
            </div>
            <div>
              完了月日レンジ（ctimer 近似） {data.params.kkmd1} … {data.params.kkmd2}
            </div>
          </div>

          <table className="w-full max-w-xl border text-sm">
            <thead className="bg-muted/80">
              <tr>
                <th className="border px-2 py-1 text-left">号機</th>
                <th className="border px-2 py-1 text-right">入荷数計</th>
                <th className="border px-2 py-1 text-right">処理量（近似）</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td className="border px-2 py-1">1号機</td>
                <td className="border px-2 py-1 text-right">{data.goukei.ric1.inboundQty}</td>
                <td className="border px-2 py-1 text-right">{data.goukei.ric1.simpleProcessingApprox}</td>
              </tr>
              <tr>
                <td className="border px-2 py-1">2号機</td>
                <td className="border px-2 py-1 text-right">{data.goukei.ric2.inboundQty}</td>
                <td className="border px-2 py-1 text-right">{data.goukei.ric2.simpleProcessingApprox}</td>
              </tr>
              <tr>
                <td className="border px-2 py-1">3号機</td>
                <td className="border px-2 py-1 text-right">{data.goukei.ric3.inboundQty}</td>
                <td className="border px-2 py-1 text-right">{data.goukei.ric3.simpleProcessingApprox}</td>
              </tr>
            </tbody>
          </table>

          <ul className="text-muted-foreground list-disc space-y-1 pl-5 text-xs">
            {data.notes.map((n) => (
              <li key={n}>{n}</li>
            ))}
          </ul>

          <pre className="bg-muted max-h-[280px] overflow-auto rounded-md p-3 text-xs">{JSON.stringify(data, null, 2)}</pre>
        </div>
      ) : null}
    </div>
  );
}
