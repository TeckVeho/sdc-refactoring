"use client";

import { pdf } from "@react-pdf/renderer";
import { useCallback, useEffect, useMemo, useRef, useState } from "react";

import { PrintButton } from "@/components/shared/print-button";
import { Button } from "@/components/ui/button";
import { apiFetch } from "@/lib/api";
import { cn } from "@/lib/utils";

import {
  Ric2DoseShortagePdfDoc,
  type Ric2PdfPlanRow,
} from "./ric2-dose-shortage-pdf";

type DataPayload = {
  data: {
    summary: {
      syouk2Count: number;
      zaikoCount: number;
      kansokuCount: number;
      syouj2Count: number;
    };
    recentPlansWithZaiko: Array<{
      uno: string;
      syono: string | null;
      kainame: string | null;
      syokind: string | null;
      sdate: string | null;
      edate: string | null;
      bikou: string | null;
      zaiko: {
        kaisyacd: string;
        sehncd: string;
        kainame: string | null;
      } | null;
    }>;
    kansokuRows: Array<{
      kid: string;
      sokutei: string | null;
      bikou: string | null;
    }>;
    schemaNote: string;
  };
};

const emptyRow = (): Ric2PdfPlanRow => ({
  ukeNo: "",
  ricNoS: "",
  ricNoE: "",
  siteiSn: "",
  kagen: "",
  joug: "",
  jituP: "",
  kaiName: "",
});

export default function Ric2DoseShortageReportPage() {
  const iframeRef = useRef<HTMLIFrameElement>(null);

  const [loadError, setLoadError] = useState<string | null>(null);
  const [serverData, setServerData] = useState<DataPayload["data"] | null>(null);
  const [previewUrl, setPreviewUrl] = useState<string | null>(null);
  const [pdfBusy, setPdfBusy] = useState(false);

  const [yoshikiNo, setYoshikiNo] = useState('様式 G2-12(MVP)');
  const [senkNo, setSenkNo] = useState("");
  const [sokDate, setSokDate] = useState("");
  const [soktCd, setSoktCd] = useState("");
  const [saiSoku, setSaiSoku] = useState("");
  const [sokutSn, setSokutSn] = useState("");
  const [tani, setTani] = useState("");
  const [pass, setPass] = useState("");
  const [sikiKodo, setSikiKodo] = useState("");
  const [sosi, setSosi] = useState("");
  const [atusa, setAtusa] = useState("");
  const [senKind, setSenKind] = useState("");
  const [atusaT, setAtusaT] = useState("");
  const [keisask, setKeisask] = useState("");
  const [planRows, setPlanRows] = useState<Ric2PdfPlanRow[]>(() =>
    Array.from({ length: 8 }, emptyRow),
  );

  useEffect(() => {
    setLoadError(null);
    void apiFetch<DataPayload>("/api/reports/ric2-dose-shortage/data")
      .then((r) => setServerData(r.data))
      .catch((e: Error) => {
        setServerData(null);
        setLoadError(e.message);
      });
  }, []);

  useEffect(() => {
    return () => {
      if (previewUrl) URL.revokeObjectURL(previewUrl);
    };
  }, [previewUrl]);

  const pdfProps = useMemo(
    () => ({
      yoshikiNo,
      senkNo,
      sokDate,
      soktCd,
      saiSoku,
      sokutSn,
      tani,
      pass,
      sikiKodo,
      sosi,
      atusa,
      senKind,
      atusaT,
      keisask,
      planRows,
    }),
    [
      yoshikiNo,
      senkNo,
      sokDate,
      soktCd,
      saiSoku,
      sokutSn,
      tani,
      pass,
      sikiKodo,
      sosi,
      atusa,
      senKind,
      atusaT,
      keisask,
      planRows,
    ],
  );

  const buildPdfBlob = useCallback(async () => {
    return pdf(<Ric2DoseShortagePdfDoc {...pdfProps} />).toBlob();
  }, [pdfProps]);

  const onGeneratePreview = useCallback(async () => {
    setPdfBusy(true);
    try {
      const blob = await buildPdfBlob();
      setPreviewUrl((prev) => {
        if (prev) URL.revokeObjectURL(prev);
        return URL.createObjectURL(blob);
      });
    } finally {
      setPdfBusy(false);
    }
  }, [buildPdfBlob]);

  const onDownload = useCallback(async () => {
    setPdfBusy(true);
    try {
      const blob = await buildPdfBlob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `RIC2線量不足報告_${senkNo || "draft"}.pdf`;
      a.click();
      URL.revokeObjectURL(url);
    } finally {
      setPdfBusy(false);
    }
  }, [buildPdfBlob, senkNo]);

  const applyFromApi = useCallback(() => {
    if (!serverData?.recentPlansWithZaiko.length) return;
    const next = Array.from({ length: 8 }, emptyRow);
    serverData.recentPlansWithZaiko.slice(0, 8).forEach((p, i) => {
      next[i] = {
        ukeNo: p.uno,
        ricNoS: p.syono ?? "",
        ricNoE: p.edate ?? "",
        siteiSn: "",
        kagen: "",
        joug: "",
        jituP: "",
        kaiName: p.kainame ?? p.zaiko?.kainame ?? "",
      };
    });
    setPlanRows(next);
  }, [serverData]);

  const updatePlanRow = useCallback(
    (index: number, patch: Partial<Ric2PdfPlanRow>) => {
      setPlanRows((rows) =>
        rows.map((r, i) => (i === index ? { ...r, ...patch } : r)),
      );
    },
    [],
  );

  return (
    <div className="mx-auto max-w-6xl space-y-6 p-6 text-ink">
      <header className="space-y-1">
        <h1 className="text-xl font-semibold tracking-tight">
          RIC2 線量不足報告書
        </h1>
        <p className="text-sm text-ink-muted">
          ２号機線量不足（工程確認・追加照射指示書）Web MVP — データは API
          サマリー参照、PDF はブラウザで生成します。
        </p>
      </header>

      {loadError ? (
        <p className="rounded-lg border border-destructive/40 bg-destructive/5 px-3 py-2 text-sm text-destructive">
          データ取得: {loadError}
        </p>
      ) : null}

      {serverData ? (
        <section className="rounded-xl border border-border/80 bg-cream/40 p-4 text-sm">
          <h2 className="mb-2 font-medium">DB サマリー（参照）</h2>
          <ul className="grid gap-2 sm:grid-cols-2 md:grid-cols-4">
            <li>Syouk2: {serverData.summary.syouk2Count.toLocaleString()} 件</li>
            <li>Zaiko: {serverData.summary.zaikoCount.toLocaleString()} 件</li>
            <li>Kansoku: {serverData.summary.kansokuCount.toLocaleString()} 件</li>
            <li>Syouj2: {serverData.summary.syouj2Count.toLocaleString()} 件</li>
          </ul>
          <p className="mt-2 text-xs text-ink-muted">{serverData.schemaNote}</p>
          <div className="mt-3 flex flex-wrap gap-2">
            <Button type="button" variant="secondary" size="sm" onClick={applyFromApi}>
              直近計画から受付列を反映
            </Button>
          </div>
        </section>
      ) : null}

      <div className="grid gap-6 lg:grid-cols-2">
        <section className="space-y-4 rounded-xl border border-border/80 bg-card p-4">
          <h2 className="font-medium">帳票ヘッダ</h2>
          <div className="grid gap-3 sm:grid-cols-2">
            <label className="text-xs">
              <span className="text-ink-muted">様式番号表示</span>
              <input
                className="mt-1 w-full rounded-md border border-input bg-background px-2 py-1.5 text-sm"
                value={yoshikiNo}
                onChange={(e) => setYoshikiNo(e.target.value)}
              />
            </label>
            <label className="text-xs">
              <span className="text-ink-muted">線量計番号</span>
              <input
                className="mt-1 w-full rounded-md border border-input bg-background px-2 py-1.5 text-sm"
                value={senkNo}
                onChange={(e) => setSenkNo(e.target.value)}
              />
            </label>
            <label className="text-xs">
              <span className="text-ink-muted">測定日</span>
              <input
                className="mt-1 w-full rounded-md border border-input bg-background px-2 py-1.5 text-sm"
                value={sokDate}
                onChange={(e) => setSokDate(e.target.value)}
              />
            </label>
            <label className="text-xs">
              <span className="text-ink-muted">測定者コード</span>
              <input
                className="mt-1 w-full rounded-md border border-input bg-background px-2 py-1.5 text-sm"
                value={soktCd}
                onChange={(e) => setSoktCd(e.target.value)}
              />
            </label>
            <label className="text-xs">
              <span className="text-ink-muted">再測定者</span>
              <input
                className="mt-1 w-full rounded-md border border-input bg-background px-2 py-1.5 text-sm"
                value={saiSoku}
                onChange={(e) => setSaiSoku(e.target.value)}
              />
            </label>
            <label className="text-xs">
              <span className="text-ink-muted">測定線量</span>
              <input
                className="mt-1 w-full rounded-md border border-input bg-background px-2 py-1.5 text-sm"
                value={sokutSn}
                onChange={(e) => setSokutSn(e.target.value)}
              />
            </label>
            <label className="text-xs">
              <span className="text-ink-muted">照射目的</span>
              <input
                className="mt-1 w-full rounded-md border border-input bg-background px-2 py-1.5 text-sm"
                value={tani}
                onChange={(e) => setTani(e.target.value)}
              />
            </label>
            <label className="text-xs">
              <span className="text-ink-muted">指定パス数</span>
              <input
                className="mt-1 w-full rounded-md border border-input bg-background px-2 py-1.5 text-sm"
                value={pass}
                onChange={(e) => setPass(e.target.value)}
              />
            </label>
            <label className="text-xs">
              <span className="text-ink-muted">計算式コード</span>
              <input
                className="mt-1 w-full rounded-md border border-input bg-background px-2 py-1.5 text-sm"
                value={sikiKodo}
                onChange={(e) => setSikiKodo(e.target.value)}
              />
            </label>
            <label className="text-xs">
              <span className="text-ink-muted">素子異常</span>
              <input
                className="mt-1 w-full rounded-md border border-input bg-background px-2 py-1.5 text-sm"
                value={sosi}
                onChange={(e) => setSosi(e.target.value)}
              />
            </label>
            <label className="text-xs">
              <span className="text-ink-muted">素子厚さ</span>
              <input
                className="mt-1 w-full rounded-md border border-input bg-background px-2 py-1.5 text-sm"
                value={atusa}
                onChange={(e) => setAtusa(e.target.value)}
              />
            </label>
            <label className="text-xs">
              <span className="text-ink-muted">線量計種類</span>
              <input
                className="mt-1 w-full rounded-md border border-input bg-background px-2 py-1.5 text-sm"
                value={senKind}
                onChange={(e) => setSenKind(e.target.value)}
              />
            </label>
            <label className="text-xs">
              <span className="text-ink-muted">線量計厚さ</span>
              <input
                className="mt-1 w-full rounded-md border border-input bg-background px-2 py-1.5 text-sm"
                value={atusaT}
                onChange={(e) => setAtusaT(e.target.value)}
              />
            </label>
            <label className="text-xs sm:col-span-2">
              <span className="text-ink-muted">計算式（表示用）</span>
              <input
                className="mt-1 w-full rounded-md border border-input bg-background px-2 py-1.5 text-sm"
                value={keisask}
                onChange={(e) => setKeisask(e.target.value)}
              />
            </label>
          </div>

          <h3 className="pt-2 text-sm font-medium">照射計画（8行）</h3>
          <div className="max-h-[280px] overflow-auto rounded-md border border-border/60">
            <table className="w-full min-w-[720px] text-left text-xs">
              <thead className="sticky top-0 bg-muted/80">
                <tr>
                  <th className="px-1 py-1">#</th>
                  <th className="px-1 py-1">受付</th>
                  <th className="px-1 py-1">RIC開始</th>
                  <th className="px-1 py-1">RIC終了</th>
                  <th className="px-1 py-1">指定線量</th>
                  <th className="px-1 py-1">下限</th>
                  <th className="px-1 py-1">上限</th>
                  <th className="px-1 py-1">実パス</th>
                  <th className="px-1 py-1">会社名</th>
                </tr>
              </thead>
              <tbody>
                {planRows.map((row, i) => (
                  <tr key={i} className="border-t border-border/50">
                    <td className="px-1 py-0.5 text-ink-muted">{i + 1}</td>
                    {(
                      [
                        "ukeNo",
                        "ricNoS",
                        "ricNoE",
                        "siteiSn",
                        "kagen",
                        "joug",
                        "jituP",
                        "kaiName",
                      ] as const
                    ).map((key) => (
                      <td key={key} className="p-0">
                        <input
                          className={cn(
                            "w-full min-w-[56px] border-0 bg-transparent px-1 py-0.5",
                            "focus:bg-background focus:ring-1 focus:ring-ring",
                          )}
                          value={row[key]}
                          onChange={(e) =>
                            updatePlanRow(i, { [key]: e.target.value })
                          }
                        />
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          {serverData && serverData.kansokuRows.length > 0 ? (
            <details className="text-xs text-ink-muted">
              <summary className="cursor-pointer">観測マスタ（kid）参照</summary>
              <ul className="mt-2 max-h-24 overflow-auto font-mono">
                {serverData.kansokuRows.slice(0, 30).map((k) => (
                  <li key={k.kid}>
                    {k.kid}: {k.sokutei ?? ""}
                  </li>
                ))}
              </ul>
            </details>
          ) : null}
        </section>

        <section className="flex flex-col gap-3 rounded-xl border border-border/80 bg-card p-4">
          <h2 className="font-medium">PDF</h2>
          <div className="flex flex-wrap gap-2">
            <Button
              type="button"
              onClick={() => void onGeneratePreview()}
              disabled={pdfBusy}
            >
              {pdfBusy ? "生成中…" : "プレビュー生成"}
            </Button>
            <Button
              type="button"
              variant="secondary"
              onClick={() => void onDownload()}
              disabled={pdfBusy}
            >
              ダウンロード
            </Button>
            <PrintButton
              onPrint={() => iframeRef.current?.contentWindow?.print()}
            />
          </div>
          <div className="min-h-[480px] flex-1 overflow-hidden rounded-lg border border-dashed border-border bg-muted/20">
            {previewUrl ? (
              <iframe
                ref={iframeRef}
                title="RIC2 PDF preview"
                className="h-[70vh] w-full"
                src={previewUrl}
              />
            ) : (
              <p className="p-6 text-sm text-ink-muted">
                「プレビュー生成」で blob URL を iframe に表示します。
              </p>
            )}
          </div>
        </section>
      </div>
    </div>
  );
}
