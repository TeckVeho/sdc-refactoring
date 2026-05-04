"use client";

import {
  Document,
  Font,
  Page,
  pdf,
  StyleSheet,
  Text,
  View,
} from "@react-pdf/renderer";
import { useCallback, useEffect, useState } from "react";

import { Button } from "@/components/ui/button";
import { apiFetch } from "@/lib/api";

Font.register({
  family: "NotoSansJP",
  src: "https://cdn.jsdelivr.net/gh/googlefonts/noto-cjk@main/Sans/OTF/Japanese/NotoSansCJKjp-Regular.otf",
});

type ListPayload = {
  data: {
    kind: "list";
    shortageCount: number;
    candidates: Array<{ uno: string; syono: string | null; kainame: string | null }>;
    truncated: boolean;
  };
};

type ReportPayload = {
  data: {
    kind: "report";
    meta: { formNo: string; title: string; variant: "general" };
    uno: string;
    senkNo: string;
    doseMeterLast4: string;
    measurement: Record<string, unknown>;
    plans: Array<{
      slot: number;
      ukeNo: string | null;
      ricNoS: string | null;
      ricNoE: string | null;
      siteiSn: number | null;
      kagen: number | null;
      joug: number | null;
      jituP: number | null;
      kaiCd: string | null;
      kaiName: string | null;
    }>;
    additional: Array<{
      slot: number;
      ukeNo: string | null;
      ricNoS: string | null;
      ricNoE: string | null;
      shortage: number | null;
      extraPasses: number | null;
      kaiName: string | null;
    }>;
    doseMeter: { senKind: string | null; atusaT: string | null; keisask: string | null };
    warnings: {
      tooManyShortage: boolean;
      negativeExtraPasses: boolean;
      extraPassesMismatch: boolean;
    };
    employees: Array<{ shano: string; shaname: string }>;
  };
};

const pdfStyles = StyleSheet.create({
  page: {
    padding: 28,
    fontSize: 9,
    fontFamily: "NotoSansJP",
  },
  title: { fontSize: 13, marginBottom: 6, textAlign: "center" },
  sub: { fontSize: 8, marginBottom: 12, textAlign: "center", color: "#444" },
  section: { marginTop: 10, marginBottom: 4, fontSize: 10 },
  row: { flexDirection: "row", borderBottomWidth: 0.5, borderBottomColor: "#ccc", paddingVertical: 3 },
  cellLabel: { width: "28%", color: "#333" },
  cellVal: { width: "72%" },
  tableHead: {
    flexDirection: "row",
    borderBottomWidth: 1,
    borderBottomColor: "#333",
    paddingVertical: 4,
    fontSize: 8,
  },
  th: { flex: 1, paddingRight: 2 },
  tr: { flexDirection: "row", paddingVertical: 3, fontSize: 7 },
  td: { flex: 1, paddingRight: 2 },
  warn: { marginTop: 8, fontSize: 8, color: "#a32020" },
  mono: { fontFamily: "Helvetica", fontSize: 8 },
});

function fmt(v: unknown): string {
  if (v == null) return "—";
  if (typeof v === "number" && Number.isFinite(v)) return String(v);
  const s = String(v).trim();
  return s === "" ? "—" : s;
}

function Ric3PdfDoc({ data }: { data: ReportPayload["data"] }) {
  const m = data.measurement as {
    sokDate?: string | null;
    sokutSn?: number | null;
    pass?: number | null;
    tani?: string | null;
    sikiKodo?: string | null;
    syokind?: string | null;
    bikou?: string | null;
  };

  return (
    <Document>
      <Page size="A4" style={pdfStyles.page}>
        <Text style={pdfStyles.title}>{data.meta.title}</Text>
        <Text style={pdfStyles.sub}>
          様式 {data.meta.formNo} / UNO:{" "}
          <Text style={pdfStyles.mono}>{data.uno}</Text>
        </Text>

        <Text style={pdfStyles.section}>測定・概要</Text>
        <View style={pdfStyles.row}>
          <Text style={pdfStyles.cellLabel}>照射管理番号（SenkNo）</Text>
          <Text style={[pdfStyles.cellVal, pdfStyles.mono]}>{fmt(data.senkNo)}</Text>
        </View>
        <View style={pdfStyles.row}>
          <Text style={pdfStyles.cellLabel}>線量計番号（下4桁）</Text>
          <Text style={[pdfStyles.cellVal, pdfStyles.mono]}>{fmt(data.doseMeterLast4)}</Text>
        </View>
        <View style={pdfStyles.row}>
          <Text style={pdfStyles.cellLabel}>測定日</Text>
          <Text style={pdfStyles.cellVal}>{fmt(m.sokDate)}</Text>
        </View>
        <View style={pdfStyles.row}>
          <Text style={pdfStyles.cellLabel}>測定線量（SokutSN）</Text>
          <Text style={[pdfStyles.cellVal, pdfStyles.mono]}>{fmt(m.sokutSn)}</Text>
        </View>
        <View style={pdfStyles.row}>
          <Text style={pdfStyles.cellLabel}>指定パス（Pass）</Text>
          <Text style={[pdfStyles.cellVal, pdfStyles.mono]}>{fmt(m.pass)}</Text>
        </View>
        <View style={pdfStyles.row}>
          <Text style={pdfStyles.cellLabel}>製品目的（Tani）</Text>
          <Text style={pdfStyles.cellVal}>{fmt(m.tani)}</Text>
        </View>
        <View style={pdfStyles.row}>
          <Text style={pdfStyles.cellLabel}>計算式コード</Text>
          <Text style={[pdfStyles.cellVal, pdfStyles.mono]}>{fmt(m.sikiKodo)}</Text>
        </View>

        <Text style={pdfStyles.section}>照射計画（最大5件・MVPは1件中心）</Text>
        <View style={pdfStyles.tableHead}>
          <Text style={pdfStyles.th}>受付</Text>
          <Text style={pdfStyles.th}>RIC-S</Text>
          <Text style={pdfStyles.th}>RIC-E</Text>
          <Text style={pdfStyles.th}>指定SN</Text>
          <Text style={pdfStyles.th}>実P</Text>
          <Text style={pdfStyles.th}>会社</Text>
        </View>
        {data.plans.map((p) => (
          <View key={p.slot} style={pdfStyles.tr} wrap={false}>
            <Text style={[pdfStyles.td, pdfStyles.mono]}>{fmt(p.ukeNo)}</Text>
            <Text style={[pdfStyles.td, pdfStyles.mono]}>{fmt(p.ricNoS)}</Text>
            <Text style={[pdfStyles.td, pdfStyles.mono]}>{fmt(p.ricNoE)}</Text>
            <Text style={[pdfStyles.td, pdfStyles.mono]}>{fmt(p.siteiSn)}</Text>
            <Text style={[pdfStyles.td, pdfStyles.mono]}>{fmt(p.jituP)}</Text>
            <Text style={pdfStyles.td}>{fmt(p.kaiName)}</Text>
          </View>
        ))}

        <Text style={pdfStyles.section}>追加照射（不足量・ROUNDUP）</Text>
        <View style={pdfStyles.tableHead}>
          <Text style={pdfStyles.th}>受付</Text>
          <Text style={pdfStyles.th}>不足</Text>
          <Text style={pdfStyles.th}>追加ﾊﾟｽ</Text>
          <Text style={pdfStyles.th}>会社</Text>
        </View>
        {data.additional.map((r) => (
          <View key={r.slot} style={pdfStyles.tr} wrap={false}>
            <Text style={[pdfStyles.td, pdfStyles.mono]}>{fmt(r.ukeNo)}</Text>
            <Text style={[pdfStyles.td, pdfStyles.mono]}>{fmt(r.shortage)}</Text>
            <Text style={[pdfStyles.td, pdfStyles.mono]}>{fmt(r.extraPasses)}</Text>
            <Text style={pdfStyles.td}>{fmt(r.kaiName)}</Text>
          </View>
        ))}

        <View style={pdfStyles.row}>
          <Text style={pdfStyles.cellLabel}>線量計種類</Text>
          <Text style={pdfStyles.cellVal}>{fmt(data.doseMeter.senKind)}</Text>
        </View>

        {(data.warnings.negativeExtraPasses || data.warnings.extraPassesMismatch) && (
          <Text style={pdfStyles.warn}>
            {data.warnings.negativeExtraPasses ? "追加パス計算が負値です。測定データを確認してください。" : ""}
            {data.warnings.extraPassesMismatch ? " 追加パス数が行ごとに異なります。" : ""}
          </Text>
        )}

        <Text style={[pdfStyles.sub, { marginTop: 14 }]}>
          Syoukj3 / Syouk1 / Zaiko MVP 結合 — 下限・上限が DB に無い場合は追加パスは空です。
        </Text>
      </Page>
    </Document>
  );
}

export default function Ric3DoseShortageReportPage() {
  const [candidates, setCandidates] = useState<ListPayload["data"]["candidates"]>([]);
  const [truncated, setTruncated] = useState(false);
  const [shortageCount, setShortageCount] = useState<number | null>(null);
  const [uno, setUno] = useState("");
  const [report, setReport] = useState<ReportPayload["data"] | null>(null);
  const [loading, setLoading] = useState(false);
  const [listLoading, setListLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    let cancelled = false;
    (async () => {
      setListLoading(true);
      setError(null);
      try {
        const res = await apiFetch<ListPayload>("/api/reports/ric3-dose-shortage");
        if (cancelled) return;
        setCandidates(res.data.candidates);
        setTruncated(res.data.truncated);
        setShortageCount(res.data.shortageCount);
      } catch (e) {
        if (!cancelled) setError(e instanceof Error ? e.message : "一覧の取得に失敗しました");
      } finally {
        if (!cancelled) setListLoading(false);
      }
    })();
    return () => {
      cancelled = true;
    };
  }, []);

  const loadReport = useCallback(async () => {
    const u = uno.trim();
    if (!u) {
      setError("受付番号（uno）を選択または入力してください");
      return;
    }
    setLoading(true);
    setError(null);
    try {
      const res = await apiFetch<ReportPayload>(
        `/api/reports/ric3-dose-shortage?${new URLSearchParams({ uno: u })}`,
      );
      setReport(res.data);
    } catch (e) {
      setReport(null);
      setError(e instanceof Error ? e.message : "報告データの取得に失敗しました");
    } finally {
      setLoading(false);
    }
  }, [uno]);

  const downloadPdf = useCallback(async () => {
    if (!report) return;
    const blob = await pdf(<Ric3PdfDoc data={report} />).toBlob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `ric3-dose-shortage-${report.uno}.pdf`;
    a.click();
    URL.revokeObjectURL(url);
  }, [report]);

  return (
    <div className="mx-auto max-w-4xl space-y-6 p-4 md:p-6">
      <div>
        <h1 className="text-2xl font-semibold text-ink">RIC3線量不足報告書</h1>
        <p className="mt-1 text-sm text-ink-muted">
          ３号機線量不足報告書（工程確認・追加照射指示書）一般向け。MVP では{" "}
          <code className="rounded bg-cream-200 px-1">syoukj3</code> を主キーに{" "}
          <code className="rounded bg-cream-200 px-1">syouk1</code> /{" "}
          <code className="rounded bg-cream-200 px-1">zaiko</code> を結合します。
        </p>
      </div>

      <div className="rounded-2xl border border-cream-300 bg-cream-50 p-4 space-y-3">
        <p className="text-sm text-ink-muted">
          {shortageCount != null ? (
            <>
              <span className="font-medium text-ink">syoukj3 件数:</span> {shortageCount}{" "}
              {truncated ? "（一覧は上限で切り詰め）" : ""}
            </>
          ) : null}
        </p>
        <label className="flex flex-col gap-1 text-sm text-ink-muted">
          照射計画（uno）
          {listLoading ? (
            <span className="text-ink">読み込み中…</span>
          ) : (
            <select
              value={uno}
              onChange={(e) => setUno(e.target.value)}
              className="rounded-xl border border-cream-300 bg-white px-3 py-2 text-ink"
            >
              <option value="">（選択）</option>
              {candidates.map((c) => (
                <option key={c.uno} value={c.uno}>
                  {c.uno}
                  {c.syono ? ` / syono:${c.syono}` : ""}
                  {c.kainame ? ` — ${c.kainame}` : ""}
                </option>
              ))}
            </select>
          )}
        </label>
        <label className="flex flex-col gap-1 text-sm text-ink-muted">
          または uno を直接入力
          <input
            value={uno}
            onChange={(e) => setUno(e.target.value)}
            placeholder="10桁の受付番号など"
            className="rounded-xl border border-cream-300 bg-white px-3 py-2 text-ink focus:outline-none focus:ring-2 focus:ring-terracotta/40"
          />
        </label>
        <div className="flex flex-wrap gap-2">
          <Button type="button" onClick={() => void loadReport()} disabled={loading}>
            {loading ? "取得中…" : "報告データを取得"}
          </Button>
          <Button type="button" variant="outline" disabled={!report} onClick={() => void downloadPdf()}>
            PDF ダウンロード
          </Button>
        </div>
      </div>

      {error ? <p className="text-sm text-red-700">{error}</p> : null}

      {report ? (
        <div className="rounded-2xl border border-cream-300 bg-white p-4 text-sm space-y-2">
          <p className="font-semibold text-ink">{report.meta.title}</p>
          <p className="text-ink-muted">
            senkNo: <span className="font-mono text-ink">{report.senkNo}</span> / 線量計下4桁:{" "}
            <span className="font-mono text-ink">{report.doseMeterLast4}</span>
          </p>
          <pre className="max-h-64 overflow-auto rounded-lg bg-cream-100 p-3 text-xs">
            {JSON.stringify(report, null, 2)}
          </pre>
        </div>
      ) : null}
    </div>
  );
}
