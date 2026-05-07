"use client";

import {
  Document,
  Font,
  Page,
  StyleSheet,
  Text,
  View,
  pdf,
} from "@react-pdf/renderer";
import { type ChangeEvent, useCallback, useState } from "react";

import { Button } from "@/components/ui/button";
import { apiFetch } from "@/lib/api";

Font.register({
  family: "NotoSansJP",
  src: "https://cdn.jsdelivr.net/gh/googlefonts/noto-cjk@main/Sans/OTF/Japanese/NotoSansCJKjp-Regular.otf",
});

type WorkOrderPayload = {
  data: {
    meta: { formTitle: string; specRef: string };
    uno: string;
    zaiko: Record<string, unknown> | null;
    syouk1: Record<string, unknown> | null;
    derived: Record<string, unknown>;
    employees: { shano: string; shaname: string }[];
    ratetbl: { ratekey: string; rateval: string | null }[];
  };
};

const pdfStyles = StyleSheet.create({
  page: { padding: 28, fontSize: 9, fontFamily: "NotoSansJP" },
  title: { fontSize: 14, marginBottom: 6, textAlign: "center" },
  sub: { fontSize: 8, marginBottom: 12, textAlign: "center", color: "#555" },
  row: { flexDirection: "row", borderBottomWidth: 0.5, borderBottomColor: "#ddd", paddingVertical: 3 },
  cellLabel: { width: "34%", color: "#333" },
  cellVal: { width: "66%" },
});

function fmt(v: unknown): string {
  if (v == null) return "—";
  const s = String(v).trim();
  return s === "" ? "—" : s;
}

function WorkOrderPdf({ payload }: { payload: WorkOrderPayload["data"] }) {
  const d = payload.derived;
  return (
    <Document>
      <Page size="A4" style={pdfStyles.page}>
        <Text style={pdfStyles.title}>{payload.meta.formTitle}</Text>
        <Text style={pdfStyles.sub}>UNO: {payload.uno}</Text>

        <View style={pdfStyles.row}>
          <Text style={pdfStyles.cellLabel}>顧客名</Text>
          <Text style={pdfStyles.cellVal}>{fmt(d.kainame)}</Text>
        </View>
        <View style={pdfStyles.row}>
          <Text style={pdfStyles.cellLabel}>入荷日</Text>
          <Text style={pdfStyles.cellVal}>{fmt(d.nyukaBi)}</Text>
        </View>
        <View style={pdfStyles.row}>
          <Text style={pdfStyles.cellLabel}>指定線量 / 本数</Text>
          <Text style={pdfStyles.cellVal}>
            {fmt(d.siteiSn)} / {fmt(d.syosuu)}
          </Text>
        </View>
        <View style={pdfStyles.row}>
          <Text style={pdfStyles.cellLabel}>測定線量（senritu）</Text>
          <Text style={pdfStyles.cellVal}>{fmt(d.senritu)}</Text>
        </View>
        <View style={pdfStyles.row}>
          <Text style={pdfStyles.cellLabel}>RIC / 線種</Text>
          <Text style={pdfStyles.cellVal}>
            {fmt(d.syono)} / {fmt(d.syokind)}
          </Text>
        </View>
        <View style={pdfStyles.row}>
          <Text style={pdfStyles.cellLabel}>備考</Text>
          <Text style={pdfStyles.cellVal}>{fmt(d.bikou)}</Text>
        </View>
        <Text style={{ marginTop: 14, fontSize: 8, color: "#666" }}>
          Web MVP: 照射時間・セット方法など Excel 全項目は順次追加（{payload.meta.specRef}）。
        </Text>
      </Page>
    </Document>
  );
}

export default function Machine1WorkOrderPage() {
  const [uno, setUno] = useState("");
  const [payload, setPayload] = useState<WorkOrderPayload["data"] | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [loading, setLoading] = useState(false);

  const load = useCallback(async () => {
    const u = uno.trim();
    if (!u) {
      setError("受付番号を入力してください");
      return;
    }
    setLoading(true);
    setError(null);
    try {
      const res = await apiFetch<WorkOrderPayload>(
        `/api/irradiation/machine1/work-order?${new URLSearchParams({ uno: u })}`,
      );
      setPayload(res.data);
    } catch (e) {
      setPayload(null);
      setError(e instanceof Error ? e.message : "取得に失敗しました");
    } finally {
      setLoading(false);
    }
  }, [uno]);

  const downloadPdf = useCallback(async () => {
    if (!payload) return;
    const blob = await pdf(<WorkOrderPdf payload={payload} />).toBlob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `${payload.uno}_work-order.pdf`;
    a.click();
    URL.revokeObjectURL(url);
  }, [payload]);

  return (
    <div className="space-y-4 p-4">
      <div>
        <h1 className="text-lg font-semibold tracking-tight">1号機作業指図書</h1>
        <p className="text-muted-foreground mt-1 text-sm">
          受付番号で在庫・照射データを取得し、印刷用サマリーを出力します。
        </p>
      </div>

      <div className="flex flex-wrap items-end gap-2">
        <div className="flex flex-col gap-1">
          <label className="text-muted-foreground text-xs">受付番号（uno）</label>
          <input
            value={uno}
            onChange={(e: ChangeEvent<HTMLInputElement>) => setUno(e.target.value)}
            className="w-56 rounded-xl border border-cream-300 bg-white px-3 py-2 font-mono text-sm text-ink focus:outline-none focus:ring-2 focus:ring-terracotta/40"
          />
        </div>
        <Button type="button" size="sm" onClick={() => void load()} disabled={loading}>
          読込
        </Button>
        <Button type="button" variant="outline" size="sm" onClick={() => void downloadPdf()} disabled={!payload}>
          PDF
        </Button>
      </div>

      {error ? <p className="text-destructive text-sm">{error}</p> : null}

      {payload ? (
        <pre className="bg-muted max-h-[480px] overflow-auto rounded-md p-3 text-xs leading-relaxed">
          {JSON.stringify(payload, null, 2)}
        </pre>
      ) : null}
    </div>
  );
}
