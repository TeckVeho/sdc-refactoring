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

type JmmPayload = {
  variant: "60" | "90";
  maxRows: number;
  title: string;
  header: {
    uno: string;
    honnsuu: number | null;
    dose: number | null;
    kainame: string | null;
    kaisyacd: string | null;
    sehncd: string | null;
    syono: string | null;
    syokind: string | null;
    bikou: string | null;
  };
  senkinds: { kindcd: string; name: string | null }[];
  keicodes: { keicode: string; keisiki: string | null }[];
};

const pdfStyles = StyleSheet.create({
  page: { padding: 32, fontSize: 9, fontFamily: "NotoSansJP" },
  title: { fontSize: 13, marginBottom: 8, textAlign: "center" },
  row: { flexDirection: "row", marginBottom: 4 },
  label: { width: "32%", color: "#444" },
  val: { width: "68%" },
});

function fmt(v: unknown): string {
  if (v == null) return "—";
  const s = String(v).trim();
  return s === "" ? "—" : s;
}

function JmmPdf({ data }: { data: JmmPayload }) {
  const h = data.header;
  return (
    <Document>
      <Page size="A4" style={pdfStyles.page}>
        <Text style={pdfStyles.title}>{data.title}</Text>
        <View style={pdfStyles.row}>
          <Text style={pdfStyles.label}>受付番号</Text>
          <Text style={pdfStyles.val}>{fmt(h.uno)}</Text>
        </View>
        <View style={pdfStyles.row}>
          <Text style={pdfStyles.label}>顧客</Text>
          <Text style={pdfStyles.val}>{fmt(h.kainame)}</Text>
        </View>
        <View style={pdfStyles.row}>
          <Text style={pdfStyles.label}>会社／製品コード</Text>
          <Text style={pdfStyles.val}>
            {fmt(h.kaisyacd)} / {fmt(h.sehncd)}
          </Text>
        </View>
        <View style={pdfStyles.row}>
          <Text style={pdfStyles.label}>本数 / 指定線量</Text>
          <Text style={pdfStyles.val}>
            {fmt(h.honnsuu)} / {fmt(h.dose)}
          </Text>
        </View>
        <View style={pdfStyles.row}>
          <Text style={pdfStyles.label}>線種 / RIC</Text>
          <Text style={pdfStyles.val}>
            {fmt(h.syokind)} / {fmt(h.syono)}
          </Text>
        </View>
        <View style={pdfStyles.row}>
          <Text style={pdfStyles.label}>備考</Text>
          <Text style={pdfStyles.val}>{fmt(h.bikou)}</Text>
        </View>
        <Text style={{ marginTop: 12, fontSize: 8, color: "#666" }}>
          測定グリッド・線量多項式計算は Web MVP ではローカル入力予定（仕様: ExJMM60/90φ記入用紙）。
        </Text>
      </Page>
    </Document>
  );
}

export function JmmEntryClient({
  variant,
  apiPath,
  caption,
}: {
  variant: "60" | "90";
  apiPath: "/api/dosimetry/jmm60" | "/api/dosimetry/jmm90";
  caption: string;
}) {
  const [uno, setUno] = useState("");
  const [data, setData] = useState<JmmPayload | null>(null);
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
      const res = await apiFetch<{ data: JmmPayload }>(
        `${apiPath}?${new URLSearchParams({ uno: u })}`,
      );
      setData(res.data);
    } catch (e) {
      setData(null);
      setError(e instanceof Error ? e.message : "取得に失敗しました");
    } finally {
      setLoading(false);
    }
  }, [apiPath, uno]);

  const downloadPdf = useCallback(async () => {
    if (!data) return;
    const blob = await pdf(<JmmPdf data={data} />).toBlob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `${data.header.uno}_JMM${variant}.pdf`;
    a.click();
    URL.revokeObjectURL(url);
  }, [data, variant]);

  return (
    <div className="space-y-4 p-4">
      <div>
        <h1 className="text-lg font-semibold tracking-tight">{caption}</h1>
        <p className="text-muted-foreground mt-1 text-sm">
          受付番号で zaiko / syouk1 を読み込み、記入用紙のヘッダーを表示します（最大行数:{" "}
          {variant === "60" ? 540 : 300}）。
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
        <Button type="button" size="sm" variant="outline" onClick={() => void downloadPdf()} disabled={!data}>
          PDF
        </Button>
      </div>

      {error ? <p className="text-destructive text-sm">{error}</p> : null}

      {data ? (
        <pre className="bg-muted max-h-[420px] overflow-auto rounded-md p-3 text-xs leading-relaxed">
          {JSON.stringify(data, null, 2)}
        </pre>
      ) : null}
    </div>
  );
}
