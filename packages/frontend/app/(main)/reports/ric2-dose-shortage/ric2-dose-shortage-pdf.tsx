"use client";

import {
  Document,
  Font,
  Page,
  StyleSheet,
  Text,
  View,
} from "@react-pdf/renderer";

/** RIC2: ROUNDUP((不足量 / (測定線量 / 実パス1)) / 2, 0) * 2 */
export function extraPassesRic2(
  shortage: number,
  measuredDose: number,
  jituP1: number,
): number {
  if (!(shortage > 0) || !(measuredDose > 0) || !(jituP1 > 0)) return 0;
  const unit = shortage / (measuredDose / jituP1);
  return Math.ceil(unit / 2) * 2;
}

Font.register({
  family: "NotoSansJP",
  src: "https://fonts.gstatic.com/ea/notosansjp/v5/NotoSansJP-Regular.otf",
});

const styles = StyleSheet.create({
  page: {
    padding: 36,
    fontSize: 9,
    fontFamily: "NotoSansJP",
  },
  title: {
    fontSize: 14,
    marginBottom: 4,
    textAlign: "center",
  },
  sub: {
    fontSize: 10,
    textAlign: "center",
    marginBottom: 12,
  },
  row: {
    flexDirection: "row",
    marginBottom: 3,
  },
  label: { width: 80 },
  val: { flex: 1 },
  tableHeader: {
    flexDirection: "row",
    borderBottomWidth: 1,
    borderBottomColor: "#000",
    marginTop: 8,
    paddingBottom: 2,
  },
  cell: { flex: 1, paddingRight: 4, fontSize: 7 },
  cellNarrow: { width: 36, paddingRight: 2, fontSize: 7 },
  tableRow: {
    flexDirection: "row",
    borderBottomWidth: 0.5,
    borderBottomColor: "#ccc",
    paddingTop: 2,
    paddingBottom: 2,
  },
  section: { marginTop: 10, fontSize: 10 },
});

export type Ric2PdfPlanRow = {
  ukeNo: string;
  ricNoS: string;
  ricNoE: string;
  siteiSn: string;
  kagen: string;
  joug: string;
  jituP: string;
  kaiName: string;
};

export type Ric2PdfProps = {
  yoshikiNo: string;
  senkNo: string;
  sokDate: string;
  soktCd: string;
  saiSoku: string;
  sokutSn: string;
  tani: string;
  pass: string;
  sikiKodo: string;
  sosi: string;
  atusa: string;
  senKind: string;
  atusaT: string;
  keisask: string;
  planRows: Ric2PdfPlanRow[];
};

function parseNum(s: string): number {
  const n = Number.parseFloat(String(s).replace(/,/g, "").trim());
  return Number.isFinite(n) ? n : 0;
}

export function Ric2DoseShortagePdfDoc(props: Ric2PdfProps) {
  const measured = parseNum(props.sokutSn);
  const jituP1 = parseNum(props.planRows[0]?.jituP ?? "0");

  return (
    <Document>
      <Page size="A4" style={styles.page}>
        <Text style={styles.sub}>{props.yoshikiNo}</Text>
        <Text style={styles.title}>
          ２号機線量不足報告書（工程確認・追加照射指示書）
        </Text>

        <View style={styles.row}>
          <Text style={styles.label}>線量計番号</Text>
          <Text style={styles.val}>{props.senkNo || "—"}</Text>
          <Text style={styles.label}>測定日</Text>
          <Text style={styles.val}>{props.sokDate || "—"}</Text>
        </View>
        <View style={styles.row}>
          <Text style={styles.label}>測定者</Text>
          <Text style={styles.val}>{props.soktCd || "—"}</Text>
          <Text style={styles.label}>再測定者</Text>
          <Text style={styles.val}>{props.saiSoku || "—"}</Text>
        </View>
        <View style={styles.row}>
          <Text style={styles.label}>測定線量</Text>
          <Text style={styles.val}>{props.sokutSn || "—"}</Text>
          <Text style={styles.label}>照射目的</Text>
          <Text style={styles.val}>{props.tani || "—"}</Text>
        </View>
        <View style={styles.row}>
          <Text style={styles.label}>指定パス</Text>
          <Text style={styles.val}>{props.pass || "—"}</Text>
          <Text style={styles.label}>計算式</Text>
          <Text style={styles.val}>{props.sikiKodo || "—"}</Text>
        </View>
        <View style={styles.row}>
          <Text style={styles.label}>素子異常</Text>
          <Text style={styles.val}>{props.sosi || "—"}</Text>
          <Text style={styles.label}>素子厚さ</Text>
          <Text style={styles.val}>{props.atusa || "—"}</Text>
        </View>
        <View style={styles.row}>
          <Text style={styles.label}>線量計種類</Text>
          <Text style={styles.val}>{props.senKind || "—"}</Text>
          <Text style={styles.label}>線量計厚さ</Text>
          <Text style={styles.val}>{props.atusaT || "—"}</Text>
        </View>
        <View style={styles.row}>
          <Text style={styles.label}>計算式コード</Text>
          <Text style={styles.val}>{props.keisask || "—"}</Text>
        </View>

        <Text style={styles.section}>照射計画（最大8受付）</Text>
        <View style={styles.tableHeader}>
          <Text style={styles.cellNarrow}>No</Text>
          <Text style={styles.cell}>受付</Text>
          <Text style={styles.cell}>RIC開始</Text>
          <Text style={styles.cell}>RIC終了</Text>
          <Text style={styles.cell}>指定線量</Text>
          <Text style={styles.cell}>下限</Text>
          <Text style={styles.cell}>上限</Text>
          <Text style={styles.cell}>実パス</Text>
          <Text style={styles.cell}>会社名</Text>
        </View>
        {props.planRows.map((r, i) => (
          <View key={i} style={styles.tableRow} wrap={false}>
            <Text style={styles.cellNarrow}>{i + 1}</Text>
            <Text style={styles.cell}>{r.ukeNo}</Text>
            <Text style={styles.cell}>{r.ricNoS}</Text>
            <Text style={styles.cell}>{r.ricNoE}</Text>
            <Text style={styles.cell}>{r.siteiSn}</Text>
            <Text style={styles.cell}>{r.kagen}</Text>
            <Text style={styles.cell}>{r.joug}</Text>
            <Text style={styles.cell}>{r.jituP}</Text>
            <Text style={styles.cell}>{r.kaiName}</Text>
          </View>
        ))}

        <Text style={styles.section}>追加照射（2パス単位の算出）</Text>
        <View style={styles.tableHeader}>
          <Text style={styles.cellNarrow}>No</Text>
          <Text style={styles.cell}>受付</Text>
          <Text style={styles.cell}>線量不足量</Text>
          <Text style={styles.cell}>上限からの余裕</Text>
          <Text style={styles.cell}>追加パス数</Text>
          <Text style={styles.cell}>会社名</Text>
        </View>
        {props.planRows.map((r, i) => {
          const kagen = parseNum(r.kagen);
          const joug = parseNum(r.joug);
          const shortage = kagen > 0 && measured > 0 ? kagen - measured : 0;
          const margin = joug > 0 && measured > 0 ? joug - measured : 0;
          const extra =
            r.ukeNo.trim() === ""
              ? 0
              : extraPassesRic2(
                  Math.max(0, shortage),
                  measured,
                  jituP1 > 0 ? jituP1 : parseNum(r.jituP),
                );
          return (
            <View key={`a-${i}`} style={styles.tableRow} wrap={false}>
              <Text style={styles.cellNarrow}>{i + 1}</Text>
              <Text style={styles.cell}>{r.ukeNo || "—"}</Text>
              <Text style={styles.cell}>
                {shortage > 0 ? shortage.toFixed(3) : "—"}
              </Text>
              <Text style={styles.cell}>
                {margin !== 0 ? margin.toFixed(3) : "—"}
              </Text>
              <Text style={styles.cell}>{extra > 0 ? String(extra) : "—"}</Text>
              <Text style={styles.cell}>{r.kaiName}</Text>
            </View>
          );
        })}

        <Text
          style={{ marginTop: 16, fontSize: 7, color: "#444" }}
          fixed
        >
          本書は Web 版 MVP です。Excel 様式 G2-12 の完全準拠はスキーマ拡張後に追加できます。
        </Text>
      </Page>
    </Document>
  );
}
