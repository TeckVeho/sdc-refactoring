import type { PrismaClient } from "@prisma/client";

export type JmmVariant = "60" | "90";

export type JmmPayload = {
  variant: JmmVariant;
  maxRows: number;
  title: string;
  header: {
    uno: string;
    honnsuu: number | null;
    dose: number | null;
    sokuKiki: string | null;
    senKind: string | null;
    keicode: string | null;
    kagenn: number | null;
    jyogen: number | null;
    kainame: string | null;
    kaisyacd: string | null;
    sehncd: string | null;
    syono: string | null;
    syokind: string | null;
    bikou: string | null;
  };
  reference: {
    shipmentMethods: readonly string[];
    reportKinds: readonly string[];
  };
  senkinds: { kindcd: string; name: string | null }[];
  keicodes: { keicode: string; keisiki: string | null }[];
};

const SHIPMENT_METHODS = [
  "引取",
  "混載便",
  "保管品",
  "チャータ便",
  "納品",
  "品証扱い",
] as const;

const REPORT_KINDS = [
  "照射後Fax送信",
  "出荷後Fax送信",
  "照射後報告書発行",
  "出荷後報告書発行",
] as const;

function parseNum(raw: string | null | undefined): number | null {
  if (raw == null || raw === "") return null;
  const n = Number(String(raw).replace(/,/g, "").trim());
  return Number.isFinite(n) ? n : null;
}

export async function loadJmmPayload(
  prisma: PrismaClient,
  unoRaw: string,
  variant: JmmVariant,
): Promise<JmmPayload | null> {
  const uno = unoRaw.trim();
  if (!uno) return null;

  const [zk, sk1, senkinds, keicodes] = await Promise.all([
    prisma.zaiko.findUnique({ where: { uno } }),
    prisma.syouk1.findUnique({ where: { uno } }),
    prisma.senkind.findMany({ take: 40, orderBy: { kindcd: "asc" } }),
    prisma.keicode.findMany({ take: 120, orderBy: { keicode: "asc" } }),
  ]);

  if (!zk && !sk1) return null;

  const dose = parseNum(sk1?.siteisn ?? null);

  const title =
    variant === "60"
      ? "1 号機線量測定記録（60φPE 丸棒用）"
      : "1 号機線量測定記録（90φPE 丸棒）";

  return {
    variant,
    maxRows: variant === "60" ? 540 : 300,
    title,
    header: {
      uno,
      honnsuu: zk?.nyukasu ?? zk?.incnt ?? null,
      dose,
      sokuKiki: null,
      senKind: sk1?.syokind ?? null,
      keicode: null,
      kagenn: null,
      jyogen: null,
      kainame: zk?.kainame ?? sk1?.kainame ?? null,
      kaisyacd: zk?.kaisyacd ?? null,
      sehncd: zk?.sehncd ?? null,
      syono: sk1?.syono ?? null,
      syokind: sk1?.syokind ?? null,
      bikou: sk1?.bikou ?? null,
    },
    reference: {
      shipmentMethods: SHIPMENT_METHODS,
      reportKinds: REPORT_KINDS,
    },
    senkinds: senkinds.map((s) => ({ kindcd: s.kindcd, name: s.name ?? null })),
    keicodes: keicodes.map((k) => ({ keicode: k.keicode, keisiki: k.keisiki ?? null })),
  };
}
