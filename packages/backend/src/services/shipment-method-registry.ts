/** rowkey 規約: shipment:${kaisyacd} — JSON { hikitori, housyube } */

export const SHIPMENT_ROWKEY_PREFIX = "shipment:" as const;

export function shipmentRowkey(kaisyacd: string): string {
  return `${SHIPMENT_ROWKEY_PREFIX}${kaisyacd.trim()}`;
}

export function parseShipmentJson(json: unknown): { hikitori: string; housyube: string } {
  if (!json || typeof json !== "object") return { hikitori: "", housyube: "" };
  const o = json as Record<string, unknown>;
  return {
    hikitori: typeof o.hikitori === "string" ? o.hikitori : "",
    housyube: typeof o.housyube === "string" ? o.housyube : "",
  };
}

/** SEHMST の max(syouho) を数値化（0 は不要扱い） */
export function syouhoMaxNumeric(v: string | null | undefined): number {
  if (v == null || String(v).trim() === "") return 0;
  const n = Number(String(v).trim());
  return Number.isFinite(n) ? n : 0;
}
