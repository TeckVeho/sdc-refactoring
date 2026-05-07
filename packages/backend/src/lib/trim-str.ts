/**
 * JSON 経由や DB ドライバの揺れで string 以外が来ても落ちない trim。
 */
export function trimStr(v: unknown): string {
  if (v == null) return "";
  return typeof v === "string" ? v.trim() : String(v).trim();
}
