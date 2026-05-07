/**
 * 本番では無効。development / test でのみ、ログイン無しでも requireAuth を通す（スタブ応答とは別）。
 */
export function isRelaxedGuestAuth(): boolean {
  if (process.env.NODE_ENV === "production") return false;
  const v = process.env.SDC_RELAX_AUTH?.trim().toLowerCase();
  return v === "1" || v === "true" || v === "yes";
}
