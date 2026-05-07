/**
 * UI 側の公開デモフラグ（バックエンドの SDC_SAMPLE_DATA と揃える想定）。
 * ログイン画面スキップ・バナー表示に使用します。
 */
export function isPublicSampleDataDemo(): boolean {
  const v = process.env.NEXT_PUBLIC_SDC_SAMPLE_DATA?.trim().toLowerCase();
  return v === "1" || v === "true" || v === "yes";
}
