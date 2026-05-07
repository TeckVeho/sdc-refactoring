/**
 * 開発・デモ用: 有効時は `/api/*` の該当ルートが DB に依存せず固定サンプルを返します。
 * 本番では必ず無効にしてください。
 */
export function isSampleDataMode(): boolean {
  const v = process.env.SDC_SAMPLE_DATA?.trim().toLowerCase();
  return v === "1" || v === "true" || v === "yes";
}

/** アプリ起動時に一度だけログを出すフラグ（createApp で使用） */
let sampleWarned = false;

export function warnSampleDataModeOnce(): void {
  if (!isSampleDataMode() || sampleWarned) return;
  sampleWarned = true;
  console.warn(
    "[SDC] SDC_SAMPLE_DATA が有効です（対象 API はサンプル応答のみ）。本番環境では無効にしてください。",
  );
}
