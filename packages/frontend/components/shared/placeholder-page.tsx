export function PlaceholderPage({ title, path }: { title: string; path: string }) {
  return (
    <div className="mx-auto max-w-2xl rounded-2xl border border-cream-300 bg-cream-50 p-8 shadow-sm">
      <h1 className="text-xl font-semibold text-ink">{title}</h1>
      <p className="mt-2 font-mono text-sm text-ink-muted">{path}</p>
      <p className="mt-4 text-sm text-ink-muted">
        この画面は docs/implementation-plan.md のフェーズに沿って順次実装されます。
      </p>
    </div>
  );
}
