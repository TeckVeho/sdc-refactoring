import Link from "next/link";

import { Button } from "@/components/ui/button";

export default function DashboardPage() {
  return (
    <div className="mx-auto max-w-3xl space-y-6">
      <div className="rounded-2xl border border-cream-300 bg-cream-50 p-8 shadow-sm">
        <h1 className="text-2xl font-semibold text-ink">ホームメニュー</h1>
        <p className="mt-2 text-ink-muted">
          左のナビから業務画面を開いてください。Phase 2 以降で各画面を実装予定です。
        </p>
        <div className="mt-6 flex flex-wrap gap-3">
          <Button asChild>
            <Link href="/irradiation/machine1/monitor">1号機照射情報へ</Link>
          </Button>
          <Button variant="secondary" asChild>
            <Link href="/inventory/shipment-summary">入出荷集計へ</Link>
          </Button>
        </div>
      </div>
    </div>
  );
}
