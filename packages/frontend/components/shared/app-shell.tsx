"use client";

import Link from "next/link";
import { usePathname } from "next/navigation";
import { useState } from "react";

import { Button } from "@/components/ui/button";
import { cn } from "@/lib/utils";

type NavItem = { href: string; label: string };

type NavGroup = { title: string; items: NavItem[] };

const NAV: NavGroup[] = [
  {
    title: "ダッシュボード",
    items: [{ href: "/dashboard", label: "ホームメニュー" }],
  },
  {
    title: "照射管理",
    items: [
      { href: "/irradiation/machine1/monitor", label: "1号機照射情報" },
      { href: "/irradiation/machine1/plan", label: "1号機照射計画" },
      { href: "/irradiation/machine1/work-order", label: "1号機作業指図書" },
      { href: "/irradiation/machine-status", label: "装置運転状況" },
      { href: "/irradiation/production", label: "生産情報一覧" },
    ],
  },
  {
    title: "入出荷",
    items: [
      { href: "/inventory/shipment-summary", label: "入出荷集計" },
      { href: "/inventory/company-arrival", label: "会社別入荷集計" },
      { href: "/inventory/month-end", label: "月末在庫集計" },
      { href: "/inventory/customer-stock", label: "顧客在庫報告" },
      { href: "/inventory/ric3-repack", label: "Ric3詰替作業" },
    ],
  },
  {
    title: "線量管理",
    items: [
      { href: "/dosimetry/oven-management", label: "R3オーブン管理" },
      { href: "/dosimetry/dose-search", label: "線量検索" },
      { href: "/dosimetry/irradiation-results", label: "照射実績表示" },
      { href: "/dosimetry/jmm60", label: "JMM60φ記入用紙" },
      { href: "/dosimetry/jmm90", label: "JMM90φ記入用紙" },
      { href: "/dosimetry/dose-rate-calc", label: "線量率計算" },
    ],
  },
  {
    title: "報告書",
    items: [
      { href: "/reports/ric2-dose-shortage", label: "RIC2線量不足報告書" },
      { href: "/reports/ric3-dose-shortage", label: "RIC3線量不足報告書" },
      { href: "/reports/ric3-dose-shortage-sumiden", label: "RIC3（住電）" },
      { href: "/reports/shipment-method", label: "出荷方法報告書" },
      { href: "/reports/unissued-search", label: "報告書未発行検索" },
      { href: "/reports/gamma-results", label: "ガンマ照射課実績集計" },
    ],
  },
  {
    title: "管理・マスタ",
    items: [
      { href: "/admin/calendar", label: "カレンダー" },
      { href: "/admin/shift-schedule", label: "勤務表" },
      { href: "/admin/price-search", label: "単価検索" },
      { href: "/admin/db-browser", label: "DBブラウザ" },
      { href: "/admin/company-code-convert", label: "GM/EB会社コード変換" },
      { href: "/admin/shipment-method-verify", label: "出荷方法（検証）" },
      { href: "/admin/shipment-method-ss", label: "出荷方法（SS検証）" },
    ],
  },
];

export function AppShell({ children }: { children: React.ReactNode }) {
  const pathname = usePathname();
  const [open, setOpen] = useState(false);

  return (
    <div className="flex min-h-screen bg-cream-100">
      <aside
        className={cn(
          "fixed inset-y-0 left-0 z-40 w-72 border-r border-cream-300 bg-cream-50/95 backdrop-blur-sm transition-transform lg:static lg:translate-x-0",
          open ? "translate-x-0" : "-translate-x-full lg:translate-x-0",
        )}
      >
        <div className="flex h-14 items-center border-b border-cream-300 px-4">
          <Link href="/dashboard" className="text-lg font-semibold text-ink">
            SDC <span className="text-sm font-normal text-ink-muted">照射管理</span>
          </Link>
        </div>
        <nav className="h-[calc(100vh-3.5rem)] overflow-y-auto px-3 py-4">
          {NAV.map((group) => (
            <div key={group.title} className="mb-6">
              <p className="mb-2 px-2 text-xs font-semibold uppercase tracking-wide text-ink-muted">
                {group.title}
              </p>
              <ul className="space-y-0.5">
                {group.items.map((item) => {
                  const active = pathname === item.href;
                  return (
                    <li key={item.href}>
                      <Link
                        href={item.href}
                        onClick={() => setOpen(false)}
                        className={cn(
                          "block rounded-xl px-3 py-2 text-sm transition-colors",
                          active
                            ? "bg-terracotta/15 font-medium text-terracotta-dark"
                            : "text-ink hover:bg-cream-200",
                        )}
                      >
                        {item.label}
                      </Link>
                    </li>
                  );
                })}
              </ul>
            </div>
          ))}
        </nav>
      </aside>

      {open ? (
        <button
          type="button"
          className="fixed inset-0 z-30 bg-ink/20 lg:hidden"
          aria-label="メニューを閉じる"
          onClick={() => setOpen(false)}
        />
      ) : null}

      <div className="flex min-h-screen flex-1 flex-col lg:pl-0">
        <header className="sticky top-0 z-20 flex h-14 items-center gap-3 border-b border-cream-300 bg-cream-50/90 px-4 backdrop-blur-sm">
          <Button type="button" variant="outline" size="sm" className="lg:hidden" onClick={() => setOpen((v) => !v)}>
            メニュー
          </Button>
          <span className="text-sm text-ink-muted">ラジエ工業 — 照射管理システム（Web）</span>
        </header>
        <main className="flex-1 p-6">{children}</main>
      </div>
    </div>
  );
}
