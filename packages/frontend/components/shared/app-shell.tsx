"use client";

import Link from "next/link";
import { usePathname } from "next/navigation";
import { useState } from "react";

import { Button } from "@/components/ui/button";
import { NAV } from "@/lib/nav-config";
import { cn } from "@/lib/utils";

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
