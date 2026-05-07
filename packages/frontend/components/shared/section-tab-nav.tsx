"use client";

import Link from "next/link";
import { usePathname } from "next/navigation";

import { cn } from "@/lib/utils";

export type SectionTab = { href: string; label: string };

export function SectionTabNav({
  tabs,
  ariaLabel,
}: {
  tabs: readonly SectionTab[];
  ariaLabel?: string;
}) {
  const pathname = usePathname();

  return (
    <nav
      className="mb-6 rounded-xl border border-border/70 bg-card/50 px-2 py-2"
      aria-label={ariaLabel ?? "セクション内タブ"}
    >
      <div className="flex flex-wrap gap-1">
        {tabs.map((t) => {
          const active = pathname === t.href;
          return (
            <Link
              key={t.href}
              href={t.href}
              className={cn(
                "rounded-lg px-3 py-2 text-sm font-medium transition-colors",
                active
                  ? "bg-terracotta/15 text-terracotta-dark shadow-sm ring-1 ring-terracotta/25"
                  : "text-ink-muted hover:bg-cream-200 hover:text-ink",
              )}
            >
              {t.label}
            </Link>
          );
        })}
      </div>
    </nav>
  );
}
