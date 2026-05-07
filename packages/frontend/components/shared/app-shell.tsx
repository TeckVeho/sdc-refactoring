"use client";

import { PanelLeftClose, PanelLeft } from "lucide-react";
import Link from "next/link";
import { usePathname } from "next/navigation";
import { useCallback, useEffect, useMemo, useState } from "react";

import { Button } from "@/components/ui/button";
import { isPublicSampleDataDemo } from "@/lib/demo-mode";
import { apiUrl } from "@/lib/api";
import {
  isNavActiveForPath,
  PRIMARY_NAV,
  type PrimaryNavItem,
} from "@/lib/nav-config";
import { cn } from "@/lib/utils";

const STORAGE_KEY = "sdc-sidebar-collapsed";

function useIsLg() {
  const [isLg, setIsLg] = useState<boolean | null>(null);

  useEffect(() => {
    const mq = window.matchMedia("(min-width: 1024px)");
    const update = () => setIsLg(mq.matches);
    update();
    mq.addEventListener("change", update);
    return () => mq.removeEventListener("change", update);
  }, []);

  return isLg;
}

const showSampleDemoBanner = isPublicSampleDataDemo();

export function AppShell({ children }: { children: React.ReactNode }) {
  const pathname = usePathname();
  const [demoGuestMismatch, setDemoGuestMismatch] = useState<boolean | null>(null);
  const [drawerOpen, setDrawerOpen] = useState(false);
  const [navCollapsed, setNavCollapsed] = useState(false);
  const isLg = useIsLg();

  const effectiveCompact = Boolean(navCollapsed && isLg);

  useEffect(() => {
    try {
      if (typeof window !== "undefined" && window.localStorage.getItem(STORAGE_KEY) === "1") {
        setNavCollapsed(true);
      }
    } catch {
      /* ignore */
    }
  }, []);

  const setCollapsedPersist = useCallback((next: boolean) => {
    setNavCollapsed(next);
    try {
      window.localStorage.setItem(STORAGE_KEY, next ? "1" : "0");
    } catch {
      /* ignore */
    }
  }, []);

  useEffect(() => {
    if (!showSampleDemoBanner) return;
    void fetch(apiUrl("/api/health"))
      .then(async (r) => (r.ok ? ((await r.json()) as { data?: { guestApiAccess?: boolean } }) : null))
      .then((j) => setDemoGuestMismatch(!Boolean(j?.data?.guestApiAccess)))
      .catch(() => setDemoGuestMismatch(true));
  }, []);

  const asideWidthClass = useMemo(() => {
    if (isLg === null) return "w-72";
    return effectiveCompact ? "lg:w-[4.5rem]" : "lg:w-72";
  }, [effectiveCompact, isLg]);

  return (
    <div className="flex min-h-screen bg-cream-100">
      <aside
        className={cn(
          "fixed inset-y-0 left-0 z-40 flex shrink-0 flex-col overflow-hidden border-r border-cream-300 bg-cream-50/95 shadow-sm backdrop-blur-sm transition-[width,transform] duration-200 ease-out",
          asideWidthClass,
          "w-72",
          drawerOpen ? "translate-x-0" : "-translate-x-full lg:translate-x-0",
          "lg:static lg:z-auto",
        )}
      >
        <div
          className={cn(
            "flex h-14 shrink-0 items-center border-b border-cream-300 px-4",
            effectiveCompact && "lg:justify-center lg:px-2",
          )}
        >
          <Link
            href="/dashboard"
            className={cn(
              "text-lg font-semibold text-ink transition-opacity hover:text-ink",
              effectiveCompact && "lg:truncate",
            )}
            title={effectiveCompact ? "SDC 照射管理 — ホームへ" : undefined}
          >
            <span className={cn(effectiveCompact && "lg:hidden")}>
              SDC <span className="text-sm font-normal text-ink-muted">照射管理</span>
            </span>
            <span className={cn("hidden", effectiveCompact && "lg:inline")}>SDC</span>
          </Link>
        </div>

        <nav
          className={cn(
            "flex min-h-0 flex-1 flex-col overflow-y-auto px-2 py-3",
            effectiveCompact && "lg:px-2 lg:pb-3",
          )}
        >
          {effectiveCompact ? (
            <CompactNav pathname={pathname} onNavigate={() => setDrawerOpen(false)} />
          ) : (
            <FullNav pathname={pathname} onNavigate={() => setDrawerOpen(false)} />
          )}
        </nav>

        <div className="hidden shrink-0 border-t border-cream-300 p-2 lg:block">
          <Button
            type="button"
            variant="outline"
            size="sm"
            className={cn(
              "w-full border-cream-300 text-ink-muted hover:bg-cream-200 hover:text-ink",
              navCollapsed ? "inline-flex h-9 w-9 shrink-0 items-center justify-center p-0" : "inline-flex items-center justify-center gap-2",
            )}
            onClick={() => setCollapsedPersist(!navCollapsed)}
            aria-label={navCollapsed ? "サイドバーを広げる" : "サイドバーをしまう"}
            title={navCollapsed ? "サイドバーを広げる" : "サイドバーをしまう"}
          >
            {navCollapsed ? (
              <PanelLeft className="h-4 w-4" />
            ) : (
              <>
                <PanelLeftClose className="h-4 w-4 shrink-0" />
                <span className="text-xs">しまう</span>
              </>
            )}
          </Button>
        </div>
      </aside>

      {drawerOpen ? (
        <button
          type="button"
          className="fixed inset-0 z-30 bg-ink/20 lg:hidden"
          aria-label="メニューを閉じる"
          onClick={() => setDrawerOpen(false)}
        />
      ) : null}

      <div className="flex min-h-screen min-w-0 flex-1 flex-col lg:pl-0">
        <header className="sticky top-0 z-20 flex h-14 shrink-0 items-center gap-3 border-b border-cream-300 bg-cream-50/90 px-4 backdrop-blur-sm">
          <Button
            type="button"
            variant="outline"
            size="sm"
            className="lg:hidden"
            onClick={() => setDrawerOpen((v) => !v)}
          >
            メニュー
          </Button>
          <span className="text-sm text-ink-muted">ラジエ工業 — 照射管理システム（Web）</span>
        </header>
        {showSampleDemoBanner ? (
          <div
            className="border-b border-amber-200 bg-amber-50 px-4 py-2 text-center text-xs text-amber-950"
            role="status"
          >
            <p className="font-medium">公開デモ表示（NEXT_PUBLIC_SDC_SAMPLE_DATA）</p>
            {demoGuestMismatch === true ? (
              <p className="mt-1 leading-relaxed text-red-950">
                バックエンドでゲスト API が無効のため一覧が Unauthorized になります。{" "}
                <code className="rounded bg-white/80 px-1 font-mono text-[11px]">SDC_SAMPLE_DATA=1</code>{" "}
                または{" "}
                <code className="rounded bg-white/80 px-1 font-mono text-[11px]">SDC_RELAX_AUTH=1</code>
                を <code className="font-mono text-[11px]">packages/backend/.env</code> に設定して再起動するか、
                <Link href="/login" className="underline decoration-amber-900/50 underline-offset-2">
                  ログイン
                </Link>
                してください。
              </p>
            ) : demoGuestMismatch === false ? (
              <p className="mt-1 opacity-95">
                バックエンドのゲスト/API 設定は検出済みです。スタブ運用時は保存は無効側の応答になります。
              </p>
            ) : (
              <p className="mt-1 text-amber-900/80">バックエンド状態を確認中…</p>
            )}
          </div>
        ) : null}
        <main className="min-w-0 flex-1 p-6">{children}</main>
      </div>
    </div>
  );
}

function NavItemLink({
  item,
  active,
  onNavigate,
}: {
  item: PrimaryNavItem;
  active: boolean;
  onNavigate: () => void;
}) {
  const Icon = item.icon;
  return (
    <Link
      href={item.href}
      onClick={onNavigate}
      title={item.label}
      className={cn(
        "flex items-center gap-3 rounded-xl border-l-[3px] px-3 py-2 text-sm transition-all duration-150",
        active
          ? "border-terracotta bg-terracotta/10 font-medium text-terracotta-dark shadow-sm ring-1 ring-terracotta/10"
          : "border-transparent text-ink hover:translate-x-0.5 hover:bg-cream-200 hover:text-ink",
      )}
    >
      <Icon className={cn("h-[1.125rem] w-[1.125rem] shrink-0", active ? "text-terracotta" : "text-ink-muted")} aria-hidden />
      <span className="min-w-0 leading-snug">{item.label}</span>
    </Link>
  );
}

function FullNav({
  pathname,
  onNavigate,
}: {
  pathname: string;
  onNavigate: () => void;
}) {
  return (
    <div className="flex flex-col gap-0.5">
      {PRIMARY_NAV.map((item) => {
        const active = isNavActiveForPath(pathname, item);
        return <NavItemLink key={item.href} item={item} active={active} onNavigate={onNavigate} />;
      })}
    </div>
  );
}

function CompactNav({ pathname, onNavigate }: { pathname: string; onNavigate: () => void }) {
  return (
    <div className="flex flex-col items-center gap-1">
      {PRIMARY_NAV.map((item) => {
        const active = isNavActiveForPath(pathname, item);
        const Icon = item.icon;
        return (
          <Link
            key={item.href}
            href={item.href}
            title={item.label}
            aria-label={item.label}
            onClick={onNavigate}
            className={cn(
              "flex h-10 w-10 shrink-0 items-center justify-center rounded-xl transition-all duration-150",
              active
                ? "bg-terracotta/15 text-terracotta-dark shadow-sm ring-1 ring-terracotta/15"
                : "text-ink-muted hover:bg-cream-200 hover:text-ink",
            )}
          >
            <Icon className="h-[1.125rem] w-[1.125rem]" aria-hidden />
          </Link>
        );
      })}
    </div>
  );
}
