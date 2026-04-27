"use client";

import { useQuery } from "@tanstack/react-query";
import Link from "next/link";

import { Button } from "@/components/ui/button";
import { apiFetch } from "@/lib/api";
import { NAV } from "@/lib/nav-config";

type NoticeResponse = { data: { text: string } };

function DashboardNotice() {
  const { data, isError } = useQuery({
    queryKey: ["dashboard", "notice"],
    queryFn: () => apiFetch<NoticeResponse>("/api/dashboard/notice"),
    retry: false,
  });
  if (isError) return null;
  const text = data?.data.text?.trim() ?? "";
  if (!text) return null;
  return (
    <div
      className="rounded-2xl border border-amber-200/80 bg-amber-50/90 px-4 py-3 text-sm text-amber-950 shadow-sm"
      role="status"
    >
      <p className="font-medium text-amber-900">お知らせ</p>
      <p className="mt-1 whitespace-pre-wrap text-amber-950/90">{text}</p>
    </div>
  );
}

export default function DashboardPage() {
  return (
    <div className="mx-auto max-w-6xl space-y-8">
      <header className="space-y-2">
        <h1 className="text-2xl font-semibold text-ink md:text-3xl">SDC 照射管理</h1>
        <p className="max-w-2xl text-ink-muted">
          Web 版の業務入口です。下のカードを開くと各業務に進みます。常に表示のメニューは左のナビと同じです。
        </p>
      </header>

      <DashboardNotice />

      <div className="space-y-10">
        {NAV.map((group) => (
          <section key={group.title} className="space-y-3">
            <h2 className="text-sm font-semibold uppercase tracking-wide text-ink-muted">{group.title}</h2>
            <div className="grid grid-cols-1 gap-3 sm:grid-cols-2 lg:grid-cols-3">
              {group.items.map((item) => (
                <Button key={item.href} asChild variant="outline" className="h-auto min-h-[2.75rem] w-full justify-center px-4 py-3 text-center text-sm font-medium whitespace-normal">
                  <Link href={item.href}>{item.label}</Link>
                </Button>
              ))}
            </div>
          </section>
        ))}
      </div>
    </div>
  );
}
