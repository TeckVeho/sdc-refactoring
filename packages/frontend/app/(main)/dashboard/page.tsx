"use client";

import { useQuery } from "@tanstack/react-query";
import Link from "next/link";

import { Button } from "@/components/ui/button";
import { apiFetch } from "@/lib/api";
import { PRIMARY_NAV } from "@/lib/nav-config";

type NoticeResponse = { data: { text: string } };

function DashboardNotice() {
  const { data, isError } = useQuery({
    queryKey: ["dashboard", "notice"],
    queryFn: () => apiFetch<NoticeResponse>("/api/dashboard/notice"),
    retry: false,
  });
  if (isError) return null;
  const text = String(data?.data?.text ?? "").trim();
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
          Web 版の業務入口です。メインは左の8項目です。各項目内はタブで画面を切り替えます。
        </p>
      </header>

      <DashboardNotice />

      <div className="grid grid-cols-1 gap-3 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4">
        {PRIMARY_NAV.map((item) => {
          const Icon = item.icon;
          return (
            <Button
              key={item.href}
              asChild
              variant="outline"
              className="h-auto min-h-[2.75rem] w-full justify-start gap-3 px-4 py-3 text-left text-sm font-medium whitespace-normal"
            >
              <Link href={item.href} className="flex items-start gap-3">
                <Icon className="mt-0.5 h-4 w-4 shrink-0 text-terracotta" aria-hidden />
                <span>{item.label}</span>
              </Link>
            </Button>
          );
        })}
      </div>
    </div>
  );
}
