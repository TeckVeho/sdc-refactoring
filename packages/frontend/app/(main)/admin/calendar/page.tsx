"use client";

import { useMutation, useQuery, useQueryClient } from "@tanstack/react-query";
import { addMonths, format, subMonths } from "date-fns";
import { useEffect, useState } from "react";

import { Button } from "@/components/ui/button";
import { apiFetch } from "@/lib/api";

type HolidayRow = { ymd: string; reason: string | null };
type ListResponse = { data: HolidayRow[] };

function toYyyymm(d: Date): string {
  return format(d, "yyyyMM");
}

function ymdToLabel(ymd: string): string {
  if (ymd.length !== 8) return ymd;
  return `${ymd.slice(0, 4)}/${ymd.slice(4, 6)}/${ymd.slice(6, 8)}`;
}

export default function AdminCalendarPage() {
  const queryClient = useQueryClient();
  const [viewMonth, setViewMonth] = useState(() => {
    const d = new Date();
    d.setDate(1);
    return d;
  });
  const [newYmd, setNewYmd] = useState("");
  const [newReason, setNewReason] = useState("");
  const [formError, setFormError] = useState<string | null>(null);

  const yyyymm = toYyyymm(viewMonth);
  const listQuery = useQuery({
    queryKey: ["admin-calendar", yyyymm],
    queryFn: () =>
      apiFetch<ListResponse>(`/api/admin/calendar?from=${yyyymm}&to=${yyyymm}`),
  });

  const createMutation = useMutation({
    mutationFn: (body: { ymd: string; reason?: string }) =>
      apiFetch<{ data: HolidayRow }>("/api/admin/calendar", {
        method: "POST",
        body: JSON.stringify(body),
      }),
    onSuccess: () => {
      setFormError(null);
      setNewYmd("");
      setNewReason("");
      void queryClient.invalidateQueries({ queryKey: ["admin-calendar"] });
    },
    onError: (e: Error) => setFormError(e.message),
  });

  const patchMutation = useMutation({
    mutationFn: ({ ymd, reason }: { ymd: string; reason: string }) =>
      apiFetch<{ data: HolidayRow }>(`/api/admin/calendar/${ymd}`, {
        method: "PATCH",
        body: JSON.stringify({ reason }),
      }),
    onSuccess: () => {
      void queryClient.invalidateQueries({ queryKey: ["admin-calendar"] });
    },
  });

  const deleteMutation = useMutation({
    mutationFn: (ymd: string) =>
      apiFetch<{ data: { ok: boolean; ymd: string } }>(`/api/admin/calendar/${ymd}`, {
        method: "DELETE",
      }),
    onSuccess: () => {
      void queryClient.invalidateQueries({ queryKey: ["admin-calendar"] });
    },
  });

  const onAdd = (e: React.FormEvent) => {
    e.preventDefault();
    setFormError(null);
    const digits = newYmd.replace(/\D/g, "");
    if (digits.length !== 8) {
      setFormError("日付は8桁（YYYYMMDD）で入力してください。");
      return;
    }
    createMutation.mutate({
      ymd: digits,
      reason: newReason.trim() === "" ? undefined : newReason.trim(),
    });
  };

  return (
    <div className="space-y-8 p-6">
      <div>
        <h1 className="text-2xl font-semibold text-ink">休日カレンダー</h1>
        <p className="mt-1 text-sm text-ink-muted">
          会社休祭日（ExYasumiX）の登録・更新・削除を行います。
        </p>
      </div>

      <div className="flex flex-wrap items-center gap-3">
        <Button
          type="button"
          variant="secondary"
          size="sm"
          onClick={() => setViewMonth((d) => subMonths(d, 1))}
        >
          前月
        </Button>
        <span className="min-w-[8rem] text-center text-base font-medium text-ink">
          {format(viewMonth, "yyyy年M月")}
        </span>
        <Button
          type="button"
          variant="secondary"
          size="sm"
          onClick={() => setViewMonth((d) => addMonths(d, 1))}
        >
          次月
        </Button>
      </div>

      {listQuery.isError ? (
        <p className="text-sm text-terracotta">一覧の取得に失敗しました。ログイン状態をご確認ください。</p>
      ) : null}
      {listQuery.isLoading ? <p className="text-sm text-ink-muted">読み込み中…</p> : null}

      <div className="overflow-x-auto rounded-2xl border border-cream-300 bg-cream-50/80">
        <table className="w-full min-w-[36rem] border-collapse text-left text-sm">
          <thead>
            <tr className="border-b border-cream-300 bg-cream-100/80">
              <th className="px-4 py-3 font-semibold text-ink">日付 (YYYYMMDD)</th>
              <th className="px-4 py-3 font-semibold text-ink">理由</th>
              <th className="w-40 px-4 py-3 font-semibold text-ink">操作</th>
            </tr>
          </thead>
          <tbody>
            {(listQuery.data?.data ?? []).map((row) => (
              <ReasonRow
                key={row.ymd}
                row={row}
                onUpdate={(reason) => patchMutation.mutate({ ymd: row.ymd, reason })}
                onDelete={() => {
                  if (typeof window !== "undefined" && window.confirm(`${row.ymd} を削除しますか？`)) {
                    deleteMutation.mutate(row.ymd);
                  }
                }}
                updating={patchMutation.isPending}
                deleting={deleteMutation.isPending}
              />
            ))}
            {(listQuery.data?.data ?? []).length === 0 && !listQuery.isLoading ? (
              <tr>
                <td colSpan={3} className="px-4 py-6 text-center text-ink-muted">
                  この月に登録された休日はありません。
                </td>
              </tr>
            ) : null}
          </tbody>
        </table>
      </div>

      <form
        onSubmit={onAdd}
        className="max-w-xl space-y-4 rounded-2xl border border-cream-300 bg-cream-50/80 p-6"
      >
        <h2 className="text-lg font-semibold text-ink">休日を追加</h2>
        <div className="grid gap-4 sm:grid-cols-2">
          <div className="space-y-1">
            <label className="text-sm font-medium text-ink" htmlFor="new-ymd">
              日付
            </label>
            <input
              id="new-ymd"
              className="w-full rounded-xl border border-cream-300 bg-white px-3 py-2 text-ink outline-none ring-terracotta/0 transition focus:ring-2"
              placeholder="例: 20260429"
              value={newYmd}
              onChange={(ev) => setNewYmd(ev.target.value)}
              inputMode="numeric"
              maxLength={10}
            />
            <p className="text-xs text-ink-muted">8桁 YYYYMMDD（区切りなし、または区切り付き可）</p>
          </div>
          <div className="space-y-1">
            <label className="text-sm font-medium text-ink" htmlFor="new-reason">
              理由（任意）
            </label>
            <input
              id="new-reason"
              className="w-full rounded-xl border border-cream-300 bg-white px-3 py-2 text-ink outline-none ring-terracotta/0 transition focus:ring-2"
              placeholder="例: 社内休日"
              value={newReason}
              onChange={(ev) => setNewReason(ev.target.value)}
            />
          </div>
        </div>
        {formError ? <p className="text-sm text-terracotta">{formError}</p> : null}
        <Button type="submit" disabled={createMutation.isPending}>
          {createMutation.isPending ? "登録中…" : "登録"}
        </Button>
      </form>
    </div>
  );
}

function ReasonRow({
  row,
  onUpdate,
  onDelete,
  updating,
  deleting,
}: {
  row: HolidayRow;
  onUpdate: (reason: string) => void;
  onDelete: () => void;
  updating: boolean;
  deleting: boolean;
}) {
  const [reason, setReason] = useState(() => row.reason ?? "");

  useEffect(() => {
    setReason(row.reason ?? "");
  }, [row.ymd, row.reason]);

  return (
    <tr className="border-b border-cream-200/80 last:border-0">
      <td className="px-4 py-3 font-mono text-ink" title={row.ymd}>
        {ymdToLabel(row.ymd)}
      </td>
      <td className="px-4 py-2">
        <input
          className="w-full min-w-0 rounded-lg border border-cream-300 bg-white px-2 py-1.5 text-ink outline-none focus:ring-2 focus:ring-terracotta/30"
          value={reason}
          onChange={(e) => setReason(e.target.value)}
        />
      </td>
      <td className="px-4 py-2">
        <div className="flex flex-wrap gap-2">
          <Button
            type="button"
            size="sm"
            variant="secondary"
            disabled={updating || (row.reason ?? "") === reason}
            onClick={() => onUpdate(reason)}
          >
            更新
          </Button>
          <Button
            type="button"
            size="sm"
            variant="outline"
            disabled={deleting}
            onClick={onDelete}
          >
            削除
          </Button>
        </div>
      </td>
    </tr>
  );
}
