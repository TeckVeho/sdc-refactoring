"use client";

import { useMutation, useQuery, useQueryClient } from "@tanstack/react-query";
import { gantt, type Link } from "dhtmlx-gantt";
import "dhtmlx-gantt/codebase/dhtmlxgantt.css";
import { useCallback, useEffect, useRef, useState } from "react";

import { apiFetch } from "@/lib/api";

type PlanDetail = {
  detailId: string;
  planId: string | null;
  json: unknown;
};

type PlanRow = {
  planId: string;
  json: unknown;
  details: PlanDetail[];
};

type PlansResponse = {
  data: {
    plans: PlanRow[];
  };
};

function isGanttPayload(v: unknown): v is { data?: unknown[]; links?: unknown[] } {
  return typeof v === "object" && v !== null && ("data" in v || "links" in v);
}

function normalizePayload(json: unknown): { data: object[]; links: object[] } {
  if (isGanttPayload(json) && Array.isArray(json.data)) {
    return {
      data: json.data as object[],
      links: Array.isArray(json.links) ? (json.links as object[]) : [],
    };
  }
  return { data: [], links: [] };
}

const SAVE_DEBOUNCE_MS = 600;

export function GanttClient() {
  const hostRef = useRef<HTMLDivElement | null>(null);
  const queryClient = useQueryClient();
  const saveTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const planIdRef = useRef<string | null>(null);
  const [initError, setInitError] = useState<string | null>(null);

  const {
    data,
    isLoading,
    isError,
    error,
    refetch,
  } = useQuery({
    queryKey: ["irradiation", "machine1", "plan"],
    queryFn: () => apiFetch<PlansResponse>("/api/irradiation/machine1/plan"),
  });

  const plans = data?.data.plans ?? [];
  const activePlanId = plans[0]?.planId ?? null;

  const createPlan = useMutation({
    mutationFn: () =>
      apiFetch<{ data: { planId: string; json: unknown } }>("/api/irradiation/machine1/plan", {
        method: "POST",
        body: JSON.stringify({}),
      }),
    onSuccess: () => {
      void queryClient.invalidateQueries({ queryKey: ["irradiation", "machine1", "plan"] });
    },
  });

  const savePlan = useMutation({
    mutationFn: ({ planId, json }: { planId: string; json: unknown }) =>
      apiFetch<{ data: { planId: string; json: unknown } }>(`/api/irradiation/machine1/plan/${encodeURIComponent(planId)}`, {
        method: "PUT",
        body: JSON.stringify({ json }),
      }),
    onError: () => {
      setInitError("保存に失敗しました。ネットワークまたは権限を確認してください。");
    },
  });

  const scheduleSave = useCallback(() => {
    const planId = planIdRef.current;
    if (!planId) return;
    if (saveTimerRef.current) clearTimeout(saveTimerRef.current);
    saveTimerRef.current = setTimeout(() => {
      saveTimerRef.current = null;
      try {
        const payload = gantt.serialize();
        savePlan.mutate({ planId, json: payload });
      } catch {
        setInitError("ガントデータのシリアライズに失敗しました。");
      }
    }, SAVE_DEBOUNCE_MS);
  }, [savePlan]);

  useEffect(() => {
    planIdRef.current = activePlanId;
  }, [activePlanId]);

  useEffect(() => {
    if (isLoading || !activePlanId || !hostRef.current) return;

    const node = hostRef.current;
    setInitError(null);

    gantt.config.date_format = "%Y-%m-%d";
    gantt.config.scale_unit = "day";
    gantt.config.step = 1;
    gantt.config.duration_unit = "day";
    gantt.config.min_column_width = 80;
    gantt.config.autosize = "y";
    gantt.config.row_height = 28;
    gantt.config.bar_height = 22;

    gantt.config.columns = [
      { name: "text", label: "タスク", width: 200, tree: true, resize: true },
      { name: "start_date", label: "開始日", align: "center", width: 100, resize: true },
      { name: "duration", label: "日数", align: "center", width: 64, resize: true },
    ];

    gantt.locale = {
      ...gantt.locale,
      date: {
        ...gantt.locale.date,
        month_full: [
          "1月", "2月", "3月", "4月", "5月", "6月",
          "7月", "8月", "9月", "10月", "11月", "12月",
        ],
        month_short: ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12"],
        day_full: ["日曜日", "月曜日", "火曜日", "水曜日", "木曜日", "金曜日", "土曜日"],
        day_short: ["日", "月", "火", "水", "木", "金", "土"],
      },
    };

    gantt.init(node);

    const plan = plans.find((p) => p.planId === activePlanId);
    const { data: taskData, links } = normalizePayload(plan?.json);
    gantt.parse({
      data: taskData.length ? taskData : [{ id: 1, text: "新規タスク", start_date: new Date().toISOString().slice(0, 10), duration: 1 }],
      links: links as Link[],
    });

    const dragEv = gantt.attachEvent("onAfterTaskDrag", () => {
      scheduleSave();
    });
    const updateEv = gantt.attachEvent("onAfterTaskUpdate", () => {
      scheduleSave();
    });

    return () => {
      gantt.detachEvent(dragEv);
      gantt.detachEvent(updateEv);
      gantt.clearAll();
      gantt.destructor();
      if (saveTimerRef.current) {
        clearTimeout(saveTimerRef.current);
        saveTimerRef.current = null;
      }
    };
  }, [isLoading, activePlanId, plans, scheduleSave]);

  if (isLoading) {
    return (
      <div className="rounded-xl border border-cream-300 bg-cream-50/80 p-8 text-center text-sm text-ink-muted">
        計画データを読み込み中…
      </div>
    );
  }

  if (isError) {
    return (
      <div className="rounded-xl border border-red-200 bg-red-50/80 p-4 text-sm text-red-900">
        <p className="font-medium">読み込みエラー</p>
        <p className="mt-1">{error instanceof Error ? error.message : "不明なエラー"}</p>
        <button
          type="button"
          className="mt-3 rounded-lg border border-red-300 bg-white px-3 py-1.5 text-sm text-red-900 hover:bg-red-50"
          onClick={() => void refetch()}
        >
          再試行
        </button>
      </div>
    );
  }

  if (!activePlanId) {
    return (
      <div className="flex flex-col items-start gap-3 rounded-xl border border-cream-300 bg-cream-50/80 p-6">
        <p className="text-sm text-ink">登録されている照射計画がありません。初期計画を作成してください。</p>
        <button
          type="button"
          className="rounded-xl bg-terracotta px-4 py-2 text-sm font-medium text-white hover:bg-terracotta-dark disabled:opacity-60"
          disabled={createPlan.isPending}
          onClick={() => createPlan.mutate()}
        >
          {createPlan.isPending ? "作成中…" : "初期計画を作成"}
        </button>
        {createPlan.isError && (
          <p className="text-sm text-red-700">作成に失敗しました。再度お試しください。</p>
        )}
      </div>
    );
  }

  return (
    <div className="flex flex-col gap-3">
      {(initError || savePlan.isError) && (
        <div className="rounded-lg border border-amber-200 bg-amber-50/90 px-3 py-2 text-sm text-amber-950">
          {initError ?? "保存エラーが発生しました。"}
        </div>
      )}
      {savePlan.isPending && <p className="text-xs text-ink-muted">保存中…</p>}
      <div
        ref={hostRef}
        className="min-h-[420px] w-full overflow-auto rounded-xl border border-cream-300 bg-white p-2"
      />
    </div>
  );
}
