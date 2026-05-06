"use client";

import dynamic from "next/dynamic";

const GanttClient = dynamic(() => import("./gantt-client").then((m) => ({ default: m.GanttClient })), {
  ssr: false,
  loading: () => (
    <div className="rounded-xl border border-cream-300 bg-cream-50/80 p-8 text-center text-sm text-ink-muted">
      ガントチャートを読み込み中…
    </div>
  ),
});

export function Machine1PlanGanttDynamic() {
  return <GanttClient />;
}
