"use client";

import type { ComponentType } from "react";
import { useEffect, useState } from "react";

/**
 * dhtmlx-gantt は SSR 不可。next/dynamic 経由だと開発時 HMR で
 * __webpack_modules__[moduleId] is not a function が出ることがあるため、
 * クライアント内の import() でのみ読み込む。
 */
export function Machine1PlanGanttLoader() {
  const [Gantt, setGantt] = useState<ComponentType | null>(null);

  useEffect(() => {
    let cancelled = false;
    void import("./gantt-client").then((m) => {
      if (!cancelled && m.default) {
        setGantt(() => m.default);
      }
    });
    return () => {
      cancelled = true;
    };
  }, []);

  if (!Gantt) {
    return (
      <div className="rounded-xl border border-cream-300 bg-cream-50/80 p-8 text-center text-sm text-ink-muted">
        ガントチャートを読み込み中…
      </div>
    );
  }

  return <Gantt />;
}
