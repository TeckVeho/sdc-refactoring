import { Machine1PlanGanttDynamic } from "./gantt-dynamic";

export default function IrradiationMachine1PlanPage() {
  return (
    <div className="flex flex-col gap-6 p-4 md:p-6">
      <div>
        <h1 className="text-xl font-semibold text-ink">1号機照射計画</h1>
        <p className="mt-1 text-sm text-ink-muted">
          タイムスケールは日単位です。バーをドラッグして日付を変更すると、自動で保存されます（数秒遅延）。
        </p>
      </div>
      <Machine1PlanGanttDynamic />
    </div>
  );
}
