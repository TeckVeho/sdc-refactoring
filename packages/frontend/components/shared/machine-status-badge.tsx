import { cn } from "@/lib/utils";

export type MachineStatus = "running" | "stopped" | "error" | "unknown";

interface MachineStatusBadgeProps {
  status: MachineStatus;
  machineName?: string;
}

const STATUS_CONFIG: Record<MachineStatus, { label: string; className: string }> = {
  running: {
    label: "照射中",
    className: "bg-green-100 text-green-800 border-green-200",
  },
  stopped: {
    label: "停止",
    className: "bg-gray-100 text-gray-700 border-gray-200",
  },
  error: {
    label: "異常",
    className: "bg-red-100 text-red-800 border-red-200",
  },
  unknown: {
    label: "不明",
    className: "bg-yellow-100 text-yellow-800 border-yellow-200",
  },
};

export function MachineStatusBadge({ status, machineName }: MachineStatusBadgeProps) {
  const config = STATUS_CONFIG[status];

  return (
    <span
      className={cn(
        "inline-flex items-center gap-1.5 rounded-full border px-2.5 py-0.5 text-xs font-medium",
        config.className,
      )}
    >
      <span
        className={cn("h-1.5 w-1.5 rounded-full", {
          "bg-green-500 animate-pulse": status === "running",
          "bg-gray-400": status === "stopped",
          "bg-red-500": status === "error",
          "bg-yellow-500": status === "unknown",
        })}
      />
      {machineName ? `${machineName}: ` : ""}
      {config.label}
    </span>
  );
}
