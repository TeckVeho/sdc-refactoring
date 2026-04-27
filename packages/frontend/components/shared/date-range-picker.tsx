"use client";

import { cn } from "@/lib/utils";

interface DateRangeValue {
  from: Date;
  to: Date;
}

interface DateRangePickerProps {
  value: DateRangeValue;
  onChange: (range: DateRangeValue) => void;
  label?: string;
  className?: string;
}

function toDateInputValue(date: Date): string {
  return date.toISOString().slice(0, 10);
}

export function DateRangePicker({ value, onChange, label, className }: DateRangePickerProps) {
  const handleFromChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const from = new Date(e.target.value);
    const to = from > value.to ? from : value.to;
    onChange({ from, to });
  };

  const handleToChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const to = new Date(e.target.value);
    const from = to < value.from ? to : value.from;
    onChange({ from, to });
  };

  return (
    <div className={cn("flex flex-wrap items-center gap-2", className)}>
      {label && <span className="text-sm font-medium text-ink">{label}</span>}
      <div className="flex items-center gap-2">
        <label className="flex items-center gap-1.5">
          <span className="text-sm text-ink-muted">開始</span>
          <input
            type="date"
            value={toDateInputValue(value.from)}
            onChange={handleFromChange}
            className="rounded-xl border border-cream-300 bg-cream-50 px-3 py-1.5 text-sm text-ink focus:outline-none focus:ring-2 focus:ring-terracotta/40"
          />
        </label>
        <span className="text-sm text-ink-muted">〜</span>
        <label className="flex items-center gap-1.5">
          <span className="text-sm text-ink-muted">終了</span>
          <input
            type="date"
            value={toDateInputValue(value.to)}
            min={toDateInputValue(value.from)}
            onChange={handleToChange}
            className="rounded-xl border border-cream-300 bg-cream-50 px-3 py-1.5 text-sm text-ink focus:outline-none focus:ring-2 focus:ring-terracotta/40"
          />
        </label>
      </div>
    </div>
  );
}
