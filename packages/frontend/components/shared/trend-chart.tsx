"use client";

import {
  CartesianGrid,
  Legend,
  Line,
  LineChart,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from "recharts";

const LINE_COLORS = [
  "#c0634c",
  "#4c8bc0",
  "#4cc07a",
  "#c0a44c",
  "#9b4cc0",
  "#4cc0b8",
];

interface TrendDataItem {
  time: string;
  /** 単一系列のとき。複数系列は dataKeys のみでも可 */
  value?: number;
  [key: string]: unknown;
}

interface TrendChartProps {
  data: TrendDataItem[];
  dataKeys: string[];
  xAxisKey?: string;
  unit?: string;
  height?: number;
}

export function TrendChart({
  data,
  dataKeys,
  xAxisKey = "time",
  unit,
  height = 300,
}: TrendChartProps) {
  return (
    <ResponsiveContainer width="100%" height={height}>
      <LineChart data={data} margin={{ top: 8, right: 24, left: 0, bottom: 8 }}>
        <CartesianGrid strokeDasharray="3 3" stroke="#e8e0d8" />
        <XAxis
          dataKey={xAxisKey}
          tick={{ fontSize: 12, fill: "#6b5e52" }}
          tickLine={false}
          axisLine={{ stroke: "#d4c8be" }}
        />
        <YAxis
          tick={{ fontSize: 12, fill: "#6b5e52" }}
          tickLine={false}
          axisLine={{ stroke: "#d4c8be" }}
          unit={unit}
        />
        <Tooltip
          contentStyle={{
            borderRadius: "0.75rem",
            border: "1px solid #d4c8be",
            background: "#faf8f5",
            fontSize: 13,
          }}
          formatter={(val) => [`${String(val)}${unit ?? ""}`, undefined]}
        />
        {dataKeys.length > 1 && <Legend wrapperStyle={{ fontSize: 12 }} />}
        {dataKeys.map((key, i) => (
          <Line
            key={key}
            type="monotone"
            dataKey={key}
            stroke={LINE_COLORS[i % LINE_COLORS.length]}
            strokeWidth={2}
            dot={false}
            activeDot={{ r: 4 }}
          />
        ))}
      </LineChart>
    </ResponsiveContainer>
  );
}
