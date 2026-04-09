"use client";

import {
  BarChart, Bar, LineChart, Line, AreaChart, Area,
  PieChart, Pie, Cell, RadarChart, Radar, PolarGrid, PolarAngleAxis,
  ScatterChart, Scatter, XAxis, YAxis, CartesianGrid, Tooltip, Legend,
  ResponsiveContainer, RadialBarChart, RadialBar, FunnelChart, Funnel,
  LabelList, Treemap,
} from "recharts";
import type { ChartConfig } from "@/components/dashboard/DashboardBuilder";
import type { ExcelSheet } from "@/lib/excel";

const PALETTE: Record<string, string[]> = {
  default: ["#0A0A0A", "#3A3A3A", "#6B6B6B", "#A1A1A1", "#D1D1D1"],
  blues:   ["#1e40af", "#2563eb", "#3b82f6", "#60a5fa", "#93c5fd"],
  greens:  ["#14532d", "#15803d", "#16a34a", "#22c55e", "#4ade80"],
  reds:    ["#7f1d1d", "#b91c1c", "#dc2626", "#ef4444", "#f87171"],
  purples: ["#3b0764", "#6d28d9", "#7c3aed", "#8b5cf6", "#a78bfa"],
  warm:    ["#78350f", "#d97706", "#f59e0b", "#fbbf24", "#fde68a"],
  cool:    ["#0c4a6e", "#0369a1", "#0ea5e9", "#38bdf8", "#7dd3fc"],
  neon:    ["#4ade80", "#22d3ee", "#a78bfa", "#f472b6", "#facc15"],
};

function getColors(scheme: string) {
  return PALETTE[scheme] || PALETTE.default;
}

function prepareData(chart: ChartConfig, sheet: ExcelSheet) {
  const xCol = sheet.columns.find((c) => c.header === chart.xColumn);
  const yCols = sheet.columns.filter((c) => chart.yColumns.includes(c.header));
  if (!xCol || yCols.length === 0) return [];

  let rows = sheet.data.map((row) => {
    const entry: Record<string, string | number | null> = {
      name: String(row[xCol.header] ?? ""),
    };
    yCols.forEach((yc) => {
      const val = row[yc.header];
      entry[yc.header] = val !== null && val !== undefined ? parseFloat(String(val)) || 0 : 0;
    });
    return entry;
  });

  // Always group by X column to prevent duplicate labels
  const grouped: Record<string, Record<string, number[]>> = {};
  rows.forEach((row) => {
    const key = String(row.name);
    if (!grouped[key]) grouped[key] = {};
    yCols.forEach((yc) => {
      if (!grouped[key][yc.header]) grouped[key][yc.header] = [];
      grouped[key][yc.header].push(Number(row[yc.header]) || 0);
    });
  });
  rows = Object.entries(grouped).map(([name, cols]) => {
    const entry: Record<string, string | number | null> = { name };
    Object.entries(cols).forEach(([col, vals]) => {
      const agg = chart.aggregation;
      if (agg === "none") entry[col] = vals[0]; // First value when no aggregation
      else if (agg === "sum") entry[col] = vals.reduce((a, b) => a + b, 0);
      else if (agg === "avg") entry[col] = vals.reduce((a, b) => a + b, 0) / vals.length;
      else if (agg === "count") entry[col] = vals.length;
      else if (agg === "min") entry[col] = Math.min(...vals);
      else if (agg === "max") entry[col] = Math.max(...vals);
    });
    return entry;
  });

  if (chart.limit && chart.limit > 0) rows = rows.slice(0, chart.limit);
  if (chart.sortBy === "asc") rows.sort((a, b) => Number(a[chart.yColumns[0]]) - Number(b[chart.yColumns[0]]));
  if (chart.sortBy === "desc") rows.sort((a, b) => Number(b[chart.yColumns[0]]) - Number(a[chart.yColumns[0]]));

  return rows;
}

interface ChartRendererProps {
  chart: ChartConfig;
  sheet: ExcelSheet;
}

export function ChartRenderer({ chart, sheet }: ChartRendererProps) {
  const data = prepareData(chart, sheet);
  const colors = getColors(chart.colorScheme);
  const yKeys = chart.yColumns;
  const { showLegend, showGrid, showLabels, stacked } = chart;

  if (data.length === 0) return <EmptyState />;

  const commonProps = {
    data,
    margin: { top: 4, right: 8, bottom: 4, left: 0 },
  };

  const xAxisLabel = chart.xColumn || "";
  const yAxisLabel = yKeys.length === 1 ? yKeys[0] : yKeys.length > 1 ? "Value" : "";
  const axisLabelStyle = { fontSize: 10, fill: "#999" };

  switch (chart.type) {
    case "bar":
    case "bar_horizontal":
      return (
        <ResponsiveContainer width="100%" height="100%">
          <BarChart
            {...commonProps}
            layout={chart.type === "bar_horizontal" ? "vertical" : "horizontal"}
            margin={{ top: 8, right: 12, bottom: xAxisLabel ? 36 : 8, left: yAxisLabel ? 40 : 16 }}
          >
            {showGrid && <CartesianGrid strokeDasharray="3 3" stroke="#F0F0F0" />}
            {chart.type === "bar_horizontal" ? (
              <>
                <XAxis
                  type="number"
                  tick={{ fontSize: 10 }}
                  label={yAxisLabel ? { value: yAxisLabel, position: "insideBottom", offset: -8, style: axisLabelStyle } : undefined}
                  height={yAxisLabel ? 44 : 30}
                />
                <YAxis dataKey="name" type="category" tick={{ fontSize: 10 }} width={90}
                  label={xAxisLabel ? { value: xAxisLabel, angle: -90, position: "insideLeft", dx: 10, style: axisLabelStyle } : undefined}
                />
              </>
            ) : (
              <>
                <XAxis
                  dataKey="name"
                  tick={{ fontSize: 10 }}
                  label={xAxisLabel ? { value: xAxisLabel, position: "insideBottom", offset: -16, style: axisLabelStyle } : undefined}
                  height={xAxisLabel ? 48 : 30}
                />
                <YAxis
                  tick={{ fontSize: 10 }}
                  width={yAxisLabel ? 70 : 50}
                  label={yAxisLabel ? { value: yAxisLabel, angle: -90, position: "insideLeft", dx: -10, style: axisLabelStyle } : undefined}
                />
              </>
            )}
            <Tooltip />
            {showLegend && yKeys.length > 1 && <Legend wrapperStyle={{ fontSize: 10 }} verticalAlign="top" align="right" />}
            {yKeys.map((key, i) => (
              <Bar key={key} dataKey={key} fill={colors[i % colors.length]} stackId={stacked ? "a" : undefined} radius={[2, 2, 0, 0]}>
                {showLabels && <LabelList dataKey={key} position="top" style={{ fontSize: 9 }} />}
              </Bar>
            ))}
          </BarChart>
        </ResponsiveContainer>
      );

    case "line":
    case "combo":
      return (
        <ResponsiveContainer width="100%" height="100%">
          <LineChart {...commonProps} margin={{ top: 8, right: 12, bottom: xAxisLabel ? 36 : 8, left: yAxisLabel ? 40 : 16 }}>
            {showGrid && <CartesianGrid strokeDasharray="3 3" stroke="#F0F0F0" />}
            <XAxis
              dataKey="name"
              tick={{ fontSize: 10 }}
              label={xAxisLabel ? { value: xAxisLabel, position: "insideBottom", offset: -16, style: axisLabelStyle } : undefined}
              height={xAxisLabel ? 48 : 30}
            />
            <YAxis
              tick={{ fontSize: 10 }}
              width={yAxisLabel ? 70 : 50}
              label={yAxisLabel ? { value: yAxisLabel, angle: -90, position: "insideLeft", dx: -10, style: axisLabelStyle } : undefined}
            />
            <Tooltip />
            {showLegend && yKeys.length > 1 && <Legend wrapperStyle={{ fontSize: 10 }} verticalAlign="top" align="right" />}
            {yKeys.map((key, i) => (
              <Line key={key} type="monotone" dataKey={key} stroke={colors[i % colors.length]} strokeWidth={chart.strokeWidth ?? 2} dot={false}>
                {showLabels && <LabelList dataKey={key} position="top" style={{ fontSize: 9 }} />}
              </Line>
            ))}
          </LineChart>
        </ResponsiveContainer>
      );

    case "area":
    case "area_stacked":
      return (
        <ResponsiveContainer width="100%" height="100%">
          <AreaChart {...commonProps} margin={{ top: 8, right: 12, bottom: xAxisLabel ? 36 : 8, left: yAxisLabel ? 40 : 16 }}>
            {showGrid && <CartesianGrid strokeDasharray="3 3" stroke="#F0F0F0" />}
            <XAxis
              dataKey="name"
              tick={{ fontSize: 10 }}
              label={xAxisLabel ? { value: xAxisLabel, position: "insideBottom", offset: -16, style: axisLabelStyle } : undefined}
              height={xAxisLabel ? 48 : 30}
            />
            <YAxis
              tick={{ fontSize: 10 }}
              width={yAxisLabel ? 70 : 50}
              label={yAxisLabel ? { value: yAxisLabel, angle: -90, position: "insideLeft", dx: -10, style: axisLabelStyle } : undefined}
            />
            <Tooltip />
            {showLegend && yKeys.length > 1 && <Legend wrapperStyle={{ fontSize: 10 }} verticalAlign="top" align="right" />}
            {yKeys.map((key, i) => (
              <Area
                key={key}
                type="monotone"
                dataKey={key}
                stroke={colors[i % colors.length]}
                fill={colors[i % colors.length]}
                fillOpacity={chart.fillOpacity ?? 0.2}
                stackId={chart.type === "area_stacked" ? "s" : undefined}
              />
            ))}
          </AreaChart>
        </ResponsiveContainer>
      );

    case "pie":
    case "donut":
      return (
        <ResponsiveContainer width="100%" height="100%">
          <PieChart>
            <Pie
              data={data}
              dataKey={yKeys[0] || "value"}
              nameKey="name"
              cx="50%" cy="50%"
              innerRadius={chart.type === "donut" ? "40%" : 0}
              outerRadius="70%"
              paddingAngle={2}
              label={showLabels ? ({ name, percent }: { name?: string; percent?: number }) => `${name ?? ""} (${((percent ?? 0) * 100).toFixed(0)}%)` : false}
              labelLine={showLabels}
            >
              {data.map((_, index) => (
                <Cell key={`cell-${index}`} fill={colors[index % colors.length]} />
              ))}
            </Pie>
            <Tooltip />
            {showLegend && yKeys.length > 1 && <Legend wrapperStyle={{ fontSize: 10 }} verticalAlign="top" align="right" />}
          </PieChart>
        </ResponsiveContainer>
      );

    case "radar":
      return (
        <ResponsiveContainer width="100%" height="100%">
          <RadarChart data={data}>
            <PolarGrid />
            <PolarAngleAxis dataKey="name" tick={{ fontSize: 10 }} />
            <Tooltip />
            {showLegend && yKeys.length > 1 && <Legend wrapperStyle={{ fontSize: 10 }} verticalAlign="top" align="right" />}
            {yKeys.map((key, i) => (
              <Radar key={key} name={key} dataKey={key} stroke={colors[i % colors.length]} fill={colors[i % colors.length]} fillOpacity={0.4} />
            ))}
          </RadarChart>
        </ResponsiveContainer>
      );

    case "radial_bar":
      return (
        <ResponsiveContainer width="100%" height="100%">
          <RadialBarChart data={data.slice(0, 10)} innerRadius="10%" outerRadius="90%">
            <RadialBar dataKey={yKeys[0] || "value"} label={showLabels ? { position: "insideStart", fill: "#fff", fontSize: 9 } : false}>
              {data.slice(0, 10).map((_, index) => (
                <Cell key={`cell-${index}`} fill={colors[index % colors.length]} />
              ))}
            </RadialBar>
            <Tooltip />
            {showLegend && yKeys.length > 1 && <Legend wrapperStyle={{ fontSize: 10 }} verticalAlign="top" align="right" />}
          </RadialBarChart>
        </ResponsiveContainer>
      );

    case "scatter":
      return (
        <ResponsiveContainer width="100%" height="100%">
          <ScatterChart {...commonProps}>
            {showGrid && <CartesianGrid strokeDasharray="3 3" stroke="#F0F0F0" />}
            <XAxis dataKey={chart.xColumn} name={chart.xColumn} tick={{ fontSize: 10 }} />
            <YAxis dataKey={yKeys[0]} name={yKeys[0]} tick={{ fontSize: 10 }} />
            <Tooltip cursor={{ strokeDasharray: "3 3" }} />
            <Scatter name={yKeys[0]} data={data} fill={colors[0]} />
          </ScatterChart>
        </ResponsiveContainer>
      );

    case "funnel":
      return (
        <ResponsiveContainer width="100%" height="100%">
          <FunnelChart>
            <Tooltip />
            <Funnel dataKey={yKeys[0] || "value"} data={data} isAnimationActive>
              {data.map((_, index) => (
                <Cell key={`cell-${index}`} fill={colors[index % colors.length]} />
              ))}
              {showLabels && <LabelList position="right" fill="#000" stroke="none" dataKey="name" style={{ fontSize: 10 }} />}
            </Funnel>
          </FunnelChart>
        </ResponsiveContainer>
      );

    case "treemap":
      return (
        <ResponsiveContainer width="100%" height="100%">
          <Treemap
            data={data.map((d, i) => ({ name: d.name, size: Number(d[yKeys[0]]) || 0, fill: colors[i % colors.length] }))}
            dataKey="size"
            nameKey="name"
            aspectRatio={4 / 3}
          />
        </ResponsiveContainer>
      );

    case "histogram":
      return (
        <ResponsiveContainer width="100%" height="100%">
          <BarChart {...commonProps}>
            {showGrid && <CartesianGrid strokeDasharray="3 3" stroke="#F0F0F0" />}
            <XAxis dataKey="name" tick={{ fontSize: 10 }} />
            <YAxis tick={{ fontSize: 10 }} />
            <Tooltip />
            <Bar dataKey={yKeys[0] || "value"} fill={colors[0]} />
          </BarChart>
        </ResponsiveContainer>
      );

    case "table":
      return (
        <div className="h-full overflow-auto">
          <table className="w-full text-xs">
            <thead className="sticky top-0 bg-[#F9F9F9]">
              <tr>
                <th className="border-b border-[#E5E5E5] px-2 py-1 text-left font-medium">{chart.xColumn}</th>
                {yKeys.map((k) => (
                  <th key={k} className="border-b border-[#E5E5E5] px-2 py-1 text-right font-medium">{k}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {data.map((row, i) => (
                <tr key={i} className="border-b border-[#F5F5F5]">
                  <td className="px-2 py-1">{String(row.name)}</td>
                  {yKeys.map((k) => (
                    <td key={k} className="px-2 py-1 text-right tabular-nums">{String(row[k] ?? "")}</td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      );

    case "gauge": {
      const val = data[0] ? Number(data[0][yKeys[0]]) : 0;
      const max = Math.max(...data.map((d) => Number(d[yKeys[0]]) || 0));
      const pct = max > 0 ? val / max : 0;
      return (
        <div className="flex h-full flex-col items-center justify-center gap-2">
          <div className="relative flex h-28 w-28 items-center justify-center rounded-full border-8" style={{ borderColor: colors[0] }}>
            <span className="text-2xl font-bold">{val.toFixed(1)}</span>
          </div>
          <div className="h-2 w-48 overflow-hidden rounded-full bg-[#F0F0F0]">
            <div className="h-full rounded-full" style={{ width: `${pct * 100}%`, backgroundColor: colors[0] }} />
          </div>
          <span className="text-xs text-[#6B6B6B]">{yKeys[0]}: {val.toFixed(2)} / {max.toFixed(2)}</span>
        </div>
      );
    }

    default:
      return <EmptyState />;
  }
}

function EmptyState() {
  return (
    <div className="flex h-full items-center justify-center text-xs text-[#A1A1A1]">
      No data to display
    </div>
  );
}
