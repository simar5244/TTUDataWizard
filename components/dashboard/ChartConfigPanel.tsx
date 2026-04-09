"use client";

import { Label } from "@/components/ui/label";
import { Input } from "@/components/ui/input";
import { Switch } from "@/components/ui/switch";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import type { ChartConfig, ChartType } from "@/components/dashboard/DashboardBuilder";
import type { ExcelSheet } from "@/lib/excel";

const CHART_TYPES: { value: ChartType; label: string }[] = [
  { value: "bar", label: "Bar" },
  { value: "bar_horizontal", label: "Bar (Horizontal)" },
  { value: "line", label: "Line" },
  { value: "area", label: "Area" },
  { value: "area_stacked", label: "Area (Stacked)" },
  { value: "pie", label: "Pie" },
  { value: "donut", label: "Donut" },
  { value: "radar", label: "Radar" },
  { value: "radial_bar", label: "Radial Bar" },
  { value: "scatter", label: "Scatter" },
  { value: "bubble", label: "Bubble" },
  { value: "histogram", label: "Histogram" },
  { value: "treemap", label: "Treemap" },
  { value: "funnel", label: "Funnel" },
  { value: "waterfall", label: "Waterfall" },
  { value: "gauge", label: "Gauge" },
  { value: "combo", label: "Bar + Line Combo" },
  { value: "table", label: "Data Table" },
];

const COLOR_SCHEMES = [
  { value: "default", label: "Default (B&W)" },
  { value: "blues", label: "Blues" },
  { value: "greens", label: "Greens" },
  { value: "reds", label: "Reds" },
  { value: "purples", label: "Purples" },
  { value: "warm", label: "Warm" },
  { value: "cool", label: "Cool" },
];

const AGGREGATIONS = [
  { value: "none", label: "None (raw)" },
  { value: "sum", label: "Sum" },
  { value: "avg", label: "Average" },
  { value: "count", label: "Count" },
  { value: "min", label: "Min" },
  { value: "max", label: "Max" },
];

interface ChartConfigPanelProps {
  chart: ChartConfig;
  sheet: ExcelSheet;
  onChange: (updates: Partial<ChartConfig>) => void;
}

export function ChartConfigPanel({ chart, sheet, onChange }: ChartConfigPanelProps) {
  const cols = sheet.columns;

  return (
    <div className="flex flex-1 flex-col gap-4 overflow-auto p-4 text-sm">
      {/* Title */}
      <div className="space-y-1.5">
        <Label className="text-xs">Chart Title</Label>
        <Input
          className="h-8 text-xs"
          value={chart.title}
          onChange={(e) => onChange({ title: e.target.value })}
        />
      </div>

      {/* Chart type */}
      <div className="space-y-1.5">
        <Label className="text-xs">Chart Type</Label>
        <Select value={chart.type} onValueChange={(v) => onChange({ type: v as ChartType })}>
          <SelectTrigger className="h-8 text-xs">
            <SelectValue />
          </SelectTrigger>
          <SelectContent>
            {CHART_TYPES.map((t) => (
              <SelectItem key={t.value} value={t.value} className="text-xs">{t.label}</SelectItem>
            ))}
          </SelectContent>
        </Select>
      </div>

      {/* X Axis */}
      <div className="space-y-1.5">
        <Label className="text-xs">X Axis / Labels</Label>
        <Select value={chart.xColumn} onValueChange={(v) => onChange({ xColumn: v })}>
          <SelectTrigger className="h-8 text-xs">
            <SelectValue placeholder="Select column…" />
          </SelectTrigger>
          <SelectContent>
            {cols.map((c) => (
              <SelectItem key={c.key} value={c.header} className="text-xs">{c.header}</SelectItem>
            ))}
          </SelectContent>
        </Select>
      </div>

      {/* Y Axis */}
      <div className="space-y-1.5">
        <Label className="text-xs">Y Axis / Values (multi-select)</Label>
        <div className="max-h-40 space-y-1 overflow-auto rounded border border-[#E5E5E5] p-2">
          {cols.map((c) => {
            const selected = chart.yColumns.includes(c.header);
            return (
              <label key={c.key} className="flex cursor-pointer items-center gap-2">
                <input
                  type="checkbox"
                  checked={selected}
                  onChange={(e) => {
                    if (e.target.checked) {
                      onChange({ yColumns: [...chart.yColumns, c.header] });
                    } else {
                      onChange({ yColumns: chart.yColumns.filter((y) => y !== c.header) });
                    }
                  }}
                  className="h-3 w-3"
                />
                <span className="flex-1 text-xs">{c.header}</span>
                <span className={`text-[9px] px-1 rounded ${c.dataType === "number" ? "bg-blue-50 text-blue-500" : "bg-[#F5F5F5] text-[#A1A1A1]"}`}>
                  {c.dataType}
                </span>
              </label>
            );
          })}
        </div>
      </div>

      {/* Aggregation */}
      <div className="space-y-1.5">
        <Label className="text-xs">Aggregation</Label>
        <Select
          value={chart.aggregation}
          onValueChange={(v) => onChange({ aggregation: v as ChartConfig["aggregation"] })}
        >
          <SelectTrigger className="h-8 text-xs"><SelectValue /></SelectTrigger>
          <SelectContent>
            {AGGREGATIONS.map((a) => (
              <SelectItem key={a.value} value={a.value} className="text-xs">{a.label}</SelectItem>
            ))}
          </SelectContent>
        </Select>
      </div>

      {/* Color scheme */}
      <div className="space-y-1.5">
        <Label className="text-xs">Color Scheme</Label>
        <Select value={chart.colorScheme} onValueChange={(v) => onChange({ colorScheme: v })}>
          <SelectTrigger className="h-8 text-xs"><SelectValue /></SelectTrigger>
          <SelectContent>
            {COLOR_SCHEMES.map((s) => (
              <SelectItem key={s.value} value={s.value} className="text-xs">{s.label}</SelectItem>
            ))}
          </SelectContent>
        </Select>
      </div>

      {/* Row limit */}
      <div className="space-y-1.5">
        <Label className="text-xs">Max rows (0 = all)</Label>
        <Input
          className="h-8 text-xs"
          type="number"
          min={0}
          value={chart.limit ?? 0}
          onChange={(e) => onChange({ limit: parseInt(e.target.value) || 0 })}
        />
      </div>

      {/* Toggles */}
      <div className="space-y-2 border-t border-[#F5F5F5] pt-3">
        {[
          { key: "showLegend", label: "Show legend" },
          { key: "showGrid", label: "Show grid lines" },
          { key: "showLabels", label: "Show data labels" },
          { key: "stacked", label: "Stacked" },
        ].map(({ key, label }) => (
          <div key={key} className="flex items-center justify-between">
            <span className="text-xs text-[#6B6B6B]">{label}</span>
            <Switch
              checked={!!chart[key as keyof ChartConfig]}
              onCheckedChange={(v) => onChange({ [key]: v })}
            />
          </div>
        ))}
      </div>
    </div>
  );
}
