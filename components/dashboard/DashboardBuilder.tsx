"use client";

import { useState, useRef, useCallback, useEffect } from "react";
import { useDropzone } from "react-dropzone";
import {
  Upload, Save, Trash2, Loader2, BarChart2, Settings2,
  Eye, EyeOff, BarChart, Type, Hash, Image as ImageIcon, Filter, ChevronDown,
  Share2, ExternalLink, RefreshCw, FileText, GripVertical, ChevronLeft, ChevronRight,
} from "lucide-react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { ChartRenderer } from "@/components/dashboard/ChartRenderer";
import { ChartConfigPanel } from "@/components/dashboard/ChartConfigPanel";
import { parseExcelFile, type ExcelSheet } from "@/lib/excel";
import { useToast } from "@/hooks/use-toast";
import "react-grid-layout/css/styles.css";
import "react-resizable/css/styles.css";

// eslint-disable-next-line @typescript-eslint/no-require-imports
const _RGL = require("react-grid-layout");
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const GridLayout = (_RGL.default || _RGL) as any;

interface GridItem { i: string; x: number; y: number; w: number; h: number; }

export interface ChartConfig {
  id: string;
  widgetType?: "chart";
  title: string;
  type: ChartType;
  xColumn: string;
  yColumns: string[];
  colorScheme: string;
  aggregation: "none" | "sum" | "avg" | "count" | "min" | "max";
  sortBy?: "asc" | "desc" | "none";
  limit?: number;
  showLegend: boolean;
  showGrid: boolean;
  showLabels: boolean;
  stacked?: boolean;
  fillOpacity?: number;
  strokeWidth?: number;
  innerRadius?: number;
}

export interface TextWidget {
  id: string;
  widgetType: "text";
  title: string;
  content: string;
  fontSize?: "xs" | "sm" | "base" | "lg" | "xl" | "2xl" | "3xl";
  align?: "left" | "center" | "right";
  bgColor?: string;
  textColor?: string;
  isBold?: boolean;
}

export interface KPIWidget {
  id: string;
  widgetType: "kpi";
  title: string;
  metric: string;
  aggregation: "sum" | "avg" | "count" | "min" | "max" | "last";
  prefix?: string;
  suffix?: string;
  format?: "number" | "currency" | "percent";
  colorScheme?: string;
  comparisonMetric?: string;
}

export interface ImageWidget {
  id: string;
  widgetType: "image";
  title: string;
  url: string;
  alt?: string;
  fit?: "cover" | "contain" | "fill";
}

export interface FilterWidget {
  id: string;
  widgetType: "filter";
  title: string;
  column: string;
  filterType: "select" | "multiselect" | "range";
}

export type DashboardWidget = ChartConfig | TextWidget | KPIWidget | ImageWidget | FilterWidget;

export type ChartType =
  | "bar" | "bar_horizontal"
  | "line" | "area" | "area_stacked"
  | "pie" | "donut"
  | "scatter" | "bubble"
  | "radar" | "radial_bar"
  | "treemap" | "funnel"
  | "heatmap" | "waterfall"
  | "histogram" | "box_plot"
  | "candlestick" | "gauge"
  | "combo" | "table";

const CHART_TYPES: { value: ChartType; label: string; group: string }[] = [
  { value: "bar", label: "Bar", group: "Bar" },
  { value: "bar_horizontal", label: "Bar (Horizontal)", group: "Bar" },
  { value: "line", label: "Line", group: "Line & Area" },
  { value: "area", label: "Area", group: "Line & Area" },
  { value: "area_stacked", label: "Area (Stacked)", group: "Line & Area" },
  { value: "pie", label: "Pie", group: "Circular" },
  { value: "donut", label: "Donut", group: "Circular" },
  { value: "radar", label: "Radar", group: "Circular" },
  { value: "radial_bar", label: "Radial Bar", group: "Circular" },
  { value: "scatter", label: "Scatter", group: "Statistical" },
  { value: "bubble", label: "Bubble", group: "Statistical" },
  { value: "histogram", label: "Histogram", group: "Statistical" },
  { value: "box_plot", label: "Box Plot", group: "Statistical" },
  { value: "treemap", label: "Treemap", group: "Distribution" },
  { value: "funnel", label: "Funnel", group: "Distribution" },
  { value: "heatmap", label: "Heatmap", group: "Distribution" },
  { value: "waterfall", label: "Waterfall", group: "Financial" },
  { value: "candlestick", label: "Candlestick", group: "Financial" },
  { value: "gauge", label: "Gauge", group: "KPI" },
  { value: "combo", label: "Bar + Line (Combo)", group: "KPI" },
  { value: "table", label: "Data Table", group: "Table" },
];

export const COLOR_SCHEMES = [
  { value: "default", label: "Default (B&W)" },
  { value: "blues", label: "Blues" },
  { value: "greens", label: "Greens" },
  { value: "reds", label: "Reds" },
  { value: "purples", label: "Purples" },
  { value: "warm", label: "Warm" },
  { value: "cool", label: "Cool" },
  { value: "neon", label: "Neon" },
];

function makeDefaultChart(id: string): ChartConfig {
  return {
    id,
    widgetType: "chart",
    title: "New Chart",
    type: "bar",
    xColumn: "",
    yColumns: [],
    colorScheme: "default",
    aggregation: "none",
    showLegend: true,
    showGrid: true,
    showLabels: false,
    stacked: false,
    fillOpacity: 0.7,
    strokeWidth: 2,
    innerRadius: 0,
  };
}

function makeTextWidget(id: string): TextWidget {
  return { id, widgetType: "text", title: "Text Block", content: "Your text here", fontSize: "base", align: "left", bgColor: "#ffffff", textColor: "#0A0A0A" };
}
function makeKPIWidget(id: string): KPIWidget {
  return { id, widgetType: "kpi", title: "KPI", metric: "", aggregation: "sum", prefix: "", suffix: "", format: "number", colorScheme: "default" };
}
function makeImageWidget(id: string): ImageWidget {
  return { id, widgetType: "image", title: "Image", url: "", alt: "", fit: "contain" };
}
function makeFilterWidget(id: string): FilterWidget {
  return { id, widgetType: "filter", title: "Filter", column: "", filterType: "multiselect" };
}

function computeKPI(widget: KPIWidget, sheet: ExcelSheet): number | null {
  const col = sheet.columns.find((c) => c.header === widget.metric);
  if (!col) return null;
  const vals = sheet.data.map((r) => parseFloat(String(r[widget.metric] ?? 0))).filter((v) => !isNaN(v));
  if (vals.length === 0) return null;
  switch (widget.aggregation) {
    case "sum": return vals.reduce((a, b) => a + b, 0);
    case "avg": return vals.reduce((a, b) => a + b, 0) / vals.length;
    case "count": return vals.length;
    case "min": return Math.min(...vals);
    case "max": return Math.max(...vals);
    case "last": return vals[vals.length - 1];
    default: return null;
  }
}

function formatKPI(val: number, widget: KPIWidget): string {
  let formatted: string;
  if (widget.format === "currency") {
    formatted = new Intl.NumberFormat("en-US", { notation: "compact", maximumFractionDigits: 1 }).format(val);
  } else if (widget.format === "percent") {
    formatted = val.toFixed(1) + "%";
  } else {
    formatted = new Intl.NumberFormat("en-US", { notation: "compact", maximumFractionDigits: 1 }).format(val);
  }
  return `${widget.prefix || ""}${formatted}${widget.suffix || ""}`;
}

interface DashboardBuilderProps {
  initialDashboard?: {
    id: string;
    name: string;
    slug?: string;
    excelData: unknown;
    charts: unknown[];
    layout: unknown[];
  };
  onSaved: (id: string) => void;
  initialViewMode?: boolean;
}

export function DashboardBuilder({ initialDashboard, onSaved, initialViewMode = false }: DashboardBuilderProps) {
  const { toast } = useToast();
  const [name, setName] = useState(initialDashboard?.name ?? "");
  const [sheet, setSheet] = useState<ExcelSheet | null>(
    initialDashboard?.excelData ? (initialDashboard.excelData as ExcelSheet) : null
  );
  const [widgets, setWidgets] = useState<DashboardWidget[]>(
    (initialDashboard?.charts as DashboardWidget[]) ?? []
  );
  const [layout, setLayout] = useState<GridItem[]>(
    (initialDashboard?.layout as GridItem[]) ?? []
  );
  const [selectedWidget, setSelectedWidget] = useState<string | null>(null);
  const [saving, setSaving] = useState(false);
  const [viewMode, setViewMode] = useState(initialViewMode);
  const [addDropdown, setAddDropdown] = useState(false);
  const [containerWidth, setContainerWidth] = useState(1200);
  const [activeFilters, setActiveFilters] = useState<Record<string, string[]>>({});
  const containerRef = useRef<HTMLDivElement>(null);
  const refreshInputRef = useRef<HTMLInputElement>(null);
  const [refreshing, setRefreshing] = useState(false);
  const [leftW, setLeftW] = useState(208);
  const [rightW, setRightW] = useState(256);
  const [leftMin, setLeftMin] = useState(false);
  const [rightMin, setRightMin] = useState(false);
  const leftPanelRef = useRef<HTMLDivElement>(null);
  const rightPanelRef = useRef<HTMLDivElement>(null);
  const isDraggingLeft = useRef(false);
  const isDraggingRight = useRef(false);

  useEffect(() => {
    function updateWidth() {
      if (containerRef.current) {
        setContainerWidth(containerRef.current.clientWidth);
      }
    }
    updateWidth();
    window.addEventListener("resize", updateWidth);
    return () => window.removeEventListener("resize", updateWidth);
  }, []);

  useEffect(() => {
    function onMove(e: MouseEvent) {
      if (isDraggingLeft.current && leftPanelRef.current) {
        const rect = leftPanelRef.current.parentElement?.getBoundingClientRect();
        if (rect) setLeftW(Math.max(140, Math.min(400, e.clientX - rect.left)));
      }
      if (isDraggingRight.current && rightPanelRef.current) {
        const rect = rightPanelRef.current.parentElement?.getBoundingClientRect();
        if (rect) setRightW(Math.max(180, Math.min(440, rect.right - e.clientX)));
      }
    }
    function onUp() {
      isDraggingLeft.current = false;
      isDraggingRight.current = false;
    }
    window.addEventListener("mousemove", onMove);
    window.addEventListener("mouseup", onUp);
    return () => {
      window.removeEventListener("mousemove", onMove);
      window.removeEventListener("mouseup", onUp);
    };
  }, []);

  const onDrop = useCallback(async (files: File[]) => {
    const file = files[0];
    if (!file) return;
    const buf = await file.arrayBuffer();
    const parsed = parseExcelFile(buf, file.name, file.size);
    if (!parsed.sheets[0]) { toast({ title: "No data found", variant: "destructive" }); return; }
    setSheet(parsed.sheets[0]);
    toast({ title: `Loaded: ${file.name}`, description: `${parsed.sheets[0].rowCount} rows, ${parsed.sheets[0].columns.length} columns` });
  }, [toast]);

  const { getRootProps, getInputProps, isDragActive } = useDropzone({
    onDrop,
    accept: {
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"],
      "application/vnd.ms-excel": [".xls"],
      "text/csv": [".csv"],
    },
    noClick: !!sheet,
  });

  function nextPosition(prev: GridItem[]): { x: number; y: number } {
    if (prev.length === 0) return { x: 0, y: 0 };
    const maxY = prev.reduce((m, item) => Math.max(m, item.y + item.h), 0);
    return { x: 0, y: maxY };
  }

  function addWidget(type: "chart" | "text" | "kpi" | "image" | "filter") {
    const id = `widget_${Date.now()}`;
    let widget: DashboardWidget;
    let size = { w: 6, h: 4 };
    switch (type) {
      case "chart": widget = makeDefaultChart(id); size = { w: 12, h: 6 }; break;
      case "text": widget = makeTextWidget(id); size = { w: 12, h: 2 }; break;
      case "kpi": widget = makeKPIWidget(id); size = { w: 4, h: 2 }; break;
      case "image": widget = makeImageWidget(id); size = { w: 6, h: 4 }; break;
      case "filter": widget = makeFilterWidget(id); size = { w: 6, h: 2 }; break;
    }
    setWidgets((prev) => [...prev, widget]);
    setLayout((prev) => {
      const pos = nextPosition(prev);
      return [...prev, { i: id, x: pos.x, y: pos.y, w: size.w, h: size.h }];
    });
    setSelectedWidget(id);
    setAddDropdown(false);
  }

  function removeWidget(id: string) {
    setWidgets((prev) => prev.filter((w) => w.id !== id));
    setLayout((prev) => prev.filter((l) => l.i !== id));
    if (selectedWidget === id) setSelectedWidget(null);
  }

  function updateWidget(id: string, updates: Partial<DashboardWidget>) {
    setWidgets((prev) => prev.map((w) => (w.id === id ? { ...w, ...updates } as DashboardWidget : w)));
  }

  function getFilteredSheet(): ExcelSheet | null {
    if (!sheet) return null;
    const filterCols = Object.keys(activeFilters).filter((k) => activeFilters[k].length > 0);
    if (filterCols.length === 0) return sheet;
    const filteredData = sheet.data.filter((row) =>
      filterCols.every((col) => activeFilters[col].includes(String(row[col] ?? "")))
    );
    return { ...sheet, data: filteredData, rowCount: filteredData.length };
  }

  async function handleSave() {
    if (!name.trim()) { toast({ title: "Enter a dashboard name", variant: "destructive" }); return; }
    setSaving(true);
    try {
      const url = initialDashboard ? `/api/dashboards/${initialDashboard.id}` : "/api/dashboards";
      const method = initialDashboard ? "PUT" : "POST";
      const res = await fetch(url, {
        method,
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ name, excelData: sheet, charts: widgets, layout }),
      });
      if (!res.ok) throw new Error((await res.json()).error);
      const saved = await res.json();
      toast({ title: "Dashboard saved!" });
      onSaved(saved.id);
    } catch (e) {
      toast({ title: "Save failed", description: (e as Error).message, variant: "destructive" });
    } finally {
      setSaving(false);
    }
  }

  function copyShareUrl() {
    if (!initialDashboard?.slug) return;
    const url = `${window.location.origin}/share/${initialDashboard.slug}?type=dashboard`;
    navigator.clipboard.writeText(url);
    toast({ title: "View link copied!", description: "Anyone with the link can view this dashboard." });
  }

  function exportPDF() {
    setSelectedWidget(null);
    setTimeout(() => window.print(), 120);
  }

  async function handleRefreshFile(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    if (!file) return;
    setRefreshing(true);
    try {
      const buf = await file.arrayBuffer();
      const parsed = parseExcelFile(buf, file.name, file.size);
      const newSheet = parsed.sheets[0];
      if (!newSheet) { toast({ title: "No data found in file", variant: "destructive" }); return; }
      setSheet(newSheet);
      if (initialDashboard?.id) {
        await fetch(`/api/dashboards/${initialDashboard.id}`, {
          method: "PUT",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ excelData: newSheet }),
        });
      }
      toast({ title: "Data refreshed!", description: `${newSheet.rowCount.toLocaleString()} rows · ${newSheet.columns.length} cols` });
    } catch (err) {
      toast({ title: "Refresh failed", description: (err as Error).message, variant: "destructive" });
    } finally {
      setRefreshing(false);
      if (refreshInputRef.current) refreshInputRef.current.value = "";
    }
  }

  const activeWidget = widgets.find((w) => w.id === selectedWidget);
  const filteredSheet = getFilteredSheet();
  const canEdit = !viewMode;

  return (
    <div className="flex h-[calc(100vh-64px)] flex-col">
      {/* ─── Toolbar ──────────────────────────────────────────── */}
      <div className="no-print flex shrink-0 items-center gap-1.5 border-b border-[#E5E5E5] bg-white px-4 py-2">
        <div className="flex items-center gap-1 text-xs text-[#A1A1A1]">
          <span>Dashboards</span>
          <span>/</span>
          <span className="font-medium text-[#0A0A0A]">{name || "Untitled"}</span>
        </div>
        <div className="mx-2 h-4 w-px shrink-0 bg-[#E5E5E5]" />

        {canEdit && (
          <div className="relative">
            <Button size="sm" variant="ghost" className="h-7 gap-1 text-xs" onClick={() => setAddDropdown((v) => !v)} disabled={!sheet}>
              <span className="text-sm font-bold leading-none">+</span>Insert
              <ChevronDown className="h-3 w-3" />
            </Button>
            {addDropdown && (
              <div className="absolute left-0 top-8 z-50 min-w-[160px] rounded-lg border border-[#E5E5E5] bg-white py-1 shadow-lg">
                {[
                  { type: "chart" as const, icon: BarChart, label: "Chart" },
                  { type: "text" as const, icon: Type, label: "Text / Title" },
                  { type: "kpi" as const, icon: Hash, label: "KPI Card" },
                  { type: "image" as const, icon: ImageIcon, label: "Image / Logo" },
                  { type: "filter" as const, icon: Filter, label: "Filter" },
                ].map(({ type, icon: Icon, label }) => (
                  <button
                    key={type}
                    className="flex w-full items-center gap-2 px-3 py-2 text-xs hover:bg-[#F5F5F5]"
                    onClick={() => addWidget(type)}
                  >
                    <Icon className="h-3.5 w-3.5 text-[#6B6B6B]" />
                    {label}
                  </button>
                ))}
              </div>
            )}
          </div>
        )}

        {canEdit && sheet && (
          <Button size="sm" variant="ghost" className="h-7 text-xs" onClick={() => refreshInputRef.current?.click()} disabled={refreshing}>
            {refreshing ? <Loader2 className="mr-1 h-3 w-3 animate-spin" /> : <RefreshCw className="mr-1 h-3 w-3" />}
            Refresh
          </Button>
        )}

        <div className="flex-1" />

        <Button size="sm" variant="ghost" className="h-7 text-xs" onClick={exportPDF}>
          <FileText className="mr-1 h-3 w-3" />
          Export PDF
        </Button>

        {initialDashboard?.slug && (
          <Button size="sm" variant="ghost" className="h-7 w-7 p-0" onClick={copyShareUrl} title="Copy view link">
            <Share2 className="h-3.5 w-3.5" />
          </Button>
        )}
        {initialDashboard?.slug && (
          <Button size="sm" variant="ghost" className="h-7 w-7 p-0" title="Open share page" onClick={() => window.open(`/share/${initialDashboard.slug}?type=dashboard`, "_blank")}>
            <ExternalLink className="h-3.5 w-3.5" />
          </Button>
        )}

        <div className="mx-1 h-4 w-px shrink-0 bg-[#E5E5E5]" />

        <Button size="sm" variant={viewMode ? "default" : "ghost"} className="h-7 text-xs" onClick={() => { setViewMode((v) => !v); setSelectedWidget(null); }}>
          {viewMode ? <><EyeOff className="mr-1 h-3 w-3" />Edit</> : <><Eye className="mr-1 h-3 w-3" />Preview</>}
        </Button>

        {canEdit && (
          <Button size="sm" className="h-7 text-xs" onClick={handleSave} disabled={saving}>
            {saving ? <Loader2 className="mr-1 h-3 w-3 animate-spin" /> : <Save className="mr-1 h-3 w-3" />}
            Save
          </Button>
        )}
      </div>

      <div className="flex flex-1 overflow-hidden">
        {/* Left — data source panel */}
        {canEdit && (
          leftMin ? (
            <div
              className="no-print flex w-9 shrink-0 cursor-pointer flex-col items-center border-r border-[#E5E5E5] bg-white py-3 hover:bg-[#F9F9F9] transition-colors"
              onClick={() => setLeftMin(false)}
              title="Expand data panel"
            >
              <ChevronRight className="h-4 w-4 text-[#A1A1A1]" />
              <span className="mt-2 text-[8px] uppercase tracking-widest text-[#C0C0C0]" style={{ writingMode: "vertical-rl", transform: "rotate(180deg)" }}>Data</span>
            </div>
          ) : (
            <>
              <div
                ref={leftPanelRef}
                className="no-print flex shrink-0 flex-col border-r border-[#E5E5E5] bg-white"
                style={{ width: leftW }}
              >
                <div className="flex items-center justify-between border-b border-[#E5E5E5] px-4 py-2.5">
                  <p className="text-xs font-semibold uppercase tracking-wider text-[#6B6B6B]">Data Source</p>
                  <button onClick={() => setLeftMin(true)} className="rounded p-0.5 hover:bg-[#F5F5F5]" title="Collapse">
                    <ChevronLeft className="h-3.5 w-3.5 text-[#A1A1A1]" />
                  </button>
                </div>
                <div className="flex flex-1 flex-col overflow-auto p-3">
                  <div
                    {...getRootProps()}
                    className={`mb-3 flex cursor-pointer flex-col items-center justify-center rounded-lg border-2 border-dashed p-4 text-center transition-colors ${
                      isDragActive ? "border-black bg-[#F5F5F5]" : "border-[#E5E5E5] hover:border-[#A1A1A1]"
                    }`}
                  >
                    <input {...getInputProps()} />
                    <Upload className="mb-1.5 h-5 w-5 text-[#A1A1A1]" />
                    <p className="text-xs font-medium">{sheet ? sheet.name : "Upload Excel"}</p>
                    {sheet && <p className="text-[10px] text-[#6B6B6B]">{sheet.rowCount} rows · click to replace</p>}
                  </div>
                  {sheet && (
                    <div className="space-y-1">
                      <p className="text-xs font-medium text-[#6B6B6B]">Columns ({sheet.columns.length})</p>
                      {sheet.columns.map((col) => (
                        <div key={col.key} className="flex items-center gap-1.5 rounded border border-[#F5F5F5] bg-[#F9F9F9] px-2 py-1">
                          <span className="truncate text-[11px] font-medium">{col.header}</span>
                          <span className="ml-auto text-[9px] text-[#A1A1A1]">{col.dataType}</span>
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              </div>
              <div
                className="no-print w-1 cursor-col-resize bg-transparent hover:bg-blue-300 active:bg-blue-400 transition-colors"
                onMouseDown={() => { isDraggingLeft.current = true; }}
              />
            </>
          )
        )}

        {/* Main canvas — Word-style document */}
        <div
          className="doc-canvas flex-1 overflow-auto"
          style={{ background: "#525659", padding: "32px 24px" }}
          onClick={() => { if (addDropdown) setAddDropdown(false); }}
        >
          <div className="mx-auto w-full" style={{ maxWidth: "960px" }}>
            <div
              id="doc-page"
              className="bg-white"
              style={{
                boxShadow: "0 4px 32px rgba(0,0,0,0.35)",
                minHeight: "1056px",
                borderRadius: "1px",
                padding: "56px 64px 80px",
              }}
            >
              {/* Document title */}
              <div className="mb-8 border-b border-[#EBEBEB] pb-5">
                {canEdit ? (
                  <Input
                    className="h-auto border-0 p-0 text-2xl font-bold shadow-none focus-visible:ring-0 tracking-tight"
                    placeholder="Untitled dashboard"
                    value={name}
                    onChange={(e) => setName(e.target.value)}
                  />
                ) : (
                  <h1 className="text-2xl font-bold tracking-tight">{name || "Untitled"}</h1>
                )}
                {sheet && (
                  <p className="mt-1.5 text-xs text-[#A1A1A1]">
                    {sheet.rowCount.toLocaleString()} rows · {sheet.columns.length} columns · {sheet.name}
                  </p>
                )}
              </div>

              {/* Grid area */}
              <div ref={containerRef}>
                {widgets.length === 0 ? (
                  <div className="flex flex-col items-center justify-center py-24">
                    <BarChart2 className="mb-3 h-10 w-10 text-[#D1D1D1]" />
                    <p className="text-sm text-[#A1A1A1]">
                      {sheet
                        ? 'Use "+  Insert" in the toolbar to add charts, KPIs, or text'
                        : "Upload an Excel file from the left panel to begin"}
                    </p>
                  </div>
                ) : (
                  <GridLayout
                    className="layout"
                    layout={layout}
                    cols={12}
                    rowHeight={80}
                    width={containerWidth > 0 ? containerWidth : 800}
                    onLayoutChange={(l: GridItem[]) => setLayout(l)}
                    draggableHandle=".drag-handle"
                    isDraggable={canEdit}
                    isResizable={canEdit}
                    margin={[0, 16]}
                  >
                    {widgets.map((widget) => (
                      <div
                        key={widget.id}
                        className="group relative flex flex-col"
                        onClick={() => canEdit && setSelectedWidget(widget.id)}
                      >
                        {/* Drag strip */}
                        {canEdit && (
                          <div className="drag-handle flex h-4 shrink-0 cursor-grab items-center justify-center opacity-0 transition-opacity group-hover:opacity-100">
                            <GripVertical className="h-3 w-3 text-[#BBBBBB]" />
                          </div>
                        )}

                        {/* Content */}
                        <div className={canEdit ? "h-[calc(100%-16px)] overflow-hidden" : "h-full overflow-hidden"}>
                          <WidgetRenderer
                            widget={widget}
                            sheet={filteredSheet}
                            activeFilters={activeFilters}
                            onFilterChange={(col, vals) => setActiveFilters((prev) => ({ ...prev, [col]: vals }))}
                            canEdit={canEdit}
                          />
                        </div>

                        {/* Floating controls */}
                        {canEdit && (
                          <div
                            className={`absolute right-1 top-5 z-20 flex items-center gap-0.5 rounded-md bg-white px-0.5 py-0.5 shadow-md transition-all ${
                              selectedWidget === widget.id
                                ? "opacity-100"
                                : "pointer-events-none opacity-0 group-hover:pointer-events-auto group-hover:opacity-100"
                            }`}
                            onClick={(e) => e.stopPropagation()}
                          >
                            <button className="rounded p-1 hover:bg-[#F5F5F5]" onClick={() => setSelectedWidget(widget.id)}>
                              <Settings2 className="h-3 w-3 text-[#6B6B6B]" />
                            </button>
                            <button className="rounded p-1 hover:bg-red-50" onClick={() => removeWidget(widget.id)}>
                              <Trash2 className="h-3 w-3 text-red-400" />
                            </button>
                          </div>
                        )}

                        {/* Selection ring */}
                        {canEdit && (
                          <div
                            className={`pointer-events-none absolute inset-0 rounded-sm transition-all ${
                              selectedWidget === widget.id
                                ? "ring-2 ring-blue-500 ring-offset-1"
                                : "ring-0 group-hover:ring-1 group-hover:ring-[#C0C0C0]"
                            }`}
                          />
                        )}
                      </div>
                    ))}
                  </GridLayout>
                )}
              </div>
            </div>
          </div>
        </div>

        {/* Right — config panel */}
        {canEdit && (
          rightMin ? (
            <div
              className="no-print flex w-9 shrink-0 cursor-pointer flex-col items-center border-l border-[#E5E5E5] bg-white py-3 hover:bg-[#F9F9F9] transition-colors"
              onClick={() => setRightMin(false)}
              title="Expand config panel"
            >
              <ChevronLeft className="h-4 w-4 text-[#A1A1A1]" />
              <span className="mt-2 text-[8px] uppercase tracking-widest text-[#C0C0C0]" style={{ writingMode: "vertical-rl", transform: "rotate(180deg)" }}>Config</span>
            </div>
          ) : (
            <>
              <div
                className="no-print w-1 cursor-col-resize bg-transparent hover:bg-blue-300 active:bg-blue-400 transition-colors"
                onMouseDown={() => { isDraggingRight.current = true; }}
              />
              <div
                ref={rightPanelRef}
                className="no-print flex shrink-0 flex-col border-l border-[#E5E5E5] bg-white"
                style={{ width: rightW }}
              >
                <div className="flex items-center justify-between border-b border-[#E5E5E5] px-4 py-2.5">
                  <p className="text-xs font-semibold uppercase tracking-wider text-[#6B6B6B]">
                    {activeWidget ? `${(activeWidget.widgetType ?? "chart").toUpperCase()} CONFIG` : "Widget Config"}
                  </p>
                  <button onClick={() => setRightMin(true)} className="rounded p-0.5 hover:bg-[#F5F5F5]" title="Collapse">
                    <ChevronRight className="h-3.5 w-3.5 text-[#A1A1A1]" />
                  </button>
                </div>
              {activeWidget && sheet ? (
                <WidgetConfigPanel
                  widget={activeWidget}
                  sheet={sheet}
                  onChange={(updates) => updateWidget(activeWidget.id, updates)}
                />
              ) : (
                <div className="flex flex-1 items-center justify-center p-4">
                  <p className="text-center text-xs text-[#A1A1A1]">
                    {widgets.length === 0 ? "Add a widget to configure it" : "Click a widget to configure"}
                  </p>
                </div>
              )}
              </div>
            </>
          )
        )}
      </div>

      {/* Hidden file input for data refresh */}
      <input
        ref={refreshInputRef}
        type="file"
        accept=".xlsx,.xls,.csv"
        className="hidden"
        onChange={handleRefreshFile}
      />
    </div>
  );
}

/* ─── Widget Renderer ─────────────────────────────────────────── */

interface WidgetRendererProps {
  widget: DashboardWidget;
  sheet: ExcelSheet | null;
  activeFilters: Record<string, string[]>;
  onFilterChange: (col: string, vals: string[]) => void;
  canEdit: boolean;
}

function WidgetRenderer({ widget, sheet, activeFilters, onFilterChange, canEdit }: WidgetRendererProps) {
  const type = widget.widgetType ?? "chart";

  if (type === "chart") {
    const chart = widget as ChartConfig;
    if (!sheet || !chart.xColumn || chart.yColumns.length === 0) {
      return <EmptyPlaceholder text={canEdit ? "Configure chart →" : "No data"} />;
    }
    return <ChartRenderer chart={chart} sheet={sheet} />;
  }

  if (type === "text") {
    const w = widget as TextWidget;
    return (
      <div
        className="flex h-full w-full items-start overflow-auto rounded p-3"
        style={{ backgroundColor: w.bgColor || "#ffffff" }}
      >
        <p
          className={`whitespace-pre-wrap text-${w.fontSize || "base"} text-${w.align || "left"} leading-relaxed ${w.isBold ? "font-bold" : ""}`}
          style={{ color: w.textColor || "#0A0A0A" }}
        >
          {w.content || (canEdit ? "Enter text in config →" : "")}
        </p>
      </div>
    );
  }

  if (type === "kpi") {
    const w = widget as KPIWidget;
    const value = sheet ? computeKPI(w, sheet) : null;
    return (
      <div className="flex h-full w-full flex-col items-center justify-center gap-1 rounded p-3">
        <p className="text-[10px] font-medium uppercase tracking-wider text-[#6B6B6B]">{w.title}</p>
        <p className="text-4xl font-bold tracking-tight">
          {value !== null ? formatKPI(value, w) : <span className="text-xl text-[#A1A1A1]">{canEdit ? "Select metric →" : "—"}</span>}
        </p>
        {w.metric && <p className="text-[10px] text-[#A1A1A1]">{w.aggregation} of {w.metric}</p>}
      </div>
    );
  }

  if (type === "image") {
    const w = widget as ImageWidget;
    if (!w.url) return <EmptyPlaceholder text={canEdit ? "Enter image URL in config →" : "No image"} />;
    return (
      <div className="flex h-full w-full items-center justify-center overflow-hidden rounded">
        {/* eslint-disable-next-line @next/next/no-img-element */}
        <img src={w.url} alt={w.alt || ""} className={`h-full w-full object-${w.fit || "contain"}`} />
      </div>
    );
  }

  if (type === "filter") {
    const w = widget as FilterWidget;
    if (!sheet || !w.column) return <EmptyPlaceholder text={canEdit ? "Select column in config →" : "No filter"} />;
    const col = sheet.columns.find((c) => c.header === w.column);
    if (!col) return <EmptyPlaceholder text="Column not found" />;
    const uniqueVals = Array.from(new Set(sheet.data.map((r) => String(r[w.column] ?? "")))).sort();
    const selected = activeFilters[w.column] || [];
    return (
      <div className="flex h-full w-full flex-col gap-1 overflow-auto p-2">
        <p className="text-[10px] font-medium text-[#6B6B6B]">Filter: {w.column}</p>
        <div className="flex flex-wrap gap-1">
          {uniqueVals.map((val) => (
            <button
              key={val}
              className={`rounded-full px-2 py-0.5 text-[10px] border transition-colors ${
                selected.length === 0 || selected.includes(val)
                  ? "border-black bg-black text-white"
                  : "border-[#E5E5E5] text-[#6B6B6B] hover:border-black"
              }`}
              onClick={() => {
                if (selected.length === 0) {
                  onFilterChange(w.column, uniqueVals.filter((v) => v !== val));
                } else if (selected.includes(val)) {
                  const next = selected.filter((v) => v !== val);
                  onFilterChange(w.column, next.length === uniqueVals.length ? [] : next);
                } else {
                  onFilterChange(w.column, [...selected, val]);
                }
              }}
            >
              {val}
            </button>
          ))}
        </div>
        {selected.length > 0 && (
          <button className="mt-1 text-[10px] text-[#6B6B6B] underline" onClick={() => onFilterChange(w.column, [])}>
            Clear filter
          </button>
        )}
      </div>
    );
  }

  return null;
}

function EmptyPlaceholder({ text }: { text: string }) {
  return (
    <div className="flex h-full w-full items-center justify-center">
      <p className="text-xs text-[#A1A1A1]">{text}</p>
    </div>
  );
}

/* ─── Widget Config Panel ─────────────────────────────────────── */

interface WidgetConfigPanelProps {
  widget: DashboardWidget;
  sheet: ExcelSheet;
  onChange: (updates: Partial<DashboardWidget>) => void;
}

function WidgetConfigPanel({ widget, sheet, onChange }: WidgetConfigPanelProps) {
  const type = widget.widgetType ?? "chart";

  if (type === "chart") {
    return (
      <ChartConfigPanel
        chart={widget as ChartConfig}
        sheet={sheet}
        onChange={onChange as (u: Partial<ChartConfig>) => void}
      />
    );
  }

  if (type === "text") {
    const w = widget as TextWidget;
    return (
      <div className="flex flex-1 flex-col gap-3 overflow-auto p-4 text-sm">
        <ConfigRow label="Title">
          <Input className="h-7 text-xs" value={w.title} onChange={(e) => onChange({ title: e.target.value })} />
        </ConfigRow>
        <ConfigRow label="Content">
          <textarea
            className="w-full rounded border border-[#E5E5E5] p-2 text-xs outline-none focus:border-black"
            rows={5}
            value={w.content}
            onChange={(e) => onChange({ content: e.target.value })}
          />
        </ConfigRow>
        <ConfigRow label="Font size">
          <select className="w-full rounded border border-[#E5E5E5] p-1 text-xs" value={w.fontSize || "base"} onChange={(e) => onChange({ fontSize: e.target.value as TextWidget["fontSize"] })}>
            {["xs","sm","base","lg","xl","2xl","3xl"].map((s) => <option key={s} value={s}>{s}</option>)}
          </select>
        </ConfigRow>
        <ConfigRow label="Align">
          <div className="flex gap-1">
            {(["left","center","right"] as const).map((a) => (
              <button key={a} onClick={() => onChange({ align: a })} className={`flex-1 rounded border py-1 text-[10px] ${w.align === a ? "border-black bg-black text-white" : "border-[#E5E5E5] hover:border-black"}`}>{a}</button>
            ))}
          </div>
        </ConfigRow>
        <ConfigRow label="Text color"><input type="color" value={w.textColor || "#0A0A0A"} onChange={(e) => onChange({ textColor: e.target.value })} className="h-7 w-full rounded border border-[#E5E5E5]" /></ConfigRow>
        <ConfigRow label="Background"><input type="color" value={w.bgColor || "#ffffff"} onChange={(e) => onChange({ bgColor: e.target.value })} className="h-7 w-full rounded border border-[#E5E5E5]" /></ConfigRow>
        <div className="flex items-center justify-between">
          <span className="text-xs text-[#6B6B6B]">Bold</span>
          <input type="checkbox" checked={!!w.isBold} onChange={(e) => onChange({ isBold: e.target.checked })} />
        </div>
      </div>
    );
  }

  if (type === "kpi") {
    const w = widget as KPIWidget;
    const numericCols = sheet.columns.filter((c) => c.dataType === "number");
    return (
      <div className="flex flex-1 flex-col gap-3 overflow-auto p-4 text-sm">
        <ConfigRow label="Label">
          <Input className="h-7 text-xs" value={w.title} onChange={(e) => onChange({ title: e.target.value })} />
        </ConfigRow>
        <ConfigRow label="Metric column">
          <select className="w-full rounded border border-[#E5E5E5] p-1 text-xs" value={w.metric} onChange={(e) => onChange({ metric: e.target.value })}>
            <option value="">Select…</option>
            {numericCols.map((c) => <option key={c.key} value={c.header}>{c.header}</option>)}
          </select>
        </ConfigRow>
        <ConfigRow label="Aggregation">
          <select className="w-full rounded border border-[#E5E5E5] p-1 text-xs" value={w.aggregation} onChange={(e) => onChange({ aggregation: e.target.value as KPIWidget["aggregation"] })}>
            {["sum","avg","count","min","max","last"].map((a) => <option key={a} value={a}>{a}</option>)}
          </select>
        </ConfigRow>
        <ConfigRow label="Format">
          <select className="w-full rounded border border-[#E5E5E5] p-1 text-xs" value={w.format || "number"} onChange={(e) => onChange({ format: e.target.value as KPIWidget["format"] })}>
            <option value="number">Number</option>
            <option value="currency">Currency</option>
            <option value="percent">Percent</option>
          </select>
        </ConfigRow>
        <ConfigRow label="Prefix"><Input className="h-7 text-xs" placeholder="$" value={w.prefix || ""} onChange={(e) => onChange({ prefix: e.target.value })} /></ConfigRow>
        <ConfigRow label="Suffix"><Input className="h-7 text-xs" placeholder="k" value={w.suffix || ""} onChange={(e) => onChange({ suffix: e.target.value })} /></ConfigRow>
      </div>
    );
  }

  if (type === "image") {
    const w = widget as ImageWidget;
    return (
      <div className="flex flex-1 flex-col gap-3 overflow-auto p-4 text-sm">
        <ConfigRow label="Title"><Input className="h-7 text-xs" value={w.title} onChange={(e) => onChange({ title: e.target.value })} /></ConfigRow>
        <ConfigRow label="Image URL">
          <Input className="h-7 text-xs" placeholder="https://…" value={w.url} onChange={(e) => onChange({ url: e.target.value })} />
        </ConfigRow>
        <ConfigRow label="Alt text"><Input className="h-7 text-xs" value={w.alt || ""} onChange={(e) => onChange({ alt: e.target.value })} /></ConfigRow>
        <ConfigRow label="Fit">
          <select className="w-full rounded border border-[#E5E5E5] p-1 text-xs" value={w.fit || "contain"} onChange={(e) => onChange({ fit: e.target.value as ImageWidget["fit"] })}>
            <option value="contain">Contain</option>
            <option value="cover">Cover</option>
            <option value="fill">Fill</option>
          </select>
        </ConfigRow>
      </div>
    );
  }

  if (type === "filter") {
    const w = widget as FilterWidget;
    return (
      <div className="flex flex-1 flex-col gap-3 overflow-auto p-4 text-sm">
        <ConfigRow label="Label"><Input className="h-7 text-xs" value={w.title} onChange={(e) => onChange({ title: e.target.value })} /></ConfigRow>
        <ConfigRow label="Filter column">
          <select className="w-full rounded border border-[#E5E5E5] p-1 text-xs" value={w.column} onChange={(e) => onChange({ column: e.target.value })}>
            <option value="">Select column…</option>
            {sheet.columns.map((c) => <option key={c.key} value={c.header}>{c.header}</option>)}
          </select>
        </ConfigRow>
        <ConfigRow label="Filter type">
          <select className="w-full rounded border border-[#E5E5E5] p-1 text-xs" value={w.filterType} onChange={(e) => onChange({ filterType: e.target.value as FilterWidget["filterType"] })}>
            <option value="multiselect">Multi-select</option>
            <option value="select">Single select</option>
          </select>
        </ConfigRow>
      </div>
    );
  }

  return null;
}

function ConfigRow({ label, children }: { label: string; children: React.ReactNode }) {
  return (
    <div className="space-y-1">
      <label className="text-xs font-medium text-[#6B6B6B]">{label}</label>
      {children}
    </div>
  );
}

export { CHART_TYPES };
