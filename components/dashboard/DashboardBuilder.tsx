"use client";

import { useState, useRef, useCallback, useEffect } from "react";
import { useDropzone } from "react-dropzone";
import {
  Upload, Save, Trash2, Loader2, BarChart2, Settings2,
  Eye, EyeOff, BarChart, Type, Hash, Image as ImageIcon, Filter, ChevronDown,
  Share2, ExternalLink, RefreshCw, FileText, GripVertical, ChevronLeft, ChevronRight,
  Plus, X, Database, FileSpreadsheet, Cloud, AlertTriangle, CheckCircle, GitBranch,
} from "lucide-react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { ChartRenderer } from "@/components/dashboard/ChartRenderer";
import { ChartConfigPanel } from "@/components/dashboard/ChartConfigPanel";
import { parseExcelFile, type ExcelSheet } from "@/lib/excel";
import { evaluate } from "mathjs";
import { useToast } from "@/hooks/use-toast";
import "react-grid-layout/css/styles.css";
import "react-resizable/css/styles.css";

// eslint-disable-next-line @typescript-eslint/no-require-imports
const _RGL = require("react-grid-layout");
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const GridLayout = (_RGL.default || _RGL) as any;

interface GridItem { i: string; x: number; y: number; w: number; h: number; }

export interface DataSource {
  id: string;
  name: string;
  type: "excel" | "smartsheet" | "mapping";
  sourceId?: string; // Smartsheet sheet ID
  mappingId?: string; // Source mapping ID
  autoUpdate?: boolean; // Auto-update when mapping runs
  sheet: ExcelSheet;
  lastRefreshed: string;
}

export interface LinkedMapping {
  mappingId: string;
  mappingName: string;
  autoUpdate: boolean;
}

interface MappingNodeData {
  label?: string;
  colRef?: string;
  formula?: string;
}
interface SimpleMappingNode { id: string; type: string; data: MappingNodeData; }
interface SimpleMappingEdge { id: string; source: string; target: string; }
interface SimpleMappingVersion {
  id: string;
  versionNumber: number;
  connections?: { nodes?: SimpleMappingNode[]; edges?: SimpleMappingEdge[] };
}
interface SimpleMapping {
  id: string;
  name: string;
  smartsheetSheetId: string | null;
  currentVersionId: string | null;
  versions: SimpleMappingVersion[];
}

function parseMappingOutputData(raw: unknown): Record<string, string | number | boolean | null>[] {
  if (Array.isArray(raw)) {
    return raw as Record<string, string | number | boolean | null>[];
  }
  if (typeof raw === "string") {
    try {
      const parsed = JSON.parse(raw);
      return Array.isArray(parsed) ? (parsed as Record<string, string | number | boolean | null>[]) : [];
    } catch {
      return [];
    }
  }
  return [];
}

function getMappingOutputHeaders(
  rows: Record<string, string | number | boolean | null>[],
  mapping?: SimpleMapping
): string[] {
  const keys = new Set<string>();
  for (const row of rows) {
    if (row && typeof row === "object") {
      Object.keys(row).forEach((k) => keys.add(k));
    }
  }
  if (keys.size > 0) return Array.from(keys);

  const version = mapping?.versions?.[0];
  const targetLabels = (version?.connections?.nodes ?? [])
    .filter((n) => n.type === "ssCol")
    .map((n) => n.data.label)
    .filter((v): v is string => typeof v === "string" && v.trim().length > 0);

  return Array.from(new Set(targetLabels));
}

export interface ChartConfig {
  id: string;
  widgetType?: "chart";
  title: string;
  type: ChartType;
  dataSourceId?: string; // Which data source this chart uses
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
  dataSourceId?: string; // Which data source this KPI uses
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
  dataSourceId?: string; // Which data source this filter uses
  column: string;
  filterType: "select" | "multiselect" | "range";
}

export interface WidgetFilter {
  id: string;
  column: string;
  type: "values" | "numRange" | "dateRange";
  values: string[];
  min: string;
  max: string;
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
    dataSources?: DataSource[];
    linkedMappings?: LinkedMapping[];
    charts: unknown[];
    layout: unknown[];
  };
  onSaved: (id: string) => void;
  initialViewMode?: boolean;
}

export function DashboardBuilder({ initialDashboard, onSaved, initialViewMode = false }: DashboardBuilderProps) {
  const { toast } = useToast();
  const [name, setName] = useState(initialDashboard?.name ?? "");

  // Multi-data source state
  const [dataSources, setDataSources] = useState<DataSource[]>(() => {
    // First check if dataSources already exists (new format)
    if (initialDashboard?.dataSources && initialDashboard.dataSources.length > 0) {
      return initialDashboard.dataSources;
    }
    // Migrate from legacy excelData if present
    if (initialDashboard?.excelData) {
      const legacySheet = initialDashboard.excelData as ExcelSheet;
      return [{
        id: `ds_${Date.now()}`,
        name: legacySheet.name,
        type: "excel",
        sheet: legacySheet,
        lastRefreshed: new Date().toISOString(),
      }];
    }
    return [];
  });
  const [linkedMappings, setLinkedMappings] = useState<LinkedMapping[]>(initialDashboard?.linkedMappings ?? []);

  const [widgets, setWidgets] = useState<DashboardWidget[]>(
    (initialDashboard?.charts as DashboardWidget[]) ?? []
  );
  const [layout, setLayout] = useState<GridItem[]>(
    (initialDashboard?.layout as GridItem[]) ?? []
  );
  const [selectedWidget, setSelectedWidget] = useState<string | null>(null);
  const [selectedDataSource, setSelectedDataSource] = useState<string | null>(null);
  const [saving, setSaving] = useState(false);
  const [viewMode, setViewMode] = useState(initialViewMode);
  const [addDropdown, setAddDropdown] = useState(false);
  const [containerWidth, setContainerWidth] = useState(1200);
  const [activeFilters, setActiveFilters] = useState<Record<string, string[]>>({});
  const [widgetFilters, setWidgetFilters] = useState<Record<string, WidgetFilter[]>>({});
  const [filterPanelOpen, setFilterPanelOpen] = useState<string | null>(null);
  const containerRef = useRef<HTMLDivElement>(null);
  const refreshInputRef = useRef<HTMLInputElement>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [refreshing, setRefreshing] = useState<string | null>(null); // Which data source is being refreshed
  const [addingSheet, setAddingSheet] = useState(false);
  const [smartsheetConnected, setSmartsheetConnected] = useState(false);
  const [smartsheetSheets, setSmartsheetSheets] = useState<Array<{ id: number; name: string }>>([]);
  const [addingFromMapping, setAddingFromMapping] = useState(false);
  const [availableMappings, setAvailableMappings] = useState<SimpleMapping[]>([]);
  const [loadingMappings, setLoadingMappings] = useState(false);
  const [applyingMapping, setApplyingMapping] = useState(false);
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
    if (viewMode) return;
    const hasAutoMappings = dataSources.some((d) => d.type === "mapping" && d.autoUpdate && d.mappingId);
    if (!hasAutoMappings) return;

    const timer = window.setInterval(async () => {
      const mappingSources = dataSources.filter((d) => d.type === "mapping" && d.autoUpdate && d.mappingId);
      for (const ds of mappingSources) {
        try {
          const res = await fetch(`/api/mappings/${ds.mappingId}/runs/latest`);
          if (!res.ok) continue;
          const run = await res.json();
          const outputData = parseMappingOutputData(run.outputData);
          const outputCreatedAt = run.createdAt as string;
          if (!outputData || !outputCreatedAt) continue;

          setDataSources((prev) => prev.map((d) => {
            if (d.id !== ds.id) return d;
            const current = new Date(d.lastRefreshed).getTime();
            const incoming = new Date(outputCreatedAt).getTime();
            if (Number.isNaN(incoming) || incoming <= current) return d;
            const headers = getMappingOutputHeaders(outputData);
            return {
              ...d,
              sheet: {
                ...d.sheet,
                columns: headers.map((header, idx) => ({ key: header, header, index: idx, dataType: "string" as const, sampleValues: [] })),
                data: outputData,
                rowCount: outputData.length,
              },
              lastRefreshed: outputCreatedAt,
            };
          }));
        } catch {
          // ignore polling errors
        }
      }
    }, 5000);

    return () => window.clearInterval(timer);
  }, [viewMode, dataSources]);

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

  // Check if Smartsheet is connected on mount
  useEffect(() => {
    fetch("/api/settings/smartsheet")
      .then((res) => res.json())
      .then((data) => {
        if (data.connected) {
          setSmartsheetConnected(true);
          fetchSmartsheetSheets();
        }
      })
      .catch(() => {/* ignore */});
  }, []);

  const fetchSmartsheetSheets = async () => {
    try {
      const res = await fetch("/api/smartsheet/sheets");
      if (res.ok) {
        const sheets = await res.json();
        setSmartsheetSheets(sheets);
      }
    } catch (e) {
      console.error("Failed to fetch Smartsheet sheets:", e);
    }
  };

  const onDrop = useCallback(async (files: File[]) => {
    const file = files[0];
    if (!file) return;
    const buf = await file.arrayBuffer();
    const parsed = parseExcelFile(buf, file.name, file.size);
    if (parsed.sheets.length === 0) {
      toast({ title: "No data found", variant: "destructive" });
      return;
    }

    // Add all sheets from the file as data sources
    const newDataSources: DataSource[] = parsed.sheets.map((sheet, index) => ({
      id: `ds_${Date.now()}_${index}`,
      name: sheet.name,
      type: "excel" as const,
      sheet,
      lastRefreshed: new Date().toISOString(),
    }));

    setDataSources((prev) => [...prev, ...newDataSources]);
    toast({
      title: `Loaded: ${file.name}`,
      description: `${parsed.sheets.length} sheet(s) added as data sources`,
    });
  }, [toast]);

  const { getRootProps, getInputProps, isDragActive } = useDropzone({
    onDrop,
    accept: {
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"],
      "application/vnd.ms-excel": [".xls"],
      "text/csv": [".csv"],
    },
    noClick: false, // Allow clicking to add more sheets
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
    setWidgetFilters((prev) => { const next = { ...prev }; delete next[id]; return next; });
    if (filterPanelOpen === id) setFilterPanelOpen(null);
  }

  function updateWidget(id: string, updates: Partial<DashboardWidget>) {
    setWidgets((prev) => prev.map((w) => (w.id === id ? { ...w, ...updates } as DashboardWidget : w)));
  }

  // Get sheet for a specific widget
  function getWidgetSheet(widget: DashboardWidget): ExcelSheet | null {
    const dataSourceId = (widget as ChartConfig | KPIWidget | FilterWidget).dataSourceId;
    if (!dataSourceId) {
      // Default to first data source for backward compatibility
      return dataSources[0]?.sheet ?? null;
    }
    return dataSources.find((ds) => ds.id === dataSourceId)?.sheet ?? null;
  }

  function getFilteredSheet(forDataSourceId?: string): ExcelSheet | null {
    const ds = forDataSourceId
      ? dataSources.find((d) => d.id === forDataSourceId)
      : dataSources[0];
    if (!ds) return null;

    const sheet = ds.sheet;
    const filterCols = Object.keys(activeFilters).filter((k) => activeFilters[k].length > 0);
    if (filterCols.length === 0) return sheet;

    const filteredData = sheet.data.filter((row: Record<string, unknown>) =>
      filterCols.every((col) => activeFilters[col].includes(String(row[col] ?? "")))
    );
    return { ...sheet, data: filteredData, rowCount: filteredData.length };
  }

  // Add a Smartsheet as data source
  async function addSmartsheetDataSource(sheetId: number, sheetName: string) {
    try {
      const res = await fetch(`/api/smartsheet/sheets/${sheetId}/data`);
      if (!res.ok) throw new Error("Failed to fetch Smartsheet data");

      const data = await res.json();
      const newDataSource: DataSource = {
        id: `ds_smartsheet_${Date.now()}`,
        name: sheetName,
        type: "smartsheet",
        sourceId: String(sheetId),
        sheet: data.sheet,
        lastRefreshed: data.lastRefreshed,
      };

      setDataSources((prev) => [...prev, newDataSource]);
      toast({ title: "Smartsheet added", description: `${sheetName} (${data.sheet.rowCount} rows)` });
      setAddingSheet(false);
    } catch (e) {
      toast({ title: "Failed to add Smartsheet", description: (e as Error).message, variant: "destructive" });
    }
  }

  // Remove a data source
  function removeDataSource(id: string) {
    setDataSources((prev) => prev.filter((ds) => ds.id !== id));
    // Also remove dataSourceId from widgets that used this source
    setWidgets((prev) => prev.map((w) => {
      if ((w as ChartConfig | KPIWidget | FilterWidget).dataSourceId === id) {
        return { ...w, dataSourceId: undefined } as DashboardWidget;
      }
      return w;
    }));
    toast({ title: "Data source removed" });
  }

  // Smart refresh with column compatibility check
  async function refreshDataSource(dataSourceId: string) {
    const ds = dataSources.find((d) => d.id === dataSourceId);
    if (!ds) return;

    setRefreshing(dataSourceId);
    try {
      let newSheet: ExcelSheet;

      if (ds.type === "smartsheet" && ds.sourceId) {
        // Refresh from Smartsheet
        const res = await fetch(`/api/smartsheet/sheets/${ds.sourceId}/data`);
        if (!res.ok) throw new Error("Failed to refresh from Smartsheet");
        const data = await res.json();
        newSheet = data.sheet;
      } else {
        // For Excel, user needs to re-upload - trigger file input
        setSelectedDataSource(dataSourceId);
        fileInputRef.current?.click();
        return;
      }

      // Check column compatibility
      const oldColumns = ds.sheet.columns.map((c) => c.header);
      const newColumns = newSheet.columns.map((c) => c.header);
      const missingColumns = oldColumns.filter((c) => !newColumns.includes(c));

      if (missingColumns.length > 0) {
        toast({
          title: "Column mismatch detected",
          description: `Missing columns: ${missingColumns.join(", ")}. Charts may need reconfiguration.`,
          variant: "destructive",
        });
      }

      // Update the data source
      setDataSources((prev) =>
        prev.map((d) =>
          d.id === dataSourceId
            ? { ...d, sheet: newSheet, lastRefreshed: new Date().toISOString() }
            : d
        )
      );

      toast({
        title: "Data refreshed!",
        description: `${newSheet.rowCount.toLocaleString()} rows · ${newSheet.columns.length} cols`,
      });
    } catch (err) {
      toast({ title: "Refresh failed", description: (err as Error).message, variant: "destructive" });
    } finally {
      setRefreshing(null);
    }
  }

  // Handle Excel file refresh for a specific data source
  async function handleExcelRefresh(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    if (!file || !selectedDataSource) return;

    const ds = dataSources.find((d) => d.id === selectedDataSource);
    if (!ds) return;

    try {
      const buf = await file.arrayBuffer();
      const parsed = parseExcelFile(buf, file.name, file.size);

      // Find sheet with matching name, or use first sheet
      const newSheet = parsed.sheets.find((s) => s.name === ds.sheet.name) || parsed.sheets[0];
      if (!newSheet) {
        toast({ title: "No matching sheet found", variant: "destructive" });
        return;
      }

      // Check column compatibility
      const oldColumns = ds.sheet.columns.map((c) => c.header);
      const newColumns = newSheet.columns.map((c) => c.header);
      const missingColumns = oldColumns.filter((c) => !newColumns.includes(c));

      if (missingColumns.length > 0) {
        toast({
          title: "Column mismatch detected",
          description: `Missing columns: ${missingColumns.join(", ")}. Charts may need reconfiguration.`,
          variant: "destructive",
        });
      }

      // Update the data source
      setDataSources((prev) =>
        prev.map((d) =>
          d.id === selectedDataSource
            ? { ...d, sheet: newSheet, lastRefreshed: new Date().toISOString() }
            : d
        )
      );

      toast({
        title: "Data refreshed!",
        description: `${newSheet.rowCount.toLocaleString()} rows · ${newColumns.length} cols`,
      });
    } catch (err) {
      toast({ title: "Refresh failed", description: (err as Error).message, variant: "destructive" });
    } finally {
      setSelectedDataSource(null);
      if (fileInputRef.current) fileInputRef.current.value = "";
    }
  }

  async function handleSave() {
    if (!name.trim()) { toast({ title: "Enter a dashboard name", variant: "destructive" }); return; }
    setSaving(true);
    try {
      const url = initialDashboard ? `/api/dashboards/${initialDashboard.id}` : "/api/dashboards";
      const method = initialDashboard ? "PUT" : "POST";

      // Save first data source as excelData for backward compatibility
      const excelData = dataSources[0]?.sheet || null;

      const res = await fetch(url, {
        method,
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ name, excelData, dataSources, charts: widgets, layout, linkedMappings }),
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

  const activeWidget = widgets.find((w) => w.id === selectedWidget);
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
            <Button size="sm" variant="ghost" className="h-7 gap-1 text-xs" onClick={() => setAddDropdown((v) => !v)} disabled={dataSources.length === 0}>
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
              className="no-print flex w-9 shrink-0 cursor-pointer flex-col items-center border-r border-[#E5E5E5] bg-white py-3 hover:bg-[#F9F9F9] transition-colors overflow-hidden"
              onClick={() => setLeftMin(false)}
              title="Expand data panel"
            >
              <ChevronRight className="h-4 w-4 text-[#A1A1A1] shrink-0" />
              <span className="mt-2 text-[8px] uppercase tracking-widest text-[#C0C0C0] whitespace-nowrap" style={{ writingMode: "vertical-rl", transform: "rotate(180deg)" }}>Data</span>
            </div>
          ) : (
            <>
              <div
                ref={leftPanelRef}
                className="no-print flex shrink-0 flex-col border-r border-[#E5E5E5] bg-white"
                style={{ width: leftW }}
              >
                <div className="flex items-center justify-between border-b border-[#E5E5E5] px-4 py-2.5">
                  <p className="text-xs font-semibold uppercase tracking-wider text-[#6B6B6B]">Data Sources ({dataSources.length})</p>
                  <button onClick={() => setLeftMin(true)} className="rounded p-0.5 hover:bg-[#F5F5F5]" title="Collapse">
                    <ChevronLeft className="h-3.5 w-3.5 text-[#A1A1A1]" />
                  </button>
                </div>
                <div className="flex flex-1 flex-col overflow-auto p-3">
                  {/* Add Excel Sheet */}
                  <div
                    {...getRootProps()}
                    className={`mb-2 flex cursor-pointer flex-col items-center justify-center rounded-lg border-2 border-dashed p-3 text-center transition-colors ${
                      isDragActive ? "border-black bg-[#F5F5F5]" : "border-[#E5E5E5] hover:border-[#A1A1A1]"
                    }`}
                  >
                    <input {...getInputProps()} />
                    <Upload className="mb-1 h-4 w-4 text-[#A1A1A1]" />
                    <p className="text-xs font-medium">Add Excel Sheet</p>
                    <p className="text-[9px] text-[#6B6B6B]">Drop file or click</p>
                  </div>

                  {/* Add Smartsheet Button */}
                  {smartsheetConnected && (
                    <button
                      onClick={() => setAddingSheet(true)}
                      className="mb-3 flex items-center justify-center gap-2 rounded-lg border border-[#E5E5E5] bg-white p-2 text-xs hover:bg-[#F5F5F5] transition-colors"
                    >
                      <Cloud className="h-3.5 w-3.5 text-[#6B6B6B]" />
                      Add Smartsheet
                    </button>
                  )}

                  {/* Smartsheet Selector Dialog */}
                  {addingSheet && smartsheetConnected && (
                    <div className="mb-3 rounded-lg border border-[#E5E5E5] bg-[#F9F9F9] p-2">
                      <div className="flex items-center justify-between mb-2">
                        <p className="text-[10px] font-medium text-[#6B6B6B]">Select Smartsheet</p>
                        <button onClick={() => setAddingSheet(false)} className="text-[#A1A1A1] hover:text-[#6B6B6B]">
                          <X className="h-3 w-3" />
                        </button>
                      </div>
                      <div className="max-h-32 overflow-auto space-y-1">
                        {smartsheetSheets.map((ss) => (
                          <button
                            key={ss.id}
                            onClick={() => addSmartsheetDataSource(ss.id, ss.name)}
                            className="w-full text-left rounded px-2 py-1.5 text-[10px] hover:bg-white hover:shadow-sm transition-all"
                          >
                            <div className="flex items-center gap-1.5">
                              <FileSpreadsheet className="h-3 w-3 text-[#6B6B6B]" />
                              <span className="truncate">{ss.name}</span>
                            </div>
                          </button>
                        ))}
                      </div>
                    </div>
                  )}

                  {/* Add from Mapping Button */}
                  <button
                    onClick={async () => {
                      if (availableMappings.length === 0) {
                        setLoadingMappings(true);
                        try {
                          const res = await fetch("/api/mappings");
                          if (res.ok) {
                            const data = await res.json();
                            setAvailableMappings(data.filter((m: SimpleMapping) => !m.smartsheetSheetId));
                          }
                        } catch (e) {
                          console.error("Failed to fetch mappings:", e);
                        }
                        setLoadingMappings(false);
                      }
                      setAddingFromMapping(true);
                    }}
                    className="mb-3 flex items-center justify-center gap-2 rounded-lg border border-[#E5E5E5] bg-white p-2 text-xs hover:bg-[#F5F5F5] transition-colors"
                  >
                    <GitBranch className="h-3.5 w-3.5 text-[#6B6B6B]" />
                    Add from Mapping
                  </button>

                  {/* Mapping Selector Dialog */}
                  {addingFromMapping && (
                    <div className="mb-3 rounded-lg border border-[#E5E5E5] bg-[#F9F9F9] p-2">
                      <div className="flex items-center justify-between mb-2">
                        <p className="text-[10px] font-medium text-[#6B6B6B]">Select Mapping (uses latest output)</p>
                        <button onClick={() => setAddingFromMapping(false)} className="text-[#A1A1A1] hover:text-[#6B6B6B]">
                          <X className="h-3 w-3" />
                        </button>
                      </div>
                      {loadingMappings ? (
                        <div className="flex items-center justify-center py-2">
                          <Loader2 className="h-4 w-4 animate-spin text-[#A1A1A1]" />
                        </div>
                      ) : availableMappings.length === 0 ? (
                        <p className="text-[10px] text-[#A1A1A1] py-2">No mappings found (Excel→Excel only)</p>
                      ) : (
                        <div className="max-h-32 overflow-auto space-y-1">
                          {availableMappings.map((m) => (
                            <button
                              key={m.id}
                              onClick={async () => {
                                setApplyingMapping(true);
                                try {
                                  const res = await fetch(`/api/mappings/${m.id}/runs/latest`);
                                  if (!res.ok) {
                                    if (res.status === 404) {
                                      toast({ title: "No runs yet", description: "Run the mapping first from the Mapper page", variant: "destructive" });
                                    } else {
                                      throw new Error("Failed to fetch latest run");
                                    }
                                    return;
                                  }
                                  const run = await res.json();
                                  const outputData = parseMappingOutputData(run.outputData);
                                  if (!outputData || outputData.length === 0) {
                                    throw new Error("No output data in run");
                                  }
                                  const headers = getMappingOutputHeaders(outputData, m);
                                  const columns = headers.map((header, idx) => ({
                                    key: header,
                                    header,
                                    index: idx,
                                    dataType: "string" as const,
                                    sampleValues: [],
                                  }));
                                  const mappedSheet: ExcelSheet = {
                                    name: `${m.name} Output`,
                                    columns,
                                    data: outputData,
                                    rowCount: outputData.length,
                                  };
                                  setDataSources((prev) => [...prev, {
                                    id: `ds_mapping_${Date.now()}`,
                                    name: m.name,
                                    type: "mapping",
                                    mappingId: m.id,
                                    autoUpdate: true,
                                    sheet: mappedSheet,
                                    lastRefreshed: run.createdAt || new Date().toISOString(),
                                  }]);
                                  setLinkedMappings((prev) => {
                                    if (prev.some((lm) => lm.mappingId === m.id)) return prev;
                                    return [...prev, { mappingId: m.id, mappingName: m.name, autoUpdate: true }];
                                  });
                                  toast({ title: "Mapping linked", description: `${outputData.length} rows loaded from latest run` });
                                  setAddingFromMapping(false);
                                } catch (err) {
                                  toast({ title: "Failed to load mapping", description: (err as Error).message, variant: "destructive" });
                                } finally {
                                  setApplyingMapping(false);
                                }
                              }}
                              className="w-full text-left rounded px-2 py-1.5 text-[10px] hover:bg-white hover:shadow-sm transition-all"
                            >
                              <div className="flex items-center gap-1.5">
                                <GitBranch className="h-3 w-3 text-purple-500" />
                                <span className="truncate">{m.name}</span>
                              </div>
                            </button>
                          ))}
                        </div>
                      )}
                    </div>
                  )}

                  {/* Data Sources List */}
                  <div className="space-y-2">
                    {dataSources.map((ds) => (
                      <div
                        key={ds.id}
                        className={`rounded-lg border p-2 transition-all ${
                          selectedDataSource === ds.id
                            ? "border-blue-400 bg-blue-50"
                            : "border-[#E5E5E5] bg-white hover:border-[#A1A1A1]"
                        }`}
                      >
                        <div className="flex items-center justify-between mb-1">
                          <div className="flex items-center gap-1.5 min-w-0">
                            {ds.type === "smartsheet" ? (
                              <Cloud className="h-3 w-3 text-blue-500 shrink-0" />
                            ) : ds.type === "mapping" ? (
                              <GitBranch className="h-3 w-3 text-purple-500 shrink-0" />
                            ) : (
                              <FileSpreadsheet className="h-3 w-3 text-green-500 shrink-0" />
                            )}
                            <span className="text-[11px] font-medium truncate">{ds.name}</span>
                          </div>
                          <div className="flex items-center gap-0.5 shrink-0">
                            {ds.type !== "mapping" && (
                              <button
                                onClick={() => refreshDataSource(ds.id)}
                                disabled={refreshing === ds.id}
                                className="rounded p-1 hover:bg-[#F5F5F5] text-[#6B6B6B]"
                                title="Refresh data"
                              >
                                {refreshing === ds.id ? (
                                  <Loader2 className="h-3 w-3 animate-spin" />
                                ) : (
                                  <RefreshCw className="h-3 w-3" />
                                )}
                              </button>
                            )}
                            <button
                              onClick={() => removeDataSource(ds.id)}
                              className="rounded p-1 hover:bg-red-50 text-red-400"
                              title="Remove data source"
                            >
                              <X className="h-3 w-3" />
                            </button>
                          </div>
                        </div>
                        <div className="flex flex-wrap items-center gap-2 text-[9px] text-[#A1A1A1]">
                          <span>{ds.sheet.rowCount.toLocaleString()} rows</span>
                          <span>·</span>
                          <span>{ds.sheet.columns.length} cols</span>
                          <span>·</span>
                          <span>{new Date(ds.lastRefreshed).toLocaleDateString()}</span>
                          {ds.type === "mapping" && (
                            <>
                              <span>·</span>
                              <button
                                onClick={() => {
                                  setDataSources((prev) => prev.map((d) =>
                                    d.id === ds.id ? { ...d, autoUpdate: !d.autoUpdate } : d
                                  ));
                                  if (ds.mappingId) {
                                    setLinkedMappings((prev) => prev.map((lm) =>
                                      lm.mappingId === ds.mappingId ? { ...lm, autoUpdate: !ds.autoUpdate } : lm
                                    ));
                                  }
                                }}
                                className={`rounded px-1 py-0.5 ${ds.autoUpdate ? "bg-green-100 text-green-700" : "bg-gray-100 text-[#6B6B6B]"}`}
                                title="Toggle automatic updates from mapping runs"
                              >
                                Auto {ds.autoUpdate ? "On" : "Off"}
                              </button>
                               <button
                                 onClick={async () => {
                                  if (ds.mappingId) {
                                    try {
                                      const res = await fetch(`/api/mappings/${ds.mappingId}/runs/latest`);
                                      if (res.ok) {
                                        const run = await res.json();
                                        const outputData = parseMappingOutputData(run.outputData);
                                        if (outputData) {
                                          const headers = getMappingOutputHeaders(outputData);
                                          setDataSources((prev) => prev.map((d) =>
                                            d.id === ds.id
                                              ? {
                                                  ...d,
                                                  sheet: {
                                                    ...d.sheet,
                                                    columns: headers.map((header, idx) => ({ key: header, header, index: idx, dataType: "string" as const, sampleValues: [] })),
                                                    data: outputData,
                                                    rowCount: outputData.length,
                                                  },
                                                  lastRefreshed: run.createdAt || new Date().toISOString(),
                                                }
                                              : d
                                          ));
                                          toast({ title: "Refreshed from latest run" });
                                        }
                                      }
                                    } catch (err) {
                                      toast({ title: "Failed to refresh", description: (err as Error).message, variant: "destructive" });
                                    }
                                  }
                                }}
                                className="hidden items-center gap-1 rounded px-1 py-0.5 hover:bg-[#F5F5F5] text-[#6B6B6B]"
                                title="Refresh from latest mapping run"
                              >
                                <RefreshCw className="h-2.5 w-2.5" />
                                <span>Auto</span>
                              </button>
                            </>
                          )}
                        </div>
                        {/* Columns preview */}
                        <div className="mt-1.5 max-h-20 overflow-auto space-y-0.5">
                          {ds.sheet.columns.slice(0, 5).map((col: { key: string; header: string; dataType: string }) => (
                            <div key={col.key} className="flex items-center gap-1 text-[9px]">
                              <span className="truncate text-[#6B6B6B]">{col.header}</span>
                              <span className="ml-auto text-[#A1A1A1]">{col.dataType}</span>
                            </div>
                          ))}
                          {ds.sheet.columns.length > 5 && (
                            <p className="text-[9px] text-[#A1A1A1]">+{ds.sheet.columns.length - 5} more columns</p>
                          )}
                        </div>
                      </div>
                    ))}
                  </div>

                  {dataSources.length === 0 && (
                    <div className="flex flex-col items-center justify-center py-8 text-center">
                      <Database className="h-8 w-8 text-[#D1D1D1] mb-2" />
                      <p className="text-xs text-[#A1A1A1]">No data sources</p>
                      <p className="text-[10px] text-[#C0C0C0] mt-1">Upload Excel or add Smartsheet</p>
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
              </div>

              {/* Grid area */}
              <div ref={containerRef}>
                {widgets.length === 0 ? (
                  <div className="flex flex-col items-center justify-center py-24">
                    <BarChart2 className="mb-3 h-10 w-10 text-[#D1D1D1]" />
                    <p className="text-sm text-[#A1A1A1]">
                      {dataSources.length > 0
                        ? 'Use "+  Insert" in the toolbar to add charts, KPIs, or text'
                        : "Upload an Excel file or add Smartsheet from the left panel to begin"}
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
                            dataSources={dataSources}
                            activeFilters={activeFilters}
                            onFilterChange={(col, vals) => setActiveFilters((prev) => ({ ...prev, [col]: vals }))}
                            widgetFilters={widgetFilters[widget.id] || []}
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

                        {/* Per-widget filter button — chart & KPI only, works in both edit and view modes */}
                        {(widget.widgetType === "chart" || widget.widgetType === "kpi" || widget.widgetType === undefined) && (() => {
                          const wf = widgetFilters[widget.id] || [];
                          const activeCount = wf.filter((f) =>
                            f.type === "values" ? f.values.length > 0 : f.min !== "" || f.max !== ""
                          ).length;
                          return (
                            <div
                              className={`absolute left-1 top-5 z-20 transition-all ${
                                activeCount > 0
                                  ? "opacity-100"
                                  : "pointer-events-none opacity-0 group-hover:pointer-events-auto group-hover:opacity-100"
                              }`}
                              onClick={(e) => e.stopPropagation()}
                            >
                              <button
                                className={`flex items-center gap-1 rounded-md px-1.5 py-1 shadow-md text-[10px] font-medium transition-colors ${
                                  activeCount > 0
                                    ? "bg-blue-500 text-white hover:bg-blue-600"
                                    : "bg-white text-[#6B6B6B] hover:bg-[#F5F5F5]"
                                }`}
                                onClick={() => setFilterPanelOpen((prev) => (prev === widget.id ? null : widget.id))}
                                title="Snapshot filters"
                              >
                                <Filter className="h-3 w-3" />
                                {activeCount > 0 && <span>{activeCount}</span>}
                              </button>
                            </div>
                          );
                        })()}

                        {/* Filter panel overlay */}
                        {filterPanelOpen === widget.id && (widget.widgetType === "chart" || widget.widgetType === "kpi" || widget.widgetType === undefined) && (
                          <div className="absolute inset-x-0 top-0 z-50" onClick={(e) => e.stopPropagation()}>
                            <WidgetFilterPanel
                              widget={widget}
                              dataSources={dataSources}
                              filters={widgetFilters[widget.id] || []}
                              onChange={(filters) => setWidgetFilters((prev) => ({ ...prev, [widget.id]: filters }))}
                              onClose={() => setFilterPanelOpen(null)}
                            />
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
              className="no-print flex w-9 shrink-0 cursor-pointer flex-col items-center border-l border-[#E5E5E5] bg-white py-3 hover:bg-[#F9F9F9] transition-colors overflow-hidden"
              onClick={() => setRightMin(false)}
              title="Expand config panel"
            >
              <ChevronLeft className="h-4 w-4 text-[#A1A1A1] shrink-0" />
              <span className="mt-2 text-[8px] uppercase tracking-widest text-[#C0C0C0] whitespace-nowrap" style={{ writingMode: "vertical-rl", transform: "rotate(180deg)" }}>Config</span>
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
              {activeWidget && dataSources.length > 0 ? (
                <WidgetConfigPanel
                  widget={activeWidget}
                  dataSources={dataSources}
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

      {/* Hidden file input for Excel refresh */}
      <input
        ref={fileInputRef}
        type="file"
        accept=".xlsx,.xls,.csv"
        className="hidden"
        onChange={handleExcelRefresh}
      />
    </div>
  );
}

/* ─── Widget Filter Panel ────────────────────────────────────── */

interface WidgetFilterPanelProps {
  widget: DashboardWidget;
  dataSources: DataSource[];
  filters: WidgetFilter[];
  onChange: (filters: WidgetFilter[]) => void;
  onClose: () => void;
}

function WidgetFilterPanel({ widget, dataSources, filters, onChange, onClose }: WidgetFilterPanelProps) {
  const dataSourceId = (widget as ChartConfig | KPIWidget).dataSourceId;
  const ds = dataSourceId ? dataSources.find((d) => d.id === dataSourceId) : dataSources[0];
  const sheet = ds?.sheet;

  function addFilter() {
    if (!sheet || sheet.columns.length === 0) return;
    const col = sheet.columns[0];
    const type: WidgetFilter["type"] = col.dataType === "number" ? "numRange" : col.dataType === "date" ? "dateRange" : "values";
    onChange([...filters, { id: `wf_${Date.now()}`, column: col.header, type, values: [], min: "", max: "" }]);
  }

  function updateFilter(id: string, updates: Partial<WidgetFilter>) {
    onChange(filters.map((f) => (f.id === id ? { ...f, ...updates } : f)));
  }

  function removeFilter(id: string) {
    onChange(filters.filter((f) => f.id !== id));
  }

  function handleColumnChange(filterId: string, colName: string) {
    if (!sheet) return;
    const col = sheet.columns.find((c) => c.header === colName);
    if (!col) return;
    const type: WidgetFilter["type"] = col.dataType === "number" ? "numRange" : col.dataType === "date" ? "dateRange" : "values";
    updateFilter(filterId, { column: colName, type, values: [], min: "", max: "" });
  }

  if (!sheet) {
    return (
      <div className="rounded-lg border border-[#E5E5E5] bg-white p-3 shadow-xl text-xs text-[#A1A1A1]">
        No data source connected
      </div>
    );
  }

  return (
    <div className="rounded-lg border border-[#E5E5E5] bg-white shadow-xl">
      <div className="flex items-center justify-between border-b border-[#E5E5E5] px-3 py-2">
        <p className="text-[11px] font-semibold text-[#0A0A0A]">Snapshot Filters</p>
        <div className="flex items-center gap-2">
          {filters.length > 0 && (
            <button className="text-[10px] text-[#6B6B6B] underline hover:text-[#0A0A0A]" onClick={() => onChange([])}>
              Clear all
            </button>
          )}
          <button className="rounded p-0.5 hover:bg-[#F5F5F5]" onClick={onClose}>
            <X className="h-3 w-3 text-[#6B6B6B]" />
          </button>
        </div>
      </div>

      <div className="max-h-72 overflow-auto p-2 space-y-2">
        {filters.length === 0 && (
          <p className="py-3 text-center text-[11px] text-[#A1A1A1]">No filters — click &ldquo;Add filter&rdquo; below</p>
        )}
        {filters.map((filter) => {
          const uniqueVals = Array.from(
            new Set(sheet.data.map((r) => String(r[filter.column] ?? "")))
          ).sort();
          return (
            <div key={filter.id} className="rounded-md border border-[#E5E5E5] bg-[#F9F9F9] p-2 space-y-1.5">
              <div className="flex items-center gap-1.5">
                <select
                  className="flex-1 min-w-0 rounded border border-[#E5E5E5] bg-white p-1 text-[10px] font-medium"
                  value={filter.column}
                  onChange={(e) => handleColumnChange(filter.id, e.target.value)}
                >
                  {sheet.columns.map((c) => (
                    <option key={c.key} value={c.header}>{c.header}</option>
                  ))}
                </select>
                <span className={`shrink-0 rounded-full px-1.5 py-0.5 text-[9px] font-medium ${
                  filter.type === "numRange" ? "bg-blue-100 text-blue-700" :
                  filter.type === "dateRange" ? "bg-purple-100 text-purple-700" :
                  "bg-gray-100 text-gray-600"
                }`}>
                  {filter.type === "numRange" ? "number" : filter.type === "dateRange" ? "date" : "values"}
                </span>
                <button className="shrink-0 rounded p-0.5 hover:bg-red-50" onClick={() => removeFilter(filter.id)}>
                  <X className="h-3 w-3 text-[#A1A1A1]" />
                </button>
              </div>

              {filter.type === "numRange" && (
                <div className="flex items-center gap-1.5">
                  <input
                    type="number"
                    placeholder="Min"
                    className="w-full rounded border border-[#E5E5E5] bg-white px-2 py-1 text-[10px] outline-none focus:border-blue-400"
                    value={filter.min}
                    onChange={(e) => updateFilter(filter.id, { min: e.target.value })}
                  />
                  <span className="shrink-0 text-[10px] text-[#A1A1A1]">–</span>
                  <input
                    type="number"
                    placeholder="Max"
                    className="w-full rounded border border-[#E5E5E5] bg-white px-2 py-1 text-[10px] outline-none focus:border-blue-400"
                    value={filter.max}
                    onChange={(e) => updateFilter(filter.id, { max: e.target.value })}
                  />
                </div>
              )}

              {filter.type === "dateRange" && (
                <div className="flex items-center gap-1.5">
                  <input
                    type="date"
                    className="flex-1 min-w-0 rounded border border-[#E5E5E5] bg-white px-2 py-1 text-[10px] outline-none focus:border-purple-400"
                    value={filter.min}
                    onChange={(e) => updateFilter(filter.id, { min: e.target.value })}
                  />
                  <span className="shrink-0 text-[10px] text-[#A1A1A1]">–</span>
                  <input
                    type="date"
                    className="flex-1 min-w-0 rounded border border-[#E5E5E5] bg-white px-2 py-1 text-[10px] outline-none focus:border-purple-400"
                    value={filter.max}
                    onChange={(e) => updateFilter(filter.id, { max: e.target.value })}
                  />
                </div>
              )}

              {filter.type === "values" && (
                <div className="flex flex-wrap gap-1 max-h-24 overflow-auto">
                  {uniqueVals.slice(0, 24).map((val) => (
                    <button
                      key={val}
                      className={`rounded-full px-2 py-0.5 text-[10px] border transition-colors ${
                        filter.values.length === 0 || filter.values.includes(val)
                          ? "border-black bg-black text-white"
                          : "border-[#E5E5E5] text-[#6B6B6B] hover:border-black"
                      }`}
                      onClick={() => {
                        if (filter.values.length === 0) {
                          updateFilter(filter.id, { values: uniqueVals.filter((v) => v !== val) });
                        } else if (filter.values.includes(val)) {
                          const next = filter.values.filter((v) => v !== val);
                          updateFilter(filter.id, { values: next.length === uniqueVals.length ? [] : next });
                        } else {
                          updateFilter(filter.id, { values: [...filter.values, val] });
                        }
                      }}
                    >
                      {val || "(blank)"}
                    </button>
                  ))}
                  {uniqueVals.length > 24 && (
                    <span className="self-center text-[9px] text-[#A1A1A1]">+{uniqueVals.length - 24} more</span>
                  )}
                </div>
              )}
            </div>
          );
        })}
      </div>

      <div className="border-t border-[#E5E5E5] p-2">
        <button
          className="flex w-full items-center justify-center gap-1 rounded-md border border-dashed border-[#E5E5E5] py-1.5 text-[11px] text-[#6B6B6B] hover:border-[#A1A1A1] hover:text-[#0A0A0A] transition-colors"
          onClick={addFilter}
        >
          <Plus className="h-3 w-3" />
          Add filter
        </button>
      </div>
    </div>
  );
}

/* ─── Widget Renderer ─────────────────────────────────────────── */

interface WidgetRendererProps {
  widget: DashboardWidget;
  dataSources: DataSource[];
  activeFilters: Record<string, string[]>;
  onFilterChange: (col: string, vals: string[]) => void;
  widgetFilters?: WidgetFilter[];
  canEdit: boolean;
}

function WidgetRenderer({ widget, dataSources, activeFilters, onFilterChange, widgetFilters, canEdit }: WidgetRendererProps) {
  const type = widget.widgetType ?? "chart";

  // Get the appropriate sheet for this widget
  const dataSourceId = (widget as ChartConfig | KPIWidget | FilterWidget).dataSourceId;
  const ds = dataSourceId
    ? dataSources.find((d) => d.id === dataSourceId)
    : dataSources[0];
  const sheet = ds?.sheet ?? null;

  // Apply filters to the sheet
  const getFilteredSheet = (): ExcelSheet | null => {
    if (!sheet) return null;
    let data: Record<string, unknown>[] = sheet.data as Record<string, unknown>[];

    // Apply global filters (from FilterWidget)
    const filterCols = Object.keys(activeFilters).filter((k) => activeFilters[k].length > 0);
    if (filterCols.length > 0) {
      data = data.filter((row) =>
        filterCols.every((col) => activeFilters[col].includes(String(row[col] ?? "")))
      );
    }

    // Apply per-widget snapshot filters
    if (widgetFilters && widgetFilters.length > 0) {
      for (const filter of widgetFilters) {
        if (filter.type === "values" && filter.values.length > 0) {
          data = data.filter((row) => filter.values.includes(String(row[filter.column] ?? "")));
        } else if (filter.type === "numRange" && (filter.min !== "" || filter.max !== "")) {
          data = data.filter((row) => {
            const v = parseFloat(String(row[filter.column] ?? ""));
            if (isNaN(v)) return true;
            if (filter.min !== "" && v < parseFloat(filter.min)) return false;
            if (filter.max !== "" && v > parseFloat(filter.max)) return false;
            return true;
          });
        } else if (filter.type === "dateRange" && (filter.min !== "" || filter.max !== "")) {
          data = data.filter((row) => {
            const raw = row[filter.column];
            const v = raw instanceof Date ? raw : new Date(String(raw ?? ""));
            if (isNaN(v.getTime())) return true;
            if (filter.min !== "" && v < new Date(filter.min)) return false;
            if (filter.max !== "" && v > new Date(filter.max)) return false;
            return true;
          });
        }
      }
    }

    if ((data as typeof sheet.data) === sheet.data) return sheet;
    return { ...sheet, data: data as typeof sheet.data, rowCount: data.length };
  };

  const filteredSheet = getFilteredSheet();

  if (type === "chart") {
    const chart = widget as ChartConfig;
    if (!filteredSheet || !chart.xColumn || chart.yColumns.length === 0) {
      return <EmptyPlaceholder text={canEdit ? "Configure chart →" : "No data"} />;
    }
    return <ChartRenderer chart={chart} sheet={filteredSheet} />;
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
    const value = filteredSheet ? computeKPI(w, filteredSheet) : null;
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
    if (!filteredSheet || !w.column) return <EmptyPlaceholder text={canEdit ? "Select column in config →" : "No filter"} />;
    const col = filteredSheet.columns.find((c: { header: string }) => c.header === w.column);
    if (!col) return <EmptyPlaceholder text="Column not found" />;
    const uniqueVals = Array.from(new Set(filteredSheet.data.map((r: Record<string, unknown>) => String(r[w.column] ?? "")))).sort();
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
  dataSources: DataSource[];
  onChange: (updates: Partial<DashboardWidget>) => void;
}

function WidgetConfigPanel({ widget, dataSources, onChange }: WidgetConfigPanelProps) {
  const type = widget.widgetType ?? "chart";

  // Get the current data source for this widget
  const currentDataSourceId = (widget as ChartConfig | KPIWidget | FilterWidget).dataSourceId;
  const selectedSheet = currentDataSourceId
    ? dataSources.find((ds) => ds.id === currentDataSourceId)?.sheet
    : dataSources[0]?.sheet;

  // Data source selector component
  const DataSourceSelector = () => (
    <div className="space-y-1.5">
      <Label className="text-xs">Data Source</Label>
      <Select
        value={currentDataSourceId || dataSources[0]?.id || ""}
        onValueChange={(v: string) => onChange({ dataSourceId: v })}
      >
        <SelectTrigger className="h-8 text-xs">
          <SelectValue />
        </SelectTrigger>
        <SelectContent>
          {dataSources.map((ds) => (
            <SelectItem key={ds.id} value={ds.id} className="text-xs">
              {ds.type === "smartsheet" ? "Smartsheet: " : "Excel: "}{ds.name}
            </SelectItem>
          ))}
        </SelectContent>
      </Select>
    </div>
  );

  if (type === "chart") {
    return (
      <div className="flex flex-1 flex-col overflow-hidden text-sm">
        <div className="shrink-0 px-4 pt-4 pb-0">
          <DataSourceSelector />
        </div>
        {selectedSheet && (
          <ChartConfigPanel
            chart={widget as ChartConfig}
            sheet={selectedSheet}
            onChange={onChange as (u: Partial<ChartConfig>) => void}
          />
        )}
      </div>
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
    const numericCols = selectedSheet?.columns.filter((c) => c.dataType === "number") || [];
    return (
      <div className="flex flex-1 flex-col gap-3 overflow-auto p-4 text-sm">
        <DataSourceSelector />
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
        <DataSourceSelector />
        <ConfigRow label="Label"><Input className="h-7 text-xs" value={w.title} onChange={(e) => onChange({ title: e.target.value })} /></ConfigRow>
        <ConfigRow label="Filter column">
          <select className="w-full rounded border border-[#E5E5E5] p-1 text-xs" value={w.column} onChange={(e) => onChange({ column: e.target.value })}>
            <option value="">Select column…</option>
            {selectedSheet?.columns.map((c) => <option key={c.key} value={c.header}>{c.header}</option>)}
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
