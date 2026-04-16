"use client";

import { useState, useCallback, useRef, useEffect } from "react";
import {
  ReactFlow,
  addEdge,
  useNodesState,
  useEdgesState,
  Controls,
  Background,
  BackgroundVariant,
  type Connection,
  type Edge,
  type Node,
  Panel,
  useReactFlow,
  MarkerType,
} from "@xyflow/react";
import "@xyflow/react/dist/style.css";
import {
  Upload,
  Save,
  Loader2,
  ChevronRight,
  Table2,
  Sheet,
  PanelLeft,
  PanelRight,
  Maximize2,
  Minimize2,
  Plus,
} from "lucide-react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Badge } from "@/components/ui/badge";
import { Switch } from "@/components/ui/switch";
import { Dialog, DialogContent, DialogHeader, DialogTitle } from "@/components/ui/dialog";
import { ExcelColumnNode } from "@/components/mapper/ExcelColumnNode";
import { SmartsheetColumnNode } from "@/components/mapper/SmartsheetColumnNode";
import { FormulaNode } from "@/components/mapper/FormulaNode";
import { SimpleEdge } from "@/components/mapper/SimpleEdge";
import { parseExcelFile, generateSchemaFingerprint, type ExcelSheet } from "@/lib/excel";
import { parseExcelFileComplex } from "@/lib/ComplexInput";
import { formatDynamicColumnName } from "@/lib/dynamic-column";
import { EXCEL_FORMULAS, FORMULA_CATEGORIES, getFormulasByCategory, searchFormulas } from "@/lib/excel-formulas";
import { BookOpen, Search, X, Eye } from "lucide-react";
import { Popover, PopoverContent, PopoverTrigger } from "@/components/ui/popover";
import { useToast } from "@/hooks/use-toast";
import { clsx, type ClassValue } from "clsx";
import { twMerge } from "tailwind-merge";
import { NodeActionsContext } from "@/components/mapper/NodeActionsContext";

function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

type TargetMode = "smartsheet" | "excel";

const nodeTypes = {
  excelCol: ExcelColumnNode,
  ssCol: SmartsheetColumnNode,
  formula: FormulaNode,
};
const edgeTypes = { simple: SimpleEdge };

interface SmartsheetColumn { id: number; title: string; type: string; index: number }
interface SmartsheetSheet { id: number; name: string; columns: SmartsheetColumn[]; rowCount: number }
interface HierarchyPreviewRow {
  rowId: number;
  rowNumber: number | null;
  parentId: number | null;
  depth: number;
  isParent: boolean;
  path: string;
  sectionPath: string;
  values: Record<string, unknown>;
}
interface PreviewState {
  open: boolean;
  title: string;
  columns: string[];
  rows: Record<string, unknown>[];
  rowCount: number;
  loading: boolean;
  hasHierarchy: boolean;
  hierarchyRows: HierarchyPreviewRow[];
}

interface MapperWorkspaceProps {
  initialMapping?: {
    id: string;
    name: string;
    autoPush?: boolean;
    smartsheetSheetId: string | null;
    smartsheetSheetName: string | null;
    currentVersionId: string | null;
    versions: { id: string; versionNumber: number; connections: unknown; formulas: unknown; schemaFingerprint: unknown }[];
  };
  onSaved: (id: string) => void;
}

interface SmartsheetRowPolicy {
  autoSkipFormulaRows: boolean;
  autoSkipParentRows: boolean;
  excludedRowNumbers: number[];
  excludedKeywords: string[];
}

interface InputProcessingPolicy {
  complexFormattingEnabled: boolean;
}

interface WorkbookOptions {
  inputSheetName: string;
  outputSheetName: string;
  excelNewSheetEnabled: boolean;
  excelNewSheetKeepExistingData: boolean;
}

interface DynamicTargetColumnPolicy {
  enabled: boolean;
  sourceLabel: string;
  sourceNodeId: string;
  nameTemplate: string;
  columnPosition: "start" | "end" | "custom";
  customColumnNumber: number;
}

interface LinkedSmartsheetMeta {
  id: string;
  name: string;
}

// Formula Reference Dropdown Component
function FormulaReferenceDropdown() {
  const [open, setOpen] = useState(false);
  const [category, setCategory] = useState<typeof FORMULA_CATEGORIES[number]>("All");
  const [search, setSearch] = useState("");
  const [selectedFormula, setSelectedFormula] = useState<(typeof EXCEL_FORMULAS)[number] | null>(null);

  const formulas = search ? searchFormulas(search) : getFormulasByCategory(category);

  return (
    <Popover open={open} onOpenChange={setOpen}>
      <PopoverTrigger asChild>
        <Button size="sm" variant="outline" className="gap-1.5">
          <BookOpen className="h-3.5 w-3.5" />
          Formulas
        </Button>
      </PopoverTrigger>
      <PopoverContent className="w-[440px] p-0 bg-white border border-slate-200 shadow-lg rounded-lg" align="end">
        {/* Search */}
        <div className="border-b border-slate-100 p-3">
          <div className="flex items-center gap-2 bg-slate-50 rounded-md px-2 py-1.5">
            <Search className="h-3.5 w-3.5 text-slate-400" />
            <input
              type="text"
              placeholder="Search formulas..."
              className="flex-1 text-xs bg-transparent outline-none text-slate-700 placeholder:text-slate-400"
              value={search}
              onChange={(e) => setSearch(e.target.value)}
            />
            {search && (
              <button onClick={() => setSearch("")} className="rounded p-0.5 hover:bg-slate-200">
                <X className="h-3 w-3 text-slate-400" />
              </button>
            )}
          </div>
        </div>
        {/* Category tabs */}
        <div className="flex border-b border-slate-100 overflow-x-auto">
          {FORMULA_CATEGORIES.map((cat) => (
            <button
              key={cat}
              onClick={() => { setCategory(cat); setSearch(""); }}
              className={`px-3 py-2 text-[11px] font-medium whitespace-nowrap transition-colors ${
                category === cat && !search
                  ? "text-slate-900 border-b-2 border-slate-900"
                  : "text-slate-500 hover:text-slate-700"
              }`}
            >
              {cat}
            </button>
          ))}
        </div>
        {/* Formula list */}
        <div className="max-h-[300px] overflow-y-auto">
          {formulas.length === 0 ? (
            <div className="p-4 text-center text-xs text-slate-400">No formulas found</div>
          ) : (
            <div className="divide-y divide-slate-50">
              {formulas.map((f) => (
                <button
                  key={f.name}
                  type="button"
                  onClick={() => setSelectedFormula(f)}
                  className="block w-full p-3 text-left hover:bg-slate-50 transition-colors"
                >
                  <div className="flex items-center gap-2">
                    <span className="font-semibold text-xs text-slate-900">{f.name}</span>
                    <span className="text-[10px] px-1.5 py-0.5 rounded bg-slate-100 text-slate-500">{f.category}</span>
                  </div>
                  <p className="mt-0.5 text-[11px] text-slate-600">{f.description}</p>
                  <code className="mt-1 block text-[10px] text-slate-400 font-mono">{f.syntax}</code>
                </button>
              ))}
            </div>
          )}
        </div>
        {selectedFormula && (
          <div className="border-t border-slate-100 bg-slate-50 px-3 py-2">
            <p className="text-[10px] font-semibold text-slate-600">Examples for {selectedFormula.name}</p>
            {selectedFormula.examples.slice(0, 3).map((example, idx) => (
              <code key={`${selectedFormula.name}-${idx}`} className="mt-1 block text-[10px] text-slate-500 font-mono">
                {example}
              </code>
            ))}
          </div>
        )}
        {/* Footer */}
        <div className="border-t border-slate-100 bg-white px-3 py-2 text-[10px] text-slate-400">
          {EXCEL_FORMULAS.length} Excel formulas supported
        </div>
      </PopoverContent>
    </Popover>
  );
}

// Generate Excel-style column references: A, B, C... Z, AA, AB, etc.
function getColRef(index: number): string {
  let result = '';
  let n = index;
  while (n >= 0) {
    result = String.fromCharCode(65 + (n % 26)) + result;
    n = Math.floor(n / 26) - 1;
  }
  return result;
}

function getFormulaRef(index: number): string {
  return `F${getColRef(index)}`;
}

export function MapperWorkspace({ initialMapping, onSaved }: MapperWorkspaceProps) {
  const { toast } = useToast();
  const reactFlow = useReactFlow();
  const fileRef = useRef<HTMLInputElement>(null);
  const targetFileRef = useRef<HTMLInputElement>(null);

  const [leftPanelMinimized, setLeftPanelMinimized] = useState(false);
  const [rightPanelMinimized, setRightPanelMinimized] = useState(false);
  const [leftPanelWidth, setLeftPanelWidth] = useState(280);
  const [rightPanelWidth, setRightPanelWidth] = useState(280);

  const [targetMode, setTargetMode] = useState<TargetMode>("excel");
  const [mappingName, setMappingName] = useState(initialMapping?.name ?? "");
  const [autoPush, setAutoPush] = useState<boolean>(initialMapping?.autoPush ?? false);
  const [excelSheet, setExcelSheet] = useState<ExcelSheet | null>(null);
  const [sourceWorkbookSheets, setSourceWorkbookSheets] = useState<ExcelSheet[]>([]);
  const [excelTargetSheet, setExcelTargetSheet] = useState<ExcelSheet | null>(null);
  const [ssSheets, setSsSheets] = useState<SmartsheetSheet[]>([]);
  const [selectedSsSheet, setSelectedSsSheet] = useState<SmartsheetSheet | null>(null);
  const [ssLoading, setSsLoading] = useState(false);
  const [ssConnected, setSsConnected] = useState(false);
  const [ssSearch, setSsSearch] = useState("");
  const [changeSummary, setChangeSummary] = useState("");
  const [saving, setSaving] = useState(false);
  const [preview, setPreview] = useState<PreviewState>({
    open: false,
    title: "",
    columns: [],
    rows: [],
    rowCount: 0,
    loading: false,
    hasHierarchy: false,
    hierarchyRows: [],
  });
  const [previewHierarchyMode, setPreviewHierarchyMode] = useState(true);
  const [hierarchyColumnsExpanded, setHierarchyColumnsExpanded] = useState(true);
  const [rowPolicy, setRowPolicy] = useState<SmartsheetRowPolicy>({
    autoSkipFormulaRows: true,
    autoSkipParentRows: true,
    excludedRowNumbers: [],
    excludedKeywords: ["total", "subtotal", "summary"],
  });
  const [inputProcessing, setInputProcessing] = useState<InputProcessingPolicy>({
    complexFormattingEnabled: false,
  });
  const [workbookOptions, setWorkbookOptions] = useState<WorkbookOptions>({
    inputSheetName: "",
    outputSheetName: "Mapped",
    excelNewSheetEnabled: false,
    excelNewSheetKeepExistingData: false,
  });
  const [dynamicTargetColumn, setDynamicTargetColumn] = useState<DynamicTargetColumnPolicy>({
    enabled: false,
    sourceLabel: "",
    sourceNodeId: "",
    nameTemplate: "Enrollments {{DATE}}",
    columnPosition: "end",
    customColumnNumber: 1,
  });
  const [showSsSearch, setShowSsSearch] = useState(false);
  const [addColumnOpen, setAddColumnOpen] = useState(false);
  const [addColumnSide, setAddColumnSide] = useState<"left" | "right">("left");
  const [addColumnName, setAddColumnName] = useState("");
  const [instantAddTargetColumn, setInstantAddTargetColumn] = useState(true);
  const [showInputSettings, setShowInputSettings] = useState(false);
  const [showOutputSettings, setShowOutputSettings] = useState(false);
  const [showAddColumnDestinationPolicy, setShowAddColumnDestinationPolicy] = useState(false);
  const [rowNumberInput, setRowNumberInput] = useState("");
  const [rowKeywordInput, setRowKeywordInput] = useState("total, subtotal, summary");
  const [selectedPanelTargetColumnId, setSelectedPanelTargetColumnId] = useState<string | null>(null);
  const [nodes, setNodes, onNodesChange] = useNodesState<Node>([]);
  const [edges, setEdges, onEdgesChange] = useEdgesState<Edge>([]);
  const prevEdgesRef = useRef<Edge[]>([]);

  // Clean up formula inputs when edges are deleted
  useEffect(() => {
    const prevEdges = prevEdgesRef.current;
    const deletedEdges = prevEdges.filter(pe => !edges.some(e => e.id === pe.id));
    
    if (deletedEdges.length > 0) {
      deletedEdges.forEach((edge) => {
        const targetNode = nodes.find((n) => n.id === edge.target);
        if (targetNode?.type === "formula") {
          setNodes((nds) =>
            nds.map((n) => {
              if (n.id !== edge.target) return n;
              const leftInputs = (n.data?.leftInputs as { id: string; label: string }[]) || [];
              const rightInputs = (n.data?.rightInputs as { id: string; label: string }[]) || [];
              const newLeft = leftInputs.filter((inp) => inp.id !== edge.source);
              const newRight = rightInputs.filter((inp) => inp.id !== edge.source);
              return { ...n, data: { ...n.data, leftInputs: newLeft, rightInputs: newRight } };
            })
          );
        }
      });
    }
    
    prevEdgesRef.current = edges;
  }, [edges, nodes, setNodes]);

  useEffect(() => {
    setAutoPush(initialMapping?.autoPush ?? false);
  }, [initialMapping?.id, initialMapping?.autoPush]);

  useEffect(() => {
    if (!initialMapping?.versions?.length) return;
    const currentVersion =
      initialMapping.versions.find((v) => v.id === initialMapping.currentVersionId) ??
      initialMapping.versions[initialMapping.versions.length - 1];
    if (!currentVersion?.connections) return;
    const conns = currentVersion.connections as { nodes?: Node[]; edges?: Edge[] };
    if (conns.nodes?.length) setNodes(conns.nodes as Node[]);
    if (conns.edges?.length) {
      setEdges((conns.edges as Edge[]).map((e) => ({ ...e, type: e.type ?? "simple" })));
    }
  }, []); // eslint-disable-line react-hooks/exhaustive-deps

  useEffect(() => {
    if (!initialMapping?.versions?.length) return;
    const currentVersion =
      initialMapping.versions.find((v) => v.id === initialMapping.currentVersionId) ??
      initialMapping.versions[initialMapping.versions.length - 1];
    const conns = (currentVersion?.connections as {
      meta?: {
        smartsheetRowPolicy?: Partial<SmartsheetRowPolicy>;
        inputProcessing?: Partial<InputProcessingPolicy>;
        workbookOptions?: Partial<WorkbookOptions>;
        dynamicTargetColumn?: Partial<DynamicTargetColumnPolicy>;
        targetMode?: TargetMode;
        linkedSmartsheet?: LinkedSmartsheetMeta;
      };
    } | undefined);
    const saved = conns?.meta?.smartsheetRowPolicy;
    const savedInputProcessing = conns?.meta?.inputProcessing;
    const savedWorkbookOptions = conns?.meta?.workbookOptions;
    const savedDynamicTargetColumn = conns?.meta?.dynamicTargetColumn;
    const savedTargetMode = conns?.meta?.targetMode;
    const savedLinkedSmartsheet = conns?.meta?.linkedSmartsheet;
    setInputProcessing({
      complexFormattingEnabled: savedInputProcessing?.complexFormattingEnabled === true,
    });
    setWorkbookOptions({
      inputSheetName: typeof savedWorkbookOptions?.inputSheetName === "string" ? savedWorkbookOptions.inputSheetName : "",
      outputSheetName:
        typeof savedWorkbookOptions?.outputSheetName === "string" && savedWorkbookOptions.outputSheetName.trim() !== ""
          ? savedWorkbookOptions.outputSheetName
          : "Mapped",
      excelNewSheetEnabled: savedWorkbookOptions?.excelNewSheetEnabled === true,
      excelNewSheetKeepExistingData: savedWorkbookOptions?.excelNewSheetKeepExistingData === true,
    });
    setDynamicTargetColumn({
      enabled: savedDynamicTargetColumn?.enabled === true,
      sourceLabel: typeof savedDynamicTargetColumn?.sourceLabel === "string" ? savedDynamicTargetColumn.sourceLabel : "",
      sourceNodeId: typeof savedDynamicTargetColumn?.sourceNodeId === "string" ? savedDynamicTargetColumn.sourceNodeId : "",
      nameTemplate:
        typeof savedDynamicTargetColumn?.nameTemplate === "string" && savedDynamicTargetColumn.nameTemplate.trim() !== ""
          ? savedDynamicTargetColumn.nameTemplate
          : "Enrollments {{DATE}}",
      columnPosition:
        savedDynamicTargetColumn?.columnPosition === "start" ||
        savedDynamicTargetColumn?.columnPosition === "custom" ||
        savedDynamicTargetColumn?.columnPosition === "end"
          ? savedDynamicTargetColumn.columnPosition
          : "end",
      customColumnNumber:
        typeof savedDynamicTargetColumn?.customColumnNumber === "number" && savedDynamicTargetColumn.customColumnNumber > 0
          ? Math.floor(savedDynamicTargetColumn.customColumnNumber)
          : 1,
    });
    if (savedTargetMode === "excel" || savedTargetMode === "smartsheet") {
      setTargetMode(savedTargetMode);
      if (savedTargetMode === "smartsheet") setSsConnected(true);
    } else if (initialMapping?.smartsheetSheetId) {
      setTargetMode("smartsheet");
      setSsConnected(true);
    }
    if (savedLinkedSmartsheet?.id) {
      setSsConnected(true);
      setSelectedSsSheet((prev) => prev ?? {
        id: Number(savedLinkedSmartsheet.id),
        name: savedLinkedSmartsheet.name || "Smartsheet",
        columns: [],
        rowCount: 0,
      });
    }
    const savedExcelSheet = (conns as { meta?: { excelSheet?: unknown } } | undefined)?.meta?.excelSheet;
    if (savedExcelSheet && typeof savedExcelSheet === "object" && "columns" in savedExcelSheet) {
      setExcelSheet((prev) => prev ?? (savedExcelSheet as typeof excelSheet));
    }
    if (!saved) return;
    const normalized: SmartsheetRowPolicy = {
      autoSkipFormulaRows: saved.autoSkipFormulaRows !== false,
      autoSkipParentRows: saved.autoSkipParentRows !== false,
      excludedRowNumbers: Array.isArray(saved.excludedRowNumbers) ? saved.excludedRowNumbers.filter((n) => Number.isFinite(n)) : [],
      excludedKeywords: Array.isArray(saved.excludedKeywords)
        ? saved.excludedKeywords.filter((k) => typeof k === "string" && k.trim() !== "")
        : ["total", "subtotal", "summary"],
    };
    setRowPolicy(normalized);
    setRowNumberInput(normalized.excludedRowNumbers.join(", "));
    setRowKeywordInput(normalized.excludedKeywords.join(", "));
  }, [initialMapping?.currentVersionId, initialMapping?.versions]);

  useEffect(() => {
    if (targetMode !== "smartsheet") return;
    const sheetId = selectedSsSheet?.id || Number(initialMapping?.smartsheetSheetId ?? 0);
    if (!sheetId) return;
    if (selectedSsSheet && selectedSsSheet.columns.length > 0) return;
    void selectSsSheet(String(sheetId));
  }, [targetMode, selectedSsSheet?.id, selectedSsSheet?.columns.length, initialMapping?.smartsheetSheetId]);

  useEffect(() => {
    if (!initialMapping?.versions?.length) return;
    const currentVersion =
      initialMapping.versions.find((v) => v.id === initialMapping.currentVersionId) ??
      initialMapping.versions[initialMapping.versions.length - 1];
    const conns = (currentVersion?.connections as { nodes?: Node[] } | undefined);
    const connNodes = Array.isArray(conns?.nodes) ? (conns?.nodes as Node[]) : [];

    const sourceNodes = connNodes.filter((n) => n.id.startsWith("excel_"));
    if (sourceNodes.length > 0) {
      const sourceColumns = sourceNodes
        .map((n, idx) => {
          const colKey = typeof n.data?.colKey === "string" ? n.data.colKey : `restored_${idx}`;
          return {
            key: colKey,
            header: String(n.data?.label ?? `Column ${idx + 1}`),
            index: Number.isFinite(Number(n.data?.index)) ? Number(n.data?.index) : idx,
            dataType: String(n.data?.dataType ?? "string") as "string" | "number" | "date" | "boolean" | "empty",
            sampleValues: Array.isArray(n.data?.sampleValues) ? (n.data?.sampleValues as (string | number | boolean | null)[]) : [],
          };
        })
        .sort((a, b) => a.index - b.index);
      setExcelSheet((prev) => {
        if (prev) return prev;
        return {
          name: workbookOptions.inputSheetName || "Saved source",
          columns: sourceColumns,
          rowCount: 0,
          data: [],
        };
      });
    }

    const targetNodes = connNodes.filter((n) => n.id.startsWith("ss_"));
    if (targetNodes.length > 0) {
      const targetColumns = targetNodes
        .map((n, idx) => ({
          key: String(n.data?.colId ?? `restored_target_${idx}`),
          header: String(n.data?.label ?? `Target ${idx + 1}`),
          index: Number.isFinite(Number(n.data?.index)) ? Number(n.data?.index) : idx,
          dataType: String(n.data?.colType ?? "string") as "string" | "number" | "date" | "boolean" | "empty",
          sampleValues: [],
        }))
        .sort((a, b) => a.index - b.index);
      if (targetMode === "excel") {
        setExcelTargetSheet((prev) => {
          if (prev) return prev;
          return {
            name: workbookOptions.outputSheetName || "Saved target",
            columns: targetColumns,
            rowCount: 0,
            data: [],
          };
        });
      } else {
        setSelectedSsSheet((prev) => {
          if (prev) return prev;
          return {
            id: Number(initialMapping?.smartsheetSheetId ?? 0),
            name: initialMapping?.smartsheetSheetName ?? "Saved Smartsheet target",
            columns: targetColumns.map((c, i) => ({
              id: Number(c.key) || i + 1,
              title: c.header,
              type: c.dataType,
              index: c.index,
            })),
            rowCount: 0,
          };
        });
      }
    }
  }, [initialMapping, targetMode, workbookOptions.inputSheetName, workbookOptions.outputSheetName]);

  useEffect(() => {
    if (!excelSheet) return;
    setSourceWorkbookSheets((prev) => {
      if (prev.length === 0) return [excelSheet];
      if (prev.some((s) => s.name === excelSheet.name)) return prev;
      return [excelSheet, ...prev];
    });
  }, [excelSheet]);

  const syncSmartsheetColumnUpdate = useCallback(async (
    columnId: number,
    updates: { title?: string; type?: string }
  ) => {
    const sheetId = selectedSsSheet?.id ? String(selectedSsSheet.id) : "";
    if (!sheetId) return;
    const payload: Record<string, unknown> = { columnId };
    if (typeof updates.title === "string" && updates.title.trim() !== "") payload.title = updates.title.trim();
    if (typeof updates.type === "string" && updates.type.trim() !== "") payload.type = updates.type.trim();
    if (Object.keys(payload).length <= 1) return;

    const res = await fetch(`/api/smartsheet/sheets/${sheetId}/columns`, {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error(err?.error || "Failed to update Smartsheet column");
    }
    return res.json().catch(() => null);
  }, [selectedSsSheet?.id]);

  const handleLabelChange = useCallback((nodeId: string, newLabel: string) => {
    let updatedSsColumnId: number | null = null;
    setNodes((nds) =>
      nds.map((n) => {
        const matches = n.id === nodeId || String(n.data?.colKey ?? "") === nodeId || `ss_${String(n.data?.colId ?? "")}` === nodeId;
        if (!matches) return n;
        if (n.type === "ssCol") {
          const parsed = Number(n.data?.colId ?? NaN);
          if (Number.isFinite(parsed) && parsed > 0) updatedSsColumnId = parsed;
        }
        if (n.type === "excelCol") {
          const currentSourceField = String(n.data?.sourceField ?? "").trim();
          const fallbackSourceField = String(n.data?.label ?? "").trim();
          return {
            ...n,
            data: {
              ...n.data,
              sourceField: currentSourceField || fallbackSourceField,
              label: newLabel,
            },
          };
        }
        return { ...n, data: { ...n.data, label: newLabel } };
      })
    );

    setSelectedSsSheet((prev) => {
      if (!prev) return prev;
      const nextColumns = prev.columns.map((col) => {
        const nodeMatch = `ss_${String(col.id)}` === nodeId;
        const colMatch = String(col.id) === nodeId.replace(/^ss_/, "");
        if (!nodeMatch && !colMatch) return col;
        return { ...col, title: newLabel };
      });
      return { ...prev, columns: nextColumns };
    });

    if (updatedSsColumnId !== null) {
      void syncSmartsheetColumnUpdate(updatedSsColumnId, { title: newLabel }).catch((e) => {
        toast({ title: "Rename failed", description: (e as Error).message, variant: "destructive" });
      });
    }
  }, [setNodes, setSelectedSsSheet, syncSmartsheetColumnUpdate, toast]);

  const handleTypeChange = useCallback((nodeId: string, newType: string) => {
    let updatedSsColumnId: number | null = null;
    setNodes((nds) =>
      nds.map((n) => {
        const matches = n.id === nodeId || String(n.data?.colKey ?? "") === nodeId || `ss_${String(n.data?.colId ?? "")}` === nodeId;
        if (!matches) return n;
        if (n.type === "excelCol") return { ...n, data: { ...n.data, dataType: newType } };
        if (n.type === "ssCol") {
          const parsed = Number(n.data?.colId ?? NaN);
          if (Number.isFinite(parsed) && parsed > 0) updatedSsColumnId = parsed;
          return { ...n, data: { ...n.data, colType: newType } };
        }
        return n;
      })
    );

    setSelectedSsSheet((prev) => {
      if (!prev) return prev;
      const nextColumns = prev.columns.map((col) => {
        const nodeMatch = `ss_${String(col.id)}` === nodeId;
        const colMatch = String(col.id) === nodeId.replace(/^ss_/, "");
        if (!nodeMatch && !colMatch) return col;
        return { ...col, type: newType };
      });
      return { ...prev, columns: nextColumns };
    });

    if (updatedSsColumnId !== null) {
      void syncSmartsheetColumnUpdate(updatedSsColumnId, { type: newType }).catch((e) => {
        toast({ title: "Type update failed", description: (e as Error).message, variant: "destructive" });
      });
    }
  }, [setNodes, setSelectedSsSheet, syncSmartsheetColumnUpdate, toast]);

  const handleFormulaChange = useCallback((nodeId: string, newFormula: string) => {
    setNodes((nds) =>
      nds.map((n) => (n.id === nodeId ? { ...n, data: { ...n.data, formula: newFormula } } : n))
    );
  }, [setNodes]);

  const loadSourceSheet = useCallback((sheet: ExcelSheet) => {
    setExcelSheet(sheet);
    setWorkbookOptions((prev) => ({ ...prev, inputSheetName: sheet.name }));

    const excelNodes: Node[] = sheet.columns.map((col, i) => ({
      id: `excel_${col.key}`,
      type: "excelCol",
      position: { x: 50, y: 80 + i * 65 },
      data: {
        label: col.header,
        sourceField: col.header,
        dataType: col.dataType,
        sampleValues: col.sampleValues,
        colKey: col.key,
        colRef: "L" + getColRef(i),
        index: i,
        nodeId: `excel_${col.key}`,
        onLabelChange: handleLabelChange,
        onTypeChange: handleTypeChange,
      },
    }));
    setNodes((prev) => [...prev.filter((n) => !n.id.startsWith("excel_")), ...excelNodes]);
  }, [setNodes, handleLabelChange, handleTypeChange]);

  const handleDeleteNode = useCallback((nodeId: string) => {
    setSelectedPanelTargetColumnId((prev) => {
      if (!prev) return prev;
      if (nodeId === `ss_${prev}` || nodeId === prev) return null;
      return prev;
    });

    setNodes((nds) => {
      const ssColId = nodeId.startsWith("ss_") ? nodeId.slice(3) : null;
      const toDeleteIds = nds
        .filter((n) => {
          if (n.id === nodeId) return true;
          if (ssColId && n.type === "ssCol" && String(n.data?.colId ?? "") === ssColId) return true;
          return false;
        })
        .map((n) => n.id);

      if (toDeleteIds.length === 0) return nds;

      setEdges((eds) =>
        eds.filter((e) => !toDeleteIds.includes(e.source) && !toDeleteIds.includes(e.target))
      );

      return nds.filter((n) => !toDeleteIds.includes(n.id));
    });

    if (nodeId.startsWith("ss_")) {
      const ssColId = nodeId.slice(3);
      const sheetId = selectedSsSheet?.id ? String(selectedSsSheet.id) : "";
      const ssColIdNum = Number(ssColId);
      if (sheetId && Number.isFinite(ssColIdNum) && ssColIdNum > 0) {
        void (async () => {
          try {
            const res = await fetch(`/api/smartsheet/sheets/${sheetId}/columns?columnId=${encodeURIComponent(ssColId)}`, {
              method: "DELETE",
            });
            if (!res.ok) {
              const data = await res.json().catch(() => ({}));
              throw new Error(data?.error || "Failed to delete Smartsheet column");
            }
          } catch (e) {
            toast({ title: "Delete failed", description: (e as Error).message, variant: "destructive" });
          }
        })();
      }
      setSelectedSsSheet((prev) => {
        if (!prev) return prev;
        return {
          ...prev,
          columns: (prev.columns || []).filter((c) => String(c.id) !== ssColId),
        };
      });
    }
  }, [setNodes, setEdges, setSelectedSsSheet, selectedSsSheet?.id, toast]);

  useEffect(() => {
    const onKeyDown = (e: KeyboardEvent) => {
      if (e.key !== "Delete" && e.key !== "Backspace") return;
      const active = document.activeElement as HTMLElement | null;
      if (active && (active.tagName === "INPUT" || active.tagName === "TEXTAREA" || active.isContentEditable)) {
        return;
      }
      const selectedIds = nodes.filter((n) => n.selected).map((n) => n.id);
      if (selectedIds.length === 0) {
        if (selectedPanelTargetColumnId) {
          e.preventDefault();
          handleDeleteNode(`ss_${selectedPanelTargetColumnId}`);
        }
        return;
      }
      e.preventDefault();
      selectedIds.forEach((id) => handleDeleteNode(id));
    };
    window.addEventListener("keydown", onKeyDown);
    return () => window.removeEventListener("keydown", onKeyDown);
  }, [nodes, handleDeleteNode, selectedPanelTargetColumnId]);

  useEffect(() => {
    if (!selectedPanelTargetColumnId) return;
    const exists = selectedSsSheet?.columns?.some((col) => String(col.id) === selectedPanelTargetColumnId) ?? false;
    if (!exists) setSelectedPanelTargetColumnId(null);
  }, [selectedPanelTargetColumnId, selectedSsSheet?.columns]);

  useEffect(() => {
    setNodes((nds) =>
      nds.map((n) => {
        if (n.type === "excelCol") {
          return { ...n, data: { ...n.data, onLabelChange: handleLabelChange, onTypeChange: handleTypeChange } };
        }
        if (n.type === "ssCol") {
          return { ...n, data: { ...n.data, onLabelChange: handleLabelChange, onTypeChange: handleTypeChange } };
        }
        if (n.type === "formula") {
          return {
            ...n,
            data: {
              ...n.data,
              onFormulaChange: handleFormulaChange,
              onLabelChange: handleLabelChange,
            },
          };
        }
        return n;
      })
    );
  }, [handleFormulaChange, handleLabelChange, handleTypeChange, setNodes]);

  useEffect(() => {
    const formulaNodes = nodes
      .filter((n) => n.type === "formula")
      .slice()
      .sort((a, b) => {
        const aNum = Number(String(a.id).replace(/^formula_/, ""));
        const bNum = Number(String(b.id).replace(/^formula_/, ""));
        if (Number.isFinite(aNum) && Number.isFinite(bNum)) return aNum - bNum;
        return String(a.id).localeCompare(String(b.id));
      });
    const formulaRefById = new Map<string, string>();
    formulaNodes.forEach((node, index) => {
      formulaRefById.set(String(node.id), getFormulaRef(index));
    });

    let changed = false;
    const nextNodes = nodes.map((node) => {
      const sourceRef = String(node.data?.colRef ?? "").trim();
      const normalizedSourceRef = sourceRef || String(node.data?.label ?? "").trim();
      const syncInputs = (inputs: { id: string; label: string; colRef?: string }[]) => {
        let localChanged = false;
        const updatedInputs = inputs.map((input) => {
          const sourceNode = nodes.find((candidate) => candidate.id === input.id);
          const sourceFormulaRef = formulaRefById.get(String(input.id));
          const expectedRef = sourceFormulaRef
            ?? String(sourceNode?.data?.colRef ?? "").trim()
            ?? "";
          if (expectedRef && expectedRef !== String(input.colRef ?? "").trim()) {
            localChanged = true;
            return { ...input, colRef: expectedRef };
          }
          return input;
        });
        return { updatedInputs, localChanged };
      };

      if (node.type === "formula") {
        const expectedRef = formulaRefById.get(String(node.id)) ?? "";
        const leftInputs = (node.data?.leftInputs as { id: string; label: string; colRef?: string }[] | undefined) ?? [];
        const rightInputs = (node.data?.rightInputs as { id: string; label: string; colRef?: string }[] | undefined) ?? [];
        const left = syncInputs(leftInputs);
        const right = syncInputs(rightInputs);
        if (String(node.data?.colRef ?? "") !== expectedRef || left.localChanged || right.localChanged) {
          changed = true;
          return {
            ...node,
            data: {
              ...node.data,
              colRef: expectedRef,
              leftInputs: left.updatedInputs,
              rightInputs: right.updatedInputs,
            },
          };
        }
        return node;
      }

      if (node.type === "excelCol" || node.type === "ssCol") {
        if (!sourceRef && normalizedSourceRef) {
          changed = true;
          return { ...node, data: { ...node.data, colRef: normalizedSourceRef } };
        }
      }

      return node;
    });

    if (changed) {
      setNodes(nextNodes);
    }
  }, [nodes, setNodes]);

  const addManualSourceColumn = useCallback((label: string) => {
    const id = `excel_manual_${Date.now()}`;
    const nextIndex = nodes.filter((n) => n.id.startsWith("excel_")).length;
    const newNode: Node = {
      id,
      type: "excelCol",
      position: { x: 50, y: 80 + nextIndex * 65 },
      data: {
        label,
        sourceField: label,
        dataType: "string",
        sampleValues: [],
        colKey: id,
        colRef: "L" + getColRef(nextIndex),
        synthetic: true,
        defaultValue: "",
        nodeId: id,
        onLabelChange: handleLabelChange,
        onTypeChange: handleTypeChange,
      },
    };
    setNodes((prev) => [...prev, newNode]);
    toast({ title: "Input column added", description: label });
  }, [nodes, setNodes, toast, handleLabelChange, handleTypeChange]);

  const addManualTargetColumn = useCallback((
    label: string,
    opts?: {
      colId?: string | number;
      synthetic?: boolean;
      columnPosition?: "start" | "end" | "custom";
      customColumnNumber?: number;
    }
  ) => {
    const id = opts?.colId !== undefined && opts?.colId !== null
      ? `ss_${String(opts.colId)}`
      : `ss_manual_${Date.now()}`;
    const nextIndex = nodes.filter((n) => n.id.startsWith("ss_")).length;
    const newNode: Node = {
      id,
      type: "ssCol",
      position: { x: 800, y: 80 + nextIndex * 65 },
      data: {
        label,
        colType: "TEXT_NUMBER",
        colId: opts?.colId ?? id,
        colRef: "R" + getColRef(nextIndex),
        synthetic: opts?.synthetic ?? true,
        columnPosition: opts?.columnPosition ?? "end",
        customColumnNumber:
          typeof opts?.customColumnNumber === "number" && opts.customColumnNumber > 0
            ? Math.floor(opts.customColumnNumber)
            : 1,
        nodeId: id,
        onLabelChange: handleLabelChange,
        onTypeChange: handleTypeChange,
      },
    };
    setNodes((prev) => [...prev, newNode]);
    toast({ title: "Output column added", description: label });
  }, [nodes, setNodes, toast, handleLabelChange, handleTypeChange]);

  const openAddColumnDialog = useCallback((side: "left" | "right") => {
    setAddColumnSide(side);
    setAddColumnName(side === "left" ? `Input ${Date.now().toString().slice(-4)}` : `Output ${Date.now().toString().slice(-4)}`);
    setShowAddColumnDestinationPolicy(side === "right" ? dynamicTargetColumn.enabled : false);
    setAddColumnOpen(true);
  }, [dynamicTargetColumn.enabled]);

  const handleAddColumnConfirm = useCallback(async () => {
    const label = addColumnName.trim();
    if (!label) {
      toast({ title: "Column name required", variant: "destructive" });
      return;
    }

    if (addColumnSide === "left") {
      addManualSourceColumn(label);
      setAddColumnOpen(false);
      return;
    }

    if (showAddColumnDestinationPolicy) {
      setDynamicTargetColumn((prev) => ({ ...prev, nameTemplate: label }));
    }

    const existingTargetTitles = targetMode === "smartsheet" && selectedSsSheet
      ? selectedSsSheet.columns.map((c) => c.title)
      : nodes.filter((n) => n.type === "ssCol").map((n) => String(n.data?.label ?? "")).filter(Boolean);
    const resolvedRightLabel = formatDynamicColumnName(label, {
      existingTitles: existingTargetTitles,
      ensureUnique: true,
    });

    if (targetMode === "smartsheet" && instantAddTargetColumn) {
      const sheetId = selectedSsSheet?.id ? String(selectedSsSheet.id) : "";
      if (!sheetId) {
        toast({ title: "Select a Smartsheet first", variant: "destructive" });
        return;
      }

      const index =
        dynamicTargetColumn.columnPosition === "start"
          ? 0
          : dynamicTargetColumn.columnPosition === "custom"
            ? Math.max(0, Math.floor((dynamicTargetColumn.customColumnNumber || 1) - 1))
            : Math.max(0, selectedSsSheet?.columns?.length ?? 0);

      try {
        const res = await fetch(`/api/smartsheet/sheets/${sheetId}/columns`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ title: resolvedRightLabel, type: "TEXT_NUMBER", index }),
        });
        const data = await res.json().catch(() => ({}));
        if (!res.ok) {
          throw new Error(data?.error || "Failed to add column to Smartsheet");
        }

        const createdId = Number(data?.id);
        addManualTargetColumn(String(data?.title ?? resolvedRightLabel), {
          colId: Number.isFinite(createdId) ? createdId : `manual_${Date.now()}`,
          synthetic: false,
          columnPosition: dynamicTargetColumn.columnPosition,
          customColumnNumber: dynamicTargetColumn.customColumnNumber,
        });

        if (Number.isFinite(createdId)) {
          setSelectedSsSheet((prev) => {
            if (!prev) return prev;
            if (prev.columns.some((c) => Number(c.id) === createdId)) return prev;
            return {
              ...prev,
              columns: [
                ...prev.columns,
                {
                  id: createdId,
                  title: String(data?.title ?? resolvedRightLabel),
                  type: String(data?.type ?? "TEXT_NUMBER"),
                  index: Number.isFinite(Number(data?.index)) ? Number(data.index) : prev.columns.length,
                },
              ],
            };
          });
        }

        toast({ title: "Column added to Smartsheet", description: resolvedRightLabel });
      } catch (e) {
        toast({ title: "Add column failed", description: (e as Error).message, variant: "destructive" });
        return;
      }
    } else {
      addManualTargetColumn(resolvedRightLabel, {
        columnPosition: dynamicTargetColumn.columnPosition,
        customColumnNumber: dynamicTargetColumn.customColumnNumber,
      });
    }

    setAddColumnOpen(false);
  }, [addColumnName, addColumnSide, addManualSourceColumn, addManualTargetColumn, dynamicTargetColumn.columnPosition, dynamicTargetColumn.customColumnNumber, instantAddTargetColumn, nodes, selectedSsSheet?.id, showAddColumnDestinationPolicy, targetMode, toast]);

  const onConnect = useCallback(
    (params: Connection) => {
      const sourceNode = nodes.find((n) => n.id === params.source);
      const targetNode = nodes.find((n) => n.id === params.target);
      if (!sourceNode || !targetNode) return;

      const isTargetLeft = params.targetHandle?.startsWith("left-");
      const isTargetRight = params.targetHandle?.startsWith("right-");

      // Handle any source (excelCol or ssCol) connecting to formula node
      if (targetNode.type === "formula" && (isTargetLeft || isTargetRight)) {
        const isLeft = isTargetLeft;
        const sourceLabel = (sourceNode.data?.label as string) || params.source || "input";
        const sourceColRef = String(sourceNode.data?.colRef ?? "").trim();

        setEdges((eds) =>
          addEdge(
            {
              ...params,
              type: "simple",
              markerEnd: { type: MarkerType.ArrowClosed, width: 12, height: 12 },
            },
            eds
          )
        );

        setNodes((nds) =>
          nds.map((n) => {
            if (n.id === params.target) {
              const inputArr = isLeft ? "leftInputs" : "rightInputs";
              const existing = (n.data?.[inputArr] as { id: string; label: string; colRef?: string }[]) || [];
              return {
                ...n,
                data: {
                  ...n.data,
                  [inputArr]: [...existing, { id: params.source!, label: sourceLabel, colRef: sourceColRef }],
                },
              };
            }
            return n;
          })
        );
        toast({ title: "Input added", description: `${sourceColRef ? sourceColRef + " " : ""}${sourceLabel} → Formula` });
        return;
      }

      // Handle right-side column (ssCol) as input to formula
      if ((sourceNode.type === "ssCol" || sourceNode.type === "excelCol") && targetNode.type === "formula") {
        const isTargetHandle = params.targetHandle;
        const isLeft = isTargetHandle?.startsWith("left-");
        const isRight = isTargetHandle?.startsWith("right-");
        
        if (isLeft || isRight) {
          const inputArr = isLeft ? "leftInputs" : "rightInputs";
          const sourceLabel = (sourceNode.data?.label as string) || params.source || "input";
          const sourceColRef = (sourceNode.data?.colRef as string) || "";

          setEdges((eds) =>
            addEdge(
              {
                ...params,
                type: "simple",
                markerEnd: { type: MarkerType.ArrowClosed, width: 12, height: 12 },
              },
              eds
            )
          );

          setNodes((nds) =>
            nds.map((n) => {
              if (n.id === params.target) {
                const existing = (n.data?.[inputArr] as { id: string; label: string; colRef?: string }[]) || [];
                return {
                  ...n,
                  data: {
                    ...n.data,
                    [inputArr]: [...existing, { id: params.source!, label: sourceLabel, colRef: sourceColRef }],
                  },
                };
              }
              return n;
            })
          );
          toast({ title: "Input added", description: `${sourceColRef ? sourceColRef + " " : ""}${sourceLabel} → Formula` });
          return;
        }
      }

      if (sourceNode.type === "formula" && (params.sourceHandle === "output" || params.sourceHandle === "output-right") && targetNode.type === "ssCol") {
        setEdges((eds) =>
          addEdge(
            {
              ...params,
              type: "simple",
              animated: true,
              markerEnd: { type: MarkerType.ArrowClosed, width: 12, height: 12 },
            },
            eds
          )
        );
        toast({ title: "Formula → Target", description: targetNode.data.label as string });
        return;
      }

      if (sourceNode.type === "excelCol" && targetNode.type === "ssCol") {
        setEdges((eds) =>
          addEdge(
            {
              ...params,
              type: "simple",
              markerEnd: { type: MarkerType.ArrowClosed, width: 12, height: 12 },
            },
            eds
          )
        );
        toast({ title: "Mapped", description: `${sourceNode.data.label} → ${targetNode.data.label}` });
        return;
      }

      setEdges((eds) =>
        addEdge(
          {
            ...params,
            type: "simple",
            markerEnd: { type: MarkerType.ArrowClosed, width: 12, height: 12 },
          },
          eds
        )
      );
    },
    [setEdges, setNodes, nodes, toast]
  );

  const handleEdgesDelete = useCallback((deletedEdges: Edge[]) => {
    deletedEdges.forEach((edge) => {
      const targetNode = nodes.find((n) => n.id === edge.target);
      if (targetNode?.type === "formula") {
        setNodes((nds) =>
          nds.map((n) => {
            if (n.id !== edge.target) return n;
            const leftInputs = (n.data?.leftInputs as { id: string; label: string }[]) || [];
            const rightInputs = (n.data?.rightInputs as { id: string; label: string }[]) || [];
            // Remove input that matches the disconnected edge source
            const newLeft = leftInputs.filter((inp) => inp.id !== edge.source);
            const newRight = rightInputs.filter((inp) => inp.id !== edge.source);
            return { ...n, data: { ...n.data, leftInputs: newLeft, rightInputs: newRight } };
          })
        );
      }
    });
  }, [nodes, setNodes]);

  const createFormulaNode = useCallback(() => {
    const id = `formula_${Date.now()}`;
    const position = reactFlow.screenToFlowPosition({
      x: window.innerWidth / 2 - 130,
      y: window.innerHeight / 2 - 100,
    });

    const newNode: Node = {
      id,
      type: "formula",
      position,
      data: {
        label: "Formula",
        formula: "",
        leftInputs: [],
        rightInputs: [],
        onFormulaChange: handleFormulaChange,
        onLabelChange: handleLabelChange,
      },
    };

    setNodes((nds) => [...nds, newNode]);
    toast({ title: "Formula node created", description: "Connect columns and write your formula" });
  }, [reactFlow, setNodes, toast, handleFormulaChange, handleLabelChange]);

  async function handleExcelUpload(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    if (!file) return;
    const buf = await file.arrayBuffer();
    const parsed = inputProcessing.complexFormattingEnabled
      ? parseExcelFileComplex(buf, file.name, file.size)
      : parseExcelFile(buf, file.name, file.size);
    const sheet =
      parsed.sheets.find((s) => s.name === workbookOptions.inputSheetName) ??
      parsed.sheets[0];
    if (!sheet) { toast({ title: "No data found", variant: "destructive" }); return; }
    setSourceWorkbookSheets(parsed.sheets);
    loadSourceSheet(sheet);
    toast({ title: `Source loaded: ${file.name}`, description: `${sheet.rowCount} rows · ${sheet.columns.length} columns` });
    e.target.value = "";
  }

  async function handleTargetExcelUpload(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    if (!file) return;
    const buf = await file.arrayBuffer();
    const parsed = inputProcessing.complexFormattingEnabled
      ? parseExcelFileComplex(buf, file.name, file.size)
      : parseExcelFile(buf, file.name, file.size);
    const sheet = parsed.sheets[0];
    if (!sheet) { toast({ title: "No data found in target file", variant: "destructive" }); return; }
    setExcelTargetSheet(sheet);
    setWorkbookOptions((prev) => ({ ...prev, outputSheetName: sheet.name }));

    const targetNodes: Node[] = sheet.columns.map((col, i) => ({
      id: `ss_${col.key}`,
      type: "ssCol",
      position: { x: 800, y: 80 + i * 65 },
      data: { label: col.header, colType: col.dataType, colId: col.key, colRef: "R" + getColRef(i), index: i, nodeId: `ss_${col.key}`, onLabelChange: handleLabelChange, onTypeChange: handleTypeChange },
    }));
    setNodes((prev) => [...prev.filter((n) => !n.id.startsWith("ss_")), ...targetNodes]);
    toast({ title: `Target loaded: ${file.name}`, description: `${sheet.rowCount} rows · ${sheet.columns.length} columns` });
    e.target.value = "";
  }

  async function loadSmartsheetSheets() {
    setSsLoading(true);
    try {
      const res = await fetch("/api/smartsheet/sheets");
      if (!res.ok) {
        const err = await res.json();
        throw new Error(err.error || "Failed");
      }
      const sheets = await res.json();
      setSsSheets(sheets);
      setSsConnected(true);
    } catch (e) {
      toast({ title: "Smartsheet error", description: (e as Error).message, variant: "destructive" });
    } finally {
      setSsLoading(false);
    }
  }

  async function refreshSsSheet() {
    if (!selectedSsSheet?.id) return;
    const sheetId = String(selectedSsSheet.id);
    setSsLoading(true);
    try {
      const res = await fetch(`/api/smartsheet/sheets/${sheetId}`);
      const sheet = await res.json();
      setSelectedSsSheet(sheet);
      const ssNodes: Node[] = (sheet.columns || []).map((col: SmartsheetColumn, i: number) => ({
        id: `ss_${col.id}`,
        type: "ssCol",
        position: { x: 800, y: 80 + i * 65 },
        data: { label: col.title, colType: col.type, colId: col.id, colRef: "R" + getColRef(i), index: i, nodeId: `ss_${col.id}`, onLabelChange: handleLabelChange, onTypeChange: handleTypeChange, onDelete: handleDeleteNode },
      }));
      setNodes((prev) => [...prev.filter((n) => !n.id.startsWith("ss_")), ...ssNodes]);
      toast({ title: "Smartsheet refreshed", description: `${sheet.columns?.length ?? 0} columns loaded` });
    } catch {
      toast({ title: "Refresh failed", variant: "destructive" });
    } finally {
      setSsLoading(false);
    }
  }

  async function selectSsSheet(sheetId: string) {
    const basic = ssSheets.find((s) => String(s.id) === sheetId);
    if (!basic) {
      setSsLoading(true);
      try {
        const res = await fetch(`/api/smartsheet/sheets/${sheetId}`);
        const sheet = await res.json();
        setSelectedSsSheet(sheet);
        setSsConnected(true);
        const ssNodes: Node[] = (sheet.columns || []).map((col: SmartsheetColumn, i: number) => ({
          id: `ss_${col.id}`,
          type: "ssCol",
          position: { x: 800, y: 80 + i * 65 },
          data: { label: col.title, colType: col.type, colId: col.id, colRef: "R" + getColRef(i), index: i, nodeId: `ss_${col.id}`, onLabelChange: handleLabelChange, onTypeChange: handleTypeChange, onDelete: handleDeleteNode },
        }));
        setNodes((prev) => [...prev.filter((n) => !n.id.startsWith("ss_")), ...ssNodes]);
      } catch {
        toast({ title: "Failed to load sheet", variant: "destructive" });
      } finally {
        setSsLoading(false);
      }
      return;
    }
    setSsLoading(true);
    try {
      const res = await fetch(`/api/smartsheet/sheets/${sheetId}`);
      const sheet = await res.json();
      setSelectedSsSheet(sheet);
      setSsConnected(true);

      const ssNodes: Node[] = (sheet.columns || []).map((col: SmartsheetColumn, i: number) => ({
        id: `ss_${col.id}`,
        type: "ssCol",
        position: { x: 800, y: 80 + i * 65 },
        data: { label: col.title, colType: col.type, colId: col.id, colRef: "R" + getColRef(i), index: i, nodeId: `ss_${col.id}`, onLabelChange: handleLabelChange, onTypeChange: handleTypeChange, onDelete: handleDeleteNode },
      }));
      setNodes((prev) => [...prev.filter((n) => !n.id.startsWith("ss_")), ...ssNodes]);
    } catch {
      toast({ title: "Failed to load sheet", variant: "destructive" });
    } finally {
      setSsLoading(false);
    }
  }

  function openExcelPreview(sheet: ExcelSheet, title: string) {
    const columns = sheet.columns.map((c) => c.header);
    const rows = (sheet.data as Record<string, unknown>[]).slice(0, 200);
    setPreview({
      open: true,
      title,
      columns,
      rows,
      rowCount: sheet.rowCount,
      loading: false,
      hasHierarchy: false,
      hierarchyRows: [],
    });
  }

  async function openSmartsheetPreview(sheetId: string, sheetName: string) {
    setPreview({
      open: true,
      title: `${sheetName} (Smartsheet)`,
      columns: [],
      rows: [],
      rowCount: 0,
      loading: true,
      hasHierarchy: false,
      hierarchyRows: [],
    });

    try {
      const res = await fetch(`/api/smartsheet/sheets/${sheetId}/data`);
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || "Failed to load Smartsheet preview");
      const sheet = data.sheet as ExcelSheet;
      const columns = sheet.columns.map((c) => c.header);
      const rows = (sheet.data as Record<string, unknown>[]).slice(0, 200);
      const hierarchyRows = Array.isArray(data.hierarchyRows)
        ? (data.hierarchyRows as PreviewState["hierarchyRows"]).slice(0, 200)
        : [];
      setPreview({
        open: true,
        title: `${sheet.name} (Smartsheet)`,
        columns,
        rows,
        rowCount: sheet.rowCount,
        loading: false,
        hasHierarchy: Boolean(data.hasHierarchy),
        hierarchyRows,
      });
      setPreviewHierarchyMode(Boolean(data.hasHierarchy));
      setHierarchyColumnsExpanded(false);
    } catch (e) {
      setPreview((prev) => ({ ...prev, loading: false }));
      toast({ title: "Preview failed", description: (e as Error).message, variant: "destructive" });
    }
  }

  const filteredSsSheets = ssSheets.filter((s) =>
    s.name.toLowerCase().includes(ssSearch.trim().toLowerCase())
  );
  const sourceColumnOptions = nodes
    .filter((n) => n.type === "excelCol" || n.type === "ssCol" || n.type === "formula")
    .map((n) => ({
      id: n.id,
      label: String(n.data?.label ?? n.id),
      type: n.type === "formula" ? "formula" : n.type === "ssCol" ? "target" : "source",
    }))
    .filter((option, index, arr) => arr.findIndex((x) => x.id === option.id) === index);

  async function handleSave() {
    if (!mappingName.trim()) { toast({ title: "Enter a mapping name", variant: "destructive" }); return; }
    setSaving(true);

    const connections = {
      nodes: nodes.map((n) => ({ id: n.id, type: n.type, position: n.position, data: n.data })),
      edges: edges.map((e) => ({
        id: e.id,
        source: e.source,
        target: e.target,
        sourceHandle: e.sourceHandle,
        targetHandle: e.targetHandle,
        type: e.type,
        animated: e.animated,
        markerEnd: e.markerEnd,
        data: e.data,
      })),
      meta: {
        smartsheetRowPolicy: rowPolicy,
        inputProcessing,
        workbookOptions,
        dynamicTargetColumn,
        targetMode,
        linkedSmartsheet: selectedSsSheet
          ? { id: String(selectedSsSheet.id), name: selectedSsSheet.name }
          : undefined,
        excelSheet: excelSheet ?? undefined,
      },
    };
    const formulas: Record<string, string> = {};
    nodes.forEach((n) => { if (n.type === "formula" && n.data?.formula) formulas[n.id] = n.data.formula as string; });
    const schemaFingerprint = excelSheet ? generateSchemaFingerprint(excelSheet) : {};

    try {
      const url = initialMapping ? `/api/mappings/${initialMapping.id}` : "/api/mappings";
      const method = initialMapping ? "PUT" : "POST";
      const res = await fetch(url, {
        method,
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          name: mappingName,
          smartsheetSheetId: selectedSsSheet ? String(selectedSsSheet.id) : initialMapping?.smartsheetSheetId,
          smartsheetSheetName: selectedSsSheet?.name ?? initialMapping?.smartsheetSheetName,
          autoPush,
          connections,
          formulas,
          schemaFingerprint,
          changeSummary: changeSummary || (initialMapping ? `Updated mapping` : "Initial version"),
        }),
      });
      if (!res.ok) {
        let errMsg = "Save failed";
        try { errMsg = (await res.json()).error || errMsg; } catch { errMsg = `HTTP ${res.status}`; }
        throw new Error(errMsg);
      }
      const saved = await res.json();
      toast({ title: "Mapping saved!" });
      onSaved(saved.id);
    } catch (e) {
      toast({ title: "Save failed", description: (e as Error).message, variant: "destructive" });
    } finally {
      setSaving(false);
    }
  }

  return (
    <NodeActionsContext.Provider value={{ deleteNode: handleDeleteNode }}>
    <div className="flex h-[calc(100vh-64px)] flex-col bg-slate-50">
      {/* Top bar */}
      <div className="flex items-center gap-3 border-b border-slate-200 bg-white px-5 py-3 shadow-sm">
        <Input
          className="h-8 w-52 text-sm font-medium"
          placeholder="Mapping name…"
          value={mappingName}
          onChange={(e) => setMappingName(e.target.value)}
        />
        {initialMapping && (
          <Input
            className="h-8 w-48 text-sm"
            placeholder="Change summary"
            value={changeSummary}
            onChange={(e) => setChangeSummary(e.target.value)}
          />
        )}
        <div className="flex items-center gap-2 rounded-md border border-slate-200 px-2 py-1">
          <span className="text-xs font-medium text-slate-600">Auto Push</span>
          <Switch checked={autoPush} onCheckedChange={setAutoPush} />
        </div>
        <div className="ml-auto flex items-center gap-2">
          <FormulaReferenceDropdown />
          <Button
            size="sm"
            variant="outline"
            onClick={() => openAddColumnDialog("right")}
            className="gap-1.5"
          >
            <Plus className="h-3.5 w-3.5" />
            Column
          </Button>
          <Button
            size="sm"
            variant="outline"
            onClick={createFormulaNode}
            className="gap-1.5"
          >
            <Plus className="h-3.5 w-3.5" />
            New Formula Node
          </Button>
          <Button size="sm" onClick={handleSave} disabled={saving}>
            {saving ? <Loader2 className="mr-1.5 h-3.5 w-3.5 animate-spin" /> : <Save className="mr-1.5 h-3.5 w-3.5" />}
            {initialMapping ? "Save new version" : "Save mapping"}
          </Button>
        </div>
      </div>

      <div className="flex flex-1 overflow-hidden">
        {/* Left panel - resize on right border */}
        <div
          className={cn(
            "relative flex shrink-0 flex-col border-r border-slate-200 bg-white transition-all duration-200",
            leftPanelMinimized ? "w-10" : ""
          )}
          style={{ width: leftPanelMinimized ? undefined : leftPanelWidth }}
        >
          <div className="flex items-center justify-between border-b border-slate-200 px-3 py-2">
            {!leftPanelMinimized && (
              <p className="text-xs font-semibold uppercase tracking-wider text-slate-500">Excel Source</p>
            )}
            <div className="ml-auto flex items-center gap-1">
              {!leftPanelMinimized && (
                <button
                  onClick={() => setLeftPanelWidth((w) => Math.max(200, w - 40))}
                  className="rounded p-1 hover:bg-slate-100"
                  title="Shrink"
                >
                  <PanelLeft className="h-3.5 w-3.5 text-slate-400" />
                </button>
              )}
              <button
                onClick={() => setLeftPanelMinimized((m) => !m)}
                className="rounded p-1 hover:bg-slate-100"
                title={leftPanelMinimized ? "Expand" : "Minimize"}
              >
                {leftPanelMinimized ? (
                  <Maximize2 className="h-3.5 w-3.5 text-slate-500" />
                ) : (
                  <Minimize2 className="h-3.5 w-3.5 text-slate-400" />
                )}
              </button>
            </div>
          </div>

          {!leftPanelMinimized && (
            <div className="flex flex-1 flex-col gap-2 overflow-auto p-3">
              <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={handleExcelUpload} />
              <div className="flex gap-2">
                <Button size="sm" variant="outline" className="flex-1" onClick={() => fileRef.current?.click()}>
                  <Upload className="mr-1.5 h-3.5 w-3.5" />
                  {excelSheet ? "Replace file" : "Upload Excel"}
                </Button>
                {excelSheet && (
                  <Button
                    size="sm"
                    variant="outline"
                    onClick={() => openExcelPreview(excelSheet, `${excelSheet.name} (Excel source)`)}
                  >
                    <Eye className="mr-1.5 h-3.5 w-3.5" />
                    View
                  </Button>
                )}
              </div>
              <div className="flex items-center justify-between rounded-md border border-slate-200 bg-slate-50 px-2 py-1.5">
                <span className="text-[11px] text-slate-600">Complex formatting</span>
                <Switch
                  checked={inputProcessing.complexFormattingEnabled}
                  onCheckedChange={(checked) =>
                    setInputProcessing({
                      complexFormattingEnabled: checked,
                    })
                  }
                />
              </div>
              <div className="flex items-center justify-between rounded-md border border-slate-200 bg-slate-50 px-2 py-1.5">
                <span className="text-[11px] text-slate-600">Input settings</span>
                <Switch checked={showInputSettings} onCheckedChange={setShowInputSettings} />
              </div>
              {showInputSettings && (
                <div className="space-y-1 pl-1">
                  <p className="text-[11px] text-slate-500">Choose source sheet from workbook.</p>
                  <select
                    className="h-7 w-full rounded border border-slate-200 bg-white px-2 text-[11px] text-slate-700 outline-none"
                    value={excelSheet?.name || workbookOptions.inputSheetName || ""}
                    onChange={(e) => {
                      const selected = sourceWorkbookSheets.find((s) => s.name === e.target.value);
                      if (!selected) return;
                      loadSourceSheet(selected);
                    }}
                    disabled={sourceWorkbookSheets.length === 0}
                  >
                    {sourceWorkbookSheets.length === 0 ? (
                      <option value="">Upload workbook first</option>
                    ) : (
                      sourceWorkbookSheets.map((sheet) => (
                        <option key={sheet.name} value={sheet.name}>
                          {sheet.name}
                        </option>
                      ))
                    )}
                  </select>
                </div>
              )}
              {excelSheet && (
                <div className="mt-2 space-y-1">
                  <p className="text-xs font-medium text-slate-600">{excelSheet.name} · {excelSheet.rowCount} rows</p>
                  {excelSheet.columns.map((col, i) => (
                    <div
                      key={col.key}
                      className="group flex items-center justify-between rounded border border-slate-200 bg-slate-50 px-2 py-1.5 hover:border-indigo-300 hover:bg-indigo-50/50 cursor-grab active:cursor-grabbing"
                    >
                      <span className="truncate text-xs font-medium text-slate-700">{col.header}</span>
                      <div className="flex items-center gap-1">
                        <code className="hidden text-[9px] font-mono font-bold text-indigo-500 group-hover:block">{"L" + getColRef(i)}</code>
                        <Badge variant="outline" className="text-[10px] border-slate-200">{col.dataType}</Badge>
                      </div>
                    </div>
                  ))}
                </div>
              )}
              {!excelSheet && (
                <p className="mt-4 text-center text-xs text-slate-400">Upload an Excel file to see columns</p>
              )}
            </div>
          )}

          {/* Invisible resize handle on right edge */}
          {!leftPanelMinimized && (
            <div
              className="absolute -right-1 top-0 z-10 h-full w-2 cursor-ew-resize"
              onMouseDown={(e) => {
                const startX = e.clientX;
                const startWidth = leftPanelWidth;
                const handleMouseMove = (ev: MouseEvent) => {
                  const newWidth = Math.max(180, Math.min(450, startWidth + ev.clientX - startX));
                  setLeftPanelWidth(newWidth);
                };
                const handleMouseUp = () => {
                  document.removeEventListener("mousemove", handleMouseMove);
                  document.removeEventListener("mouseup", handleMouseUp);
                };
                document.addEventListener("mousemove", handleMouseMove);
                document.addEventListener("mouseup", handleMouseUp);
              }}
            />
          )}
        </div>

        {/* Canvas */}
        <div className="flex-1 bg-slate-50">
          <ReactFlow
            nodes={nodes}
            edges={edges}
            onNodesChange={onNodesChange}
            onEdgesChange={onEdgesChange}
            onEdgesDelete={handleEdgesDelete}
            onNodesDelete={(deleted) => {
              deleted.forEach((node) => {
                handleDeleteNode(String(node.id));
              });
            }}
            onConnect={onConnect}
            nodeTypes={nodeTypes}
            edgeTypes={edgeTypes}
            deleteKeyCode={["Backspace", "Delete"]}
            fitView
            className="bg-slate-50"
            proOptions={{ hideAttribution: true }}
          >
            <Background variant={BackgroundVariant.Dots} gap={20} size={1} color="#cbd5e1" />
            <Controls 
              className="!bg-white !shadow-md !border-0 [&>button]:!border-0 [&>button]:!bg-white [&>button]:!shadow-none [&>button:hover]:!bg-slate-100" 
              position="bottom-left"
              showInteractive={false}
            />
            <Panel position="top-center">
              
            </Panel>
          </ReactFlow>
        </div>

        {/* Right panel - resize on left border */}
        <div
          className={cn(
            "relative flex shrink-0 flex-col border-l border-slate-200 bg-white transition-all duration-200",
            rightPanelMinimized ? "w-10" : ""
          )}
          style={{ width: rightPanelMinimized ? undefined : rightPanelWidth }}
        >
          <div className="flex items-center justify-between border-b border-slate-200 px-3 py-2">
            {!rightPanelMinimized && (
              <p className="text-xs font-semibold uppercase tracking-wider text-slate-500">Target</p>
            )}
            <div className="ml-auto flex items-center gap-1">
              {!rightPanelMinimized && (
                <button
                  onClick={() => setRightPanelWidth((w) => Math.max(200, w - 40))}
                  className="rounded p-1 hover:bg-slate-100"
                  title="Shrink"
                >
                  <PanelRight className="h-3.5 w-3.5 text-slate-400" />
                </button>
              )}
              <button
                onClick={() => setRightPanelMinimized((m) => !m)}
                className="rounded p-1 hover:bg-slate-100"
                title={rightPanelMinimized ? "Expand" : "Minimize"}
              >
                {rightPanelMinimized ? (
                  <Maximize2 className="h-3.5 w-3.5 text-slate-500" />
                ) : (
                  <Minimize2 className="h-3.5 w-3.5 text-slate-400" />
                )}
              </button>
            </div>
          </div>

          {!rightPanelMinimized && (
            <div className="flex flex-1 flex-col gap-2 overflow-auto p-3">
              <div className="flex gap-1">
                <button
                  onClick={() => setTargetMode("excel")}
                  className={cn(
                    "flex flex-1 items-center justify-center gap-1 rounded-md border px-2 py-1 text-[10px] font-medium transition-colors",
                    targetMode === "excel"
                      ? "border-slate-800 bg-slate-800 text-white"
                      : "border-slate-200 text-slate-600 hover:border-slate-400"
                  )}
                >
                  <Table2 className="h-3 w-3" />Excel
                </button>
                <button
                  onClick={() => setTargetMode("smartsheet")}
                  className={cn(
                    "flex flex-1 items-center justify-center gap-1 rounded-md border px-2 py-1 text-[10px] font-medium transition-colors",
                    targetMode === "smartsheet"
                      ? "border-slate-800 bg-slate-800 text-white"
                      : "border-slate-200 text-slate-600 hover:border-slate-400"
                  )}
                >
                  <Sheet className="h-3 w-3" />Smartsheet
                </button>
              </div>

              {targetMode === "excel" ? (
                <>
                  <input ref={targetFileRef} type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={handleTargetExcelUpload} />
                  <div className="flex gap-2">
                    <Button size="sm" variant="outline" className="flex-1" onClick={() => targetFileRef.current?.click()}>
                      <Upload className="mr-1.5 h-3.5 w-3.5" />
                      {excelTargetSheet ? "Replace target" : "Upload target Excel"}
                    </Button>
                    {excelTargetSheet && (
                      <Button
                        size="sm"
                        variant="outline"
                        onClick={() => openExcelPreview(excelTargetSheet, `${excelTargetSheet.name} (Excel target)`)}
                      >
                        <Eye className="mr-1.5 h-3.5 w-3.5" />
                        View
                      </Button>
                    )}
                  </div>
                  {excelTargetSheet ? (
                    <div className="mt-2 space-y-1">
                      <p className="text-xs font-medium text-slate-600">{excelTargetSheet.name} · {excelTargetSheet.rowCount} rows</p>
                      {excelTargetSheet.columns.map((col, i) => (
                        <div key={col.key} className="flex items-center gap-1.5 rounded border border-slate-200 bg-slate-50 px-2 py-1.5">
                          <ChevronRight className="h-3 w-3 shrink-0 text-slate-400" />
                          <span className="truncate text-xs font-medium text-slate-700">{col.header}</span>
                          <code className="ml-auto text-[9px] font-mono font-bold text-emerald-600">{"R" + getColRef(i)}</code>
                        </div>
                      ))}
                    </div>
                  ) : (
                    <p className="mt-4 text-center text-xs text-slate-400">Upload a target Excel file to map to</p>
                  )}
                  <div className="flex items-center justify-between rounded-md border border-slate-200 bg-slate-50 px-2 py-1.5">
                    <span className="text-[11px] text-slate-600">Output settings</span>
                    <Switch checked={showOutputSettings} onCheckedChange={setShowOutputSettings} />
                  </div>
                  {showOutputSettings && (
                    <div className="space-y-2 pl-1">
                      <p className="text-[11px] text-slate-500">Output sheet name and new-sheet behavior.</p>
                      <Input
                        className="h-7 text-[11px]"
                        value={workbookOptions.outputSheetName}
                        onChange={(e) =>
                          setWorkbookOptions((prev) => ({
                            ...prev,
                            outputSheetName: e.target.value,
                          }))
                        }
                        placeholder="Mapped"
                      />
                      <div className="flex items-center justify-between rounded border border-slate-200 bg-white px-2 py-1.5">
                        <span className="text-[11px] text-slate-600">Write mapped columns to new sheet</span>
                        <Switch
                          checked={workbookOptions.excelNewSheetEnabled}
                          onCheckedChange={(checked) =>
                            setWorkbookOptions((prev) => ({
                              ...prev,
                              excelNewSheetEnabled: checked,
                            }))
                          }
                        />
                      </div>
                      {workbookOptions.excelNewSheetEnabled && (
                        <div className="flex items-center justify-between rounded border border-slate-200 bg-white px-2 py-1.5">
                          <span className="text-[11px] text-slate-600">Keep existing row data in new sheet</span>
                          <Switch
                            checked={workbookOptions.excelNewSheetKeepExistingData}
                            onCheckedChange={(checked) =>
                              setWorkbookOptions((prev) => ({
                                ...prev,
                                excelNewSheetKeepExistingData: checked,
                              }))
                            }
                          />
                        </div>
                      )}
                    </div>
                  )}
                </>
              ) : (
                <>
                  {!ssConnected ? (
                    <Button size="sm" variant="outline" className="w-full" onClick={loadSmartsheetSheets} disabled={ssLoading}>
                      {ssLoading ? <Loader2 className="mr-1.5 h-3.5 w-3.5 animate-spin" /> : null}
                      Connect Smartsheet
                    </Button>
                  ) : (
                    <>
                      {selectedSsSheet && !showSsSearch ? (
                        <div className="flex items-center justify-between rounded-md border border-slate-200 bg-slate-50 px-2 py-1.5">
                          <span className="truncate text-xs font-medium text-slate-700">{selectedSsSheet.name}</span>
                          <button
                            onClick={() => {
                              setShowSsSearch(true);
                              setSsSearch("");
                              void loadSmartsheetSheets();
                            }}
                            className="ml-2 shrink-0 rounded px-1.5 py-0.5 text-[10px] text-slate-500 hover:bg-slate-200"
                          >
                            Replace
                          </button>
                        </div>
                      ) : (
                        <>
                          <div className="flex items-center gap-2 rounded-md border border-slate-200 px-2 py-1.5">
                            <Search className="h-3.5 w-3.5 text-slate-400" />
                            <input
                              className="flex-1 bg-transparent text-xs outline-none placeholder:text-slate-400"
                              placeholder="Search Smartsheet by name..."
                              value={ssSearch}
                              onChange={(e) => setSsSearch(e.target.value)}
                              autoFocus={showSsSearch}
                            />
                            {showSsSearch && (
                              <button
                                onClick={() => setShowSsSearch(false)}
                                className="rounded p-0.5 hover:bg-slate-100"
                              >
                                <X className="h-3 w-3 text-slate-400" />
                              </button>
                            )}
                            {!showSsSearch && <span className="text-[10px] text-slate-400">{filteredSsSheets.length}</span>}
                          </div>
                          {ssLoading ? (
                            <div className="flex items-center justify-center py-4">
                              <Loader2 className="h-4 w-4 animate-spin text-slate-400" />
                            </div>
                          ) : (
                          <div className="max-h-44 overflow-auto rounded-md border border-slate-200">
                            {filteredSsSheets.length === 0 ? (
                              <p className="p-3 text-center text-xs text-slate-400">No sheets match your search</p>
                            ) : (
                              filteredSsSheets.map((s) => (
                                <div key={s.id} className="flex items-center gap-2 border-b border-slate-100 px-2 py-1.5 last:border-b-0">
                                  <button
                                    onClick={() => { void selectSsSheet(String(s.id)); setShowSsSearch(false); }}
                                    className="flex-1 truncate rounded px-1 py-1 text-left text-xs font-medium text-slate-700 hover:bg-slate-100"
                                    title={s.name}
                                  >
                                    {s.name}
                                  </button>
                                  <Button
                                    size="sm"
                                    variant="outline"
                                    className="h-6 px-2 text-[10px]"
                                    onClick={() => openSmartsheetPreview(String(s.id), s.name)}
                                  >
                                    View
                                  </Button>
                                </div>
                              ))
                            )}
                          </div>
                          )}
                        </>
                      )}
                      {selectedSsSheet && !showSsSearch && (
                        <div className="mt-2 space-y-1">
                          <div className="flex items-center justify-between">
                            <p className="text-xs font-medium text-slate-600">{selectedSsSheet.name} · {selectedSsSheet.columns.length} columns</p>
                            <div className="flex items-center gap-1">
                              <Button
                                size="sm"
                                variant="outline"
                                className="h-6 px-2 text-[10px]"
                                onClick={() => void refreshSsSheet()}
                                disabled={ssLoading}
                                title="Refresh columns from Smartsheet"
                              >
                                {ssLoading ? <Loader2 className="h-3 w-3 animate-spin" /> : "↺"}
                              </Button>
                              <Button
                                size="sm"
                                variant="outline"
                                className="h-6 px-2 text-[10px]"
                                onClick={() => openSmartsheetPreview(String(selectedSsSheet.id), selectedSsSheet.name)}
                              >
                                <Eye className="mr-1 h-3 w-3" />
                                View data
                              </Button>
                            </div>
                          </div>
                          {selectedSsSheet.columns.map((col) => {
                            const isPanelSelected = selectedPanelTargetColumnId === String(col.id);
                            return (
                            <div
                              key={col.id}
                              tabIndex={0}
                              onClick={() => setSelectedPanelTargetColumnId(String(col.id))}
                              onKeyDown={(e) => {
                                if ((e.key === "Delete" || e.key === "Backspace") && selectedPanelTargetColumnId === String(col.id)) {
                                  e.preventDefault();
                                  handleDeleteNode(`ss_${String(col.id)}`);
                                }
                              }}
                              className={cn(
                                "flex items-center gap-1.5 rounded border px-2 py-1.5 outline-none",
                                isPanelSelected
                                  ? "border-red-300 bg-red-50"
                                  : "border-slate-200 bg-slate-50",
                                "focus:ring-2 focus:ring-red-100"
                              )}
                            >
                              <ChevronRight className="h-3 w-3 shrink-0 text-slate-400" />
                              <span className="truncate text-xs font-medium text-slate-700">{col.title}</span>
                              <button
                                onClick={() => {
                                  const nodeId = `ss_${String(col.id)}`;
                                  handleDeleteNode(nodeId);
                                }}
                                className="ml-auto rounded p-0.5 text-[10px] text-red-500 hover:bg-red-50"
                                title="Delete column"
                              >
                                <X className="h-3 w-3" />
                              </button>
                            </div>
                          )})}
                        </div>
                      )}
                    </>
                  )}
                </>
              )}
            </div>
          )}

          {/* Invisible resize handle on left edge */}
          {!rightPanelMinimized && (
            <div
              className="absolute -left-1 top-0 z-10 h-full w-2 cursor-ew-resize"
              onMouseDown={(e) => {
                const startX = e.clientX;
                const startWidth = rightPanelWidth;
                const handleMouseMove = (ev: MouseEvent) => {
                  const newWidth = Math.max(180, Math.min(450, startWidth - (ev.clientX - startX)));
                  setRightPanelWidth(newWidth);
                };
                const handleMouseUp = () => {
                  document.removeEventListener("mousemove", handleMouseMove);
                  document.removeEventListener("mouseup", handleMouseUp);
                };
                document.addEventListener("mousemove", handleMouseMove);
                document.addEventListener("mouseup", handleMouseUp);
              }}
            />
          )}
        </div>
      </div>

      <Dialog
        open={preview.open}
        onOpenChange={(open) => {
          setPreview((prev) => ({ ...prev, open }));
          if (open) setHierarchyColumnsExpanded(true);
        }}
      >
        <DialogContent className="max-w-6xl max-h-[85vh] flex flex-col">
          <DialogHeader>
            <DialogTitle>{preview.title}</DialogTitle>
          </DialogHeader>
          {preview.loading ? (
            <div className="flex flex-1 items-center justify-center py-12">
              <Loader2 className="h-5 w-5 animate-spin text-slate-400" />
            </div>
          ) : (
            <>
              <div className="flex items-center justify-between">
                <p className="text-xs text-slate-500">Showing {preview.rows.length} of {preview.rowCount} rows</p>
                {preview.hasHierarchy && (
                  <div className="flex items-center gap-2 rounded-md border border-slate-200 px-2 py-1">
                    <span className="text-[11px] font-medium text-slate-600">Hierarchy view</span>
                    <Switch checked={previewHierarchyMode} onCheckedChange={setPreviewHierarchyMode} />
                    {previewHierarchyMode && (
                      <>
                        <span className="ml-2 text-[11px] font-medium text-slate-600">All columns</span>
                        <Switch checked={hierarchyColumnsExpanded} onCheckedChange={setHierarchyColumnsExpanded} />
                      </>
                    )}
                  </div>
                )}
              </div>
              <div className="flex-1 overflow-auto rounded-md border border-slate-200">
                {preview.hasHierarchy && previewHierarchyMode ? (
                  (() => {
                    const visibleColumns = hierarchyColumnsExpanded ? preview.columns : preview.columns.slice(0, 12);
                    const hiddenCount = Math.max(0, preview.columns.length - visibleColumns.length);
                    return (
                  <div className="space-y-4 p-3">
                    {hiddenCount > 0 && (
                      <div className="rounded-md border border-amber-200 bg-amber-50 px-3 py-2 text-[11px] text-amber-800">
                        Showing {visibleColumns.length} of {preview.columns.length} columns in compact mode. Enable <strong>All columns</strong> to inspect full width.
                      </div>
                    )}
                    {Array.from(
                      preview.hierarchyRows.reduce((acc, row) => {
                        const key = row.sectionPath || "Top level";
                        if (!acc.has(key)) acc.set(key, [] as HierarchyPreviewRow[]);
                        acc.get(key)?.push(row);
                        return acc;
                      }, new Map<string, HierarchyPreviewRow[]>())
                    ).map(([section, sectionRows]) => (
                      <div key={section} className="overflow-hidden rounded-md border border-slate-200">
                        <div className="flex items-center justify-between border-b border-slate-200 bg-slate-50 px-3 py-2">
                          <p className="text-xs font-medium text-slate-700">{section}</p>
                          <Badge variant="outline" className="text-[10px]">{sectionRows.length} rows</Badge>
                        </div>
                        <table className="w-full text-xs">
                          <thead className="bg-slate-50/70">
                            <tr className="border-b border-slate-200">
                              <th className="px-3 py-2 text-left font-medium text-slate-600">Path</th>
                              {visibleColumns.map((col) => (
                                <th key={`${section}-${col}`} className="px-3 py-2 text-left font-medium text-slate-600">{col}</th>
                              ))}
                              {hiddenCount > 0 && <th className="px-3 py-2 text-left font-medium text-slate-500">…</th>}
                            </tr>
                          </thead>
                          <tbody>
                            {sectionRows.map((row) => (
                              <tr key={row.rowId} className="border-b border-slate-100 last:border-0">
                                <td className="px-3 py-2 text-[11px] text-slate-700">
                                  <span style={{ paddingLeft: `${row.depth * 12}px` }} className="inline-flex items-center gap-1">
                                    {row.isParent ? <span className="text-[10px] text-slate-500">[parent]</span> : null}
                                    {row.path}
                                  </span>
                                </td>
                                {visibleColumns.map((col) => (
                                  <td key={`${row.rowId}-${col}`} className="px-3 py-2 font-mono text-[11px] text-slate-700">
                                    {String(row.values[col] ?? "")}
                                  </td>
                                ))}
                                {hiddenCount > 0 && <td className="px-3 py-2 text-slate-400">…</td>}
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    ))}
                  </div>
                    );
                  })()
                ) : preview.columns.length > 0 ? (
                  <table className="w-full text-xs">
                    <thead className="sticky top-0 bg-slate-50">
                      <tr className="border-b border-slate-200">
                        {preview.columns.map((col) => (
                          <th key={col} className="px-3 py-2 text-left font-medium text-slate-600">{col}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {preview.rows.map((row, idx) => (
                        <tr key={idx} className="border-b border-slate-100 last:border-0">
                          {preview.columns.map((col) => (
                            <td key={`${idx}-${col}`} className="px-3 py-2 font-mono text-[11px] text-slate-700">
                              {String(row[col] ?? "")}
                            </td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                ) : (
                  <div className="flex h-full items-center justify-center py-12 text-xs text-slate-400">No preview data available</div>
                )}
              </div>
            </>
          )}
        </DialogContent>
      </Dialog>

      <Dialog open={addColumnOpen} onOpenChange={setAddColumnOpen}>
        <DialogContent className="max-h-[85vh] max-w-md overflow-y-auto">
          <DialogHeader>
            <DialogTitle>Add Column</DialogTitle>
          </DialogHeader>
          <div className="space-y-3">
            <div className="space-y-1">
              <p className="text-xs font-medium text-slate-600">Column header</p>
              <Input
                value={addColumnName}
                onChange={(e) => {
                  const next = e.target.value;
                  setAddColumnName(next);
                  if (addColumnSide === "right") {
                    setDynamicTargetColumn((prev) => ({
                      ...prev,
                      nameTemplate: next,
                    }));
                  }
                }}
                placeholder={addColumnSide === "left" ? "Input header" : "Output header"}
              />
            </div>
            {addColumnSide === "right" && (
            <div className="space-y-1">
              {targetMode === "smartsheet" && (
                <div className="flex items-center justify-between rounded border border-slate-200 bg-slate-50 px-2 py-1.5">
                  <span className="text-[11px] text-slate-600">Instant add to Smartsheet now</span>
                  <Switch checked={instantAddTargetColumn} onCheckedChange={setInstantAddTargetColumn} />
                </div>
              )}
              <p className="text-xs font-medium text-slate-600">Column position</p>
              <div className="grid grid-cols-3 gap-2">
                <button
                  type="button"
                  onClick={() => setDynamicTargetColumn((prev) => ({ ...prev, columnPosition: "start" }))}
                  className={cn(
                    "rounded border px-2 py-1.5 text-[11px]",
                    dynamicTargetColumn.columnPosition === "start" ? "border-slate-800 bg-slate-800 text-white" : "border-slate-200 text-slate-600"
                  )}
                >
                  Start
                </button>
                <button
                  type="button"
                  onClick={() => setDynamicTargetColumn((prev) => ({ ...prev, columnPosition: "end" }))}
                  className={cn(
                    "rounded border px-2 py-1.5 text-[11px]",
                    dynamicTargetColumn.columnPosition === "end" ? "border-slate-800 bg-slate-800 text-white" : "border-slate-200 text-slate-600"
                  )}
                >
                  End
                </button>
                <button
                  type="button"
                  onClick={() => setDynamicTargetColumn((prev) => ({ ...prev, columnPosition: "custom" }))}
                  className={cn(
                    "rounded border px-2 py-1.5 text-[11px]",
                    dynamicTargetColumn.columnPosition === "custom" ? "border-slate-800 bg-slate-800 text-white" : "border-slate-200 text-slate-600"
                  )}
                >
                  Custom #
                </button>
              </div>
              {dynamicTargetColumn.columnPosition === "custom" && (
                <Input
                  className="h-8 text-[11px]"
                  type="number"
                  min={1}
                  value={String(dynamicTargetColumn.customColumnNumber || 1)}
                  onChange={(e) => {
                    const n = Number(e.target.value);
                    setDynamicTargetColumn((prev) => ({
                      ...prev,
                      customColumnNumber: Number.isFinite(n) && n > 0 ? Math.floor(n) : 1,
                    }));
                  }}
                  placeholder="Column number"
                />
              )}
            </div>
            )}
            <div className="flex justify-end gap-2 pt-1">
              <Button size="sm" variant="outline" onClick={() => setAddColumnOpen(false)}>Cancel</Button>
              <Button size="sm" onClick={() => void handleAddColumnConfirm()}>Add column</Button>
            </div>
          </div>
        </DialogContent>
      </Dialog>
    </div>
    </NodeActionsContext.Provider>
  );
}
