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
import { EXCEL_FORMULAS, FORMULA_CATEGORIES, getFormulasByCategory, searchFormulas } from "@/lib/excel-formulas";
import { BookOpen, Search, X, Eye } from "lucide-react";
import { Popover, PopoverContent, PopoverTrigger } from "@/components/ui/popover";
import { useToast } from "@/hooks/use-toast";
import { clsx, type ClassValue } from "clsx";
import { twMerge } from "tailwind-merge";

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

// Formula Reference Dropdown Component
function FormulaReferenceDropdown() {
  const [open, setOpen] = useState(false);
  const [category, setCategory] = useState<typeof FORMULA_CATEGORIES[number]>("All");
  const [search, setSearch] = useState("");

  const formulas = search ? searchFormulas(search) : getFormulasByCategory(category);

  return (
    <Popover open={open} onOpenChange={setOpen}>
      <PopoverTrigger asChild>
        <Button size="sm" variant="outline" className="gap-1.5">
          <BookOpen className="h-3.5 w-3.5" />
          Formula List
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
                <div key={f.name} className="p-3 hover:bg-slate-50 transition-colors">
                  <div className="flex items-center gap-2">
                    <span className="font-semibold text-xs text-slate-900">{f.name}</span>
                    <span className="text-[10px] px-1.5 py-0.5 rounded bg-slate-100 text-slate-500">{f.category}</span>
                  </div>
                  <p className="mt-0.5 text-[11px] text-slate-600">{f.description}</p>
                  <code className="mt-1 block text-[10px] text-slate-400 font-mono">{f.syntax}</code>
                </div>
              ))}
            </div>
          )}
        </div>
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
  const [rowNumberInput, setRowNumberInput] = useState("");
  const [rowKeywordInput, setRowKeywordInput] = useState("total, subtotal, summary");
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
    const conns = (currentVersion?.connections as { meta?: { smartsheetRowPolicy?: Partial<SmartsheetRowPolicy> } } | undefined);
    const saved = conns?.meta?.smartsheetRowPolicy;
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

      if (sourceNode.type === "formula" && params.sourceHandle === "output" && targetNode.type === "ssCol") {
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

  const handleFormulaChange = useCallback((nodeId: string, newFormula: string) => {
    setNodes((nds) =>
      nds.map((n) => (n.id === nodeId ? { ...n, data: { ...n.data, formula: newFormula } } : n))
    );
  }, [setNodes]);

  const handleDeleteNode = useCallback((nodeId: string) => {
    setNodes((nds) => nds.filter((n) => n.id !== nodeId));
    setEdges((eds) => eds.filter((e) => e.source !== nodeId && e.target !== nodeId));
  }, [setNodes, setEdges]);

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

  const handleLabelChange = useCallback((nodeId: string, newLabel: string) => {
    setNodes((nds) =>
      nds.map((n) => (n.id === nodeId ? { ...n, data: { ...n.data, label: newLabel } } : n))
    );
  }, [setNodes]);

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
        onDelete: handleDeleteNode,
        onLabelChange: handleLabelChange,
      },
    };

    setNodes((nds) => [...nds, newNode]);
    toast({ title: "Formula node created", description: "Connect columns and write your formula" });
  }, [reactFlow, setNodes, toast, handleFormulaChange, handleDeleteNode, handleLabelChange]);

  async function handleExcelUpload(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    if (!file) return;
    const buf = await file.arrayBuffer();
    const parsed = parseExcelFile(buf, file.name, file.size);
    const sheet = parsed.sheets[0];
    if (!sheet) { toast({ title: "No data found", variant: "destructive" }); return; }
    setExcelSheet(sheet);

    const excelNodes: Node[] = sheet.columns.map((col, i) => ({
      id: `excel_${col.key}`,
      type: "excelCol",
      position: { x: 50, y: 80 + i * 65 },
      data: { label: col.header, dataType: col.dataType, sampleValues: col.sampleValues, colKey: col.key, colRef: "L" + getColRef(i) },
    }));
    setNodes((prev) => [...prev.filter((n) => !n.id.startsWith("excel_")), ...excelNodes]);
    toast({ title: `Source loaded: ${file.name}`, description: `${sheet.rowCount} rows · ${sheet.columns.length} columns` });
  }

  async function handleTargetExcelUpload(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    if (!file) return;
    const buf = await file.arrayBuffer();
    const parsed = parseExcelFile(buf, file.name, file.size);
    const sheet = parsed.sheets[0];
    if (!sheet) { toast({ title: "No data found in target file", variant: "destructive" }); return; }
    setExcelTargetSheet(sheet);

    const targetNodes: Node[] = sheet.columns.map((col, i) => ({
      id: `ss_${col.key}`,
      type: "ssCol",
      position: { x: 800, y: 80 + i * 65 },
      data: { label: col.header, colType: col.dataType, colId: col.key, colRef: "R" + getColRef(i) },
    }));
    setNodes((prev) => [...prev.filter((n) => !n.id.startsWith("ss_")), ...targetNodes]);
    toast({ title: `Target loaded: ${file.name}`, description: `${sheet.rowCount} rows · ${sheet.columns.length} columns` });
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

  async function selectSsSheet(sheetId: string) {
    const basic = ssSheets.find((s) => String(s.id) === sheetId);
    if (!basic) return;
    setSsLoading(true);
    try {
      const res = await fetch(`/api/smartsheet/sheets/${sheetId}`);
      const sheet = await res.json();
      setSelectedSsSheet(sheet);

      const ssNodes: Node[] = (sheet.columns || []).map((col: SmartsheetColumn, i: number) => ({
        id: `ss_${col.id}`,
        type: "ssCol",
        position: { x: 800, y: 80 + i * 65 },
        data: { label: col.title, colType: col.type, colId: col.id, colRef: "R" + getColRef(i) },
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

  async function handleSave() {
    if (!mappingName.trim()) { toast({ title: "Enter a mapping name", variant: "destructive" }); return; }
    setSaving(true);

    const connections = {
      nodes: nodes.map((n) => ({ id: n.id, type: n.type, position: n.position, data: n.data })),
      edges: edges.map((e) => ({ id: e.id, source: e.source, target: e.target, data: e.data })),
      meta: {
        smartsheetRowPolicy: rowPolicy,
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
            className="h-8 w-52 text-sm"
            placeholder="Change summary (optional)"
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
            onConnect={onConnect}
            nodeTypes={nodeTypes}
            edgeTypes={edgeTypes}
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
                      <div className="flex items-center gap-2 rounded-md border border-slate-200 px-2 py-1.5">
                        <Search className="h-3.5 w-3.5 text-slate-400" />
                        <input
                          className="flex-1 bg-transparent text-xs outline-none placeholder:text-slate-400"
                          placeholder="Search Smartsheet by name..."
                          value={ssSearch}
                          onChange={(e) => setSsSearch(e.target.value)}
                        />
                        <span className="text-[10px] text-slate-400">{filteredSsSheets.length}</span>
                      </div>
                      <div className="max-h-44 overflow-auto rounded-md border border-slate-200">
                        {filteredSsSheets.length === 0 ? (
                          <p className="p-3 text-center text-xs text-slate-400">No sheets match your search</p>
                        ) : (
                          filteredSsSheets.map((s) => (
                            <div key={s.id} className="flex items-center gap-2 border-b border-slate-100 px-2 py-1.5 last:border-b-0">
                              <button
                                onClick={() => selectSsSheet(String(s.id))}
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
                      {selectedSsSheet && (
                        <div className="mt-2 space-y-1">
                          <div className="flex items-center justify-between">
                            <p className="text-xs font-medium text-slate-600">{selectedSsSheet.name} · {selectedSsSheet.columns.length} columns</p>
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
                          {selectedSsSheet.columns.map((col) => (
                            <div key={col.id} className="flex items-center gap-1.5 rounded border border-slate-200 bg-slate-50 px-2 py-1.5">
                              <ChevronRight className="h-3 w-3 shrink-0 text-slate-400" />
                              <span className="truncate text-xs font-medium text-slate-700">{col.title}</span>
                            </div>
                          ))}
                          <div className="mt-3 space-y-2 rounded-md border border-slate-200 bg-slate-50 p-2">
                            <p className="text-[11px] font-semibold text-slate-700">Destination row policy</p>
                            <div className="flex items-center justify-between rounded border border-slate-200 bg-white px-2 py-1.5">
                              <span className="text-[11px] text-slate-600">Auto-skip formula/summary rows</span>
                              <Switch
                                checked={rowPolicy.autoSkipFormulaRows}
                                onCheckedChange={(checked) => setRowPolicy((prev) => ({ ...prev, autoSkipFormulaRows: checked }))}
                              />
                            </div>
                            <div className="flex items-center justify-between rounded border border-slate-200 bg-white px-2 py-1.5">
                              <span className="text-[11px] text-slate-600">Auto-skip parent (group header) rows</span>
                              <Switch
                                checked={rowPolicy.autoSkipParentRows}
                                onCheckedChange={(checked) => setRowPolicy((prev) => ({ ...prev, autoSkipParentRows: checked }))}
                              />
                            </div>
                            <div className="rounded border border-slate-200 bg-white p-2">
                              <p className="text-[11px] font-medium text-slate-600">Manual exclude destination row numbers</p>
                              <Input
                                className="mt-1 h-7 text-[11px]"
                                placeholder="8, 18, 32"
                                value={rowNumberInput}
                                onChange={(e) => {
                                  const raw = e.target.value;
                                  setRowNumberInput(raw);
                                  const numbers = raw
                                    .split(",")
                                    .map((s) => Number(s.trim()))
                                    .filter((n) => Number.isFinite(n) && n > 0)
                                    .map((n) => Math.floor(n));
                                  setRowPolicy((prev) => ({ ...prev, excludedRowNumbers: numbers }));
                                }}
                              />
                            </div>
                            <div className="rounded border border-slate-200 bg-white p-2">
                              <p className="text-[11px] font-medium text-slate-600">Manual exclude by row keyword</p>
                              <Input
                                className="mt-1 h-7 text-[11px]"
                                placeholder="total, subtotal, summary"
                                value={rowKeywordInput}
                                onChange={(e) => {
                                  const raw = e.target.value;
                                  setRowKeywordInput(raw);
                                  const keywords = raw
                                    .split(",")
                                    .map((s) => s.trim().toLowerCase())
                                    .filter(Boolean);
                                  setRowPolicy((prev) => ({ ...prev, excludedKeywords: keywords }));
                                }}
                              />
                            </div>
                          </div>
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
    </div>
  );
}
