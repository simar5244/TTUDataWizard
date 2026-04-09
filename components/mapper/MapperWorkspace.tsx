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
} from "@xyflow/react";
import "@xyflow/react/dist/style.css";
import { Upload, Save, Loader2, ChevronRight, AlertCircle, Table2, Sheet } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Badge } from "@/components/ui/badge";
import { ExcelColumnNode } from "@/components/mapper/ExcelColumnNode";
import { SmartsheetColumnNode } from "@/components/mapper/SmartsheetColumnNode";
import { FormulaEdge } from "@/components/mapper/FormulaEdge";
import { parseExcelFile, generateSchemaFingerprint, type ExcelSheet } from "@/lib/excel";
import { useToast } from "@/hooks/use-toast";

type TargetMode = "smartsheet" | "excel";

const nodeTypes = { excelCol: ExcelColumnNode, ssCol: SmartsheetColumnNode };
const edgeTypes = { formula: FormulaEdge };

interface SmartsheetColumn { id: number; title: string; type: string; index: number }
interface SmartsheetSheet { id: number; name: string; columns: SmartsheetColumn[]; rowCount: number }

interface MapperWorkspaceProps {
  initialMapping?: {
    id: string;
    name: string;
    smartsheetSheetId: string | null;
    smartsheetSheetName: string | null;
    currentVersionId: string | null;
    versions: { id: string; versionNumber: number; connections: unknown; formulas: unknown; schemaFingerprint: unknown }[];
  };
  onSaved: (id: string) => void;
}

export function MapperWorkspace({ initialMapping, onSaved }: MapperWorkspaceProps) {
  const { toast } = useToast();
  const fileRef = useRef<HTMLInputElement>(null);
  const targetFileRef = useRef<HTMLInputElement>(null);

  const [targetMode, setTargetMode] = useState<TargetMode>("excel");
  const [mappingName, setMappingName] = useState(initialMapping?.name ?? "");
  const [excelSheet, setExcelSheet] = useState<ExcelSheet | null>(null);
  const [excelTargetSheet, setExcelTargetSheet] = useState<ExcelSheet | null>(null);
  const [ssSheets, setSsSheets] = useState<SmartsheetSheet[]>([]);
  const [selectedSsSheet, setSelectedSsSheet] = useState<SmartsheetSheet | null>(null);
  const [ssLoading, setSsLoading] = useState(false);
  const [ssConnected, setSsConnected] = useState(false);
  const [changeSummary, setChangeSummary] = useState("");
  const [saving, setSaving] = useState(false);
  const [nodes, setNodes, onNodesChange] = useNodesState<Node>([]);
  const [edges, setEdges, onEdgesChange] = useEdgesState<Edge>([]);

  useEffect(() => {
    if (!initialMapping?.versions?.length) return;
    const currentVersion =
      initialMapping.versions.find((v) => v.id === initialMapping.currentVersionId) ??
      initialMapping.versions[initialMapping.versions.length - 1];
    if (!currentVersion?.connections) return;
    const conns = currentVersion.connections as { nodes?: Node[]; edges?: Edge[] };
    if (conns.nodes?.length) setNodes(conns.nodes as Node[]);
    if (conns.edges?.length) {
      setEdges((conns.edges as Edge[]).map((e) => ({ ...e, type: e.type ?? "formula" })));
    }
  }, []); // eslint-disable-line react-hooks/exhaustive-deps

  const onConnect = useCallback(
    (params: Connection) => {
      const sourceNode = nodes.find((n) => n.id === params.source);
      const targetNode = nodes.find((n) => n.id === params.target);
      const sourceLabel = (sourceNode?.data?.label as string) || params.source || "A";
      const targetLabel = (targetNode?.data?.label as string) || params.target || "B";
      setEdges((eds) =>
        addEdge({ ...params, type: "formula", data: { formula: "", sourceLabel, targetLabel, isNew: true } }, eds)
      );
    },
    [setEdges, nodes]
  );

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
      position: { x: 60, y: 60 + i * 70 },
      data: { label: col.header, dataType: col.dataType, sampleValues: col.sampleValues, colKey: col.key },
    }));
    setNodes((prev) => [...prev.filter((n) => !n.id.startsWith("excel_")), ...excelNodes]);
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
      position: { x: 700, y: 60 + i * 70 },
      data: { label: col.header, colType: col.dataType, colId: col.key },
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
        position: { x: 700, y: 60 + i * 70 },
        data: { label: col.title, colType: col.type, colId: col.id },
      }));
      setNodes((prev) => [...prev.filter((n) => !n.id.startsWith("ss_")), ...ssNodes]);
    } catch {
      toast({ title: "Failed to load sheet", variant: "destructive" });
    } finally {
      setSsLoading(false);
    }
  }

  async function handleSave() {
    if (!mappingName.trim()) { toast({ title: "Enter a mapping name", variant: "destructive" }); return; }
    setSaving(true);

    const connections = {
      nodes: nodes.map((n) => ({ id: n.id, type: n.type, position: n.position, data: n.data })),
      edges: edges.map((e) => ({ id: e.id, source: e.source, target: e.target, data: e.data })),
    };
    const formulas: Record<string, string> = {};
    edges.forEach((e) => { if (e.data?.formula) formulas[e.id] = e.data.formula as string; });
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
          connections,
          formulas,
          schemaFingerprint,
          changeSummary: changeSummary || (initialMapping ? `Updated mapping` : "Initial version"),
        }),
      });
      if (!res.ok) throw new Error((await res.json()).error);
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
    <div className="flex h-[calc(100vh-64px)] flex-col">
      {/* Top bar */}
      <div className="flex items-center gap-3 border-b border-[#E5E5E5] bg-white px-5 py-3">
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
        <div className="ml-auto flex items-center gap-2">
          <Button size="sm" onClick={handleSave} disabled={saving}>
            {saving ? <Loader2 className="mr-1.5 h-3.5 w-3.5 animate-spin" /> : <Save className="mr-1.5 h-3.5 w-3.5" />}
            {initialMapping ? "Save new version" : "Save mapping"}
          </Button>
        </div>
      </div>

      <div className="flex flex-1 overflow-hidden">
        {/* Left panel — Excel */}
        <div className="flex w-56 shrink-0 flex-col border-r border-[#E5E5E5] bg-white">
          <div className="border-b border-[#E5E5E5] px-4 py-3">
            <p className="text-xs font-semibold uppercase tracking-wider text-[#6B6B6B]">Excel Source</p>
          </div>
          <div className="flex flex-1 flex-col gap-2 overflow-auto p-3">
            <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={handleExcelUpload} />
            <Button size="sm" variant="outline" className="w-full" onClick={() => fileRef.current?.click()}>
              <Upload className="mr-1.5 h-3.5 w-3.5" />
              {excelSheet ? "Replace file" : "Upload Excel"}
            </Button>
            {excelSheet && (
              <div className="mt-2 space-y-1">
                <p className="text-xs font-medium text-[#6B6B6B]">{excelSheet.name} · {excelSheet.rowCount} rows</p>
                {excelSheet.columns.map((col) => (
                  <div key={col.key} className="flex items-center justify-between rounded border border-[#E5E5E5] bg-[#F9F9F9] px-2 py-1">
                    <span className="truncate text-xs font-medium">{col.header}</span>
                    <Badge variant="outline" className="text-[10px]">{col.dataType}</Badge>
                  </div>
                ))}
              </div>
            )}
            {!excelSheet && (
              <p className="mt-4 text-center text-xs text-[#A1A1A1]">Upload an Excel file to see columns</p>
            )}
          </div>
        </div>

        {/* Canvas */}
        <div className="flex-1 bg-[#F9F9F9]">
          <ReactFlow
            nodes={nodes}
            edges={edges}
            onNodesChange={onNodesChange}
            onEdgesChange={onEdgesChange}
            onConnect={onConnect}
            nodeTypes={nodeTypes}
            edgeTypes={edgeTypes}
            fitView
            className="bg-[#F9F9F9]"
          >
            <Background variant={BackgroundVariant.Dots} gap={16} size={1} color="#E5E5E5" />
            <Controls />
            <Panel position="top-center">
              {!excelSheet && !selectedSsSheet && !excelTargetSheet && (
                <div className="flex items-center gap-2 rounded-lg border border-[#E5E5E5] bg-white px-4 py-2 text-xs text-[#6B6B6B] shadow-sm">
                  <AlertCircle className="h-3.5 w-3.5" />
                  Upload Source Excel (left) · Upload Target Excel or connect Smartsheet (right) · Draw connections
                </div>
              )}
            </Panel>
          </ReactFlow>
        </div>

        {/* Right panel — Target (Excel or Smartsheet) */}
        <div className="flex w-56 shrink-0 flex-col border-l border-[#E5E5E5] bg-white">
          <div className="border-b border-[#E5E5E5] px-3 py-2">
            <p className="mb-1.5 text-xs font-semibold uppercase tracking-wider text-[#6B6B6B]">Target</p>
            <div className="flex gap-1">
              <button
                onClick={() => setTargetMode("excel")}
                className={`flex flex-1 items-center justify-center gap-1 rounded-md border px-2 py-1 text-[10px] font-medium transition-colors ${
                  targetMode === "excel" ? "border-black bg-black text-white" : "border-[#E5E5E5] text-[#6B6B6B] hover:border-black"
                }`}
              >
                <Table2 className="h-3 w-3" />Excel
              </button>
              <button
                onClick={() => setTargetMode("smartsheet")}
                className={`flex flex-1 items-center justify-center gap-1 rounded-md border px-2 py-1 text-[10px] font-medium transition-colors ${
                  targetMode === "smartsheet" ? "border-black bg-black text-white" : "border-[#E5E5E5] text-[#6B6B6B] hover:border-black"
                }`}
              >
                <Sheet className="h-3 w-3" />Smartsheet
              </button>
            </div>
          </div>

          <div className="flex flex-1 flex-col gap-2 overflow-auto p-3">
            {targetMode === "excel" ? (
              <>
                <input ref={targetFileRef} type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={handleTargetExcelUpload} />
                <Button size="sm" variant="outline" className="w-full" onClick={() => targetFileRef.current?.click()}>
                  <Upload className="mr-1.5 h-3.5 w-3.5" />
                  {excelTargetSheet ? "Replace target" : "Upload target Excel"}
                </Button>
                {excelTargetSheet ? (
                  <div className="mt-2 space-y-1">
                    <p className="text-xs font-medium text-[#6B6B6B]">{excelTargetSheet.name} · {excelTargetSheet.rowCount} rows</p>
                    {excelTargetSheet.columns.map((col) => (
                      <div key={col.key} className="flex items-center gap-1.5 rounded border border-[#E5E5E5] bg-[#F9F9F9] px-2 py-1">
                        <ChevronRight className="h-3 w-3 shrink-0 text-[#A1A1A1]" />
                        <span className="truncate text-xs font-medium">{col.header}</span>
                        <span className="ml-auto text-[9px] text-[#A1A1A1]">{col.dataType}</span>
                      </div>
                    ))}
                  </div>
                ) : (
                  <p className="mt-4 text-center text-xs text-[#A1A1A1]">Upload a target Excel file to map to</p>
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
                    <Select onValueChange={selectSsSheet}>
                      <SelectTrigger className="h-8 text-xs">
                        <SelectValue placeholder="Select sheet…" />
                      </SelectTrigger>
                      <SelectContent>
                        {ssSheets.map((s) => (
                          <SelectItem key={s.id} value={String(s.id)}>{s.name}</SelectItem>
                        ))}
                      </SelectContent>
                    </Select>
                    {selectedSsSheet && (
                      <div className="mt-2 space-y-1">
                        <p className="text-xs font-medium text-[#6B6B6B]">{selectedSsSheet.columns.length} columns</p>
                        {selectedSsSheet.columns.map((col) => (
                          <div key={col.id} className="flex items-center gap-1.5 rounded border border-[#E5E5E5] bg-[#F9F9F9] px-2 py-1">
                            <ChevronRight className="h-3 w-3 shrink-0 text-[#A1A1A1]" />
                            <span className="truncate text-xs font-medium">{col.title}</span>
                          </div>
                        ))}
                      </div>
                    )}
                  </>
                )}
              </>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}
