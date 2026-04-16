"use client";

import { useEffect, useState, type ChangeEvent } from "react";
import { useDropzone } from "react-dropzone";
import { Upload, CheckCircle2, AlertCircle, ArrowRight, Loader2, Download, Table2, Sheet } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Dialog, DialogContent, DialogHeader, DialogTitle } from "@/components/ui/dialog";
import { Switch } from "@/components/ui/switch";
import { Input } from "@/components/ui/input";
import { parseExcelFile, validateExcelAgainstFingerprint, type SchemaFingerprint, type ExcelSheet } from "@/lib/excel";
import { parseExcelFileComplex } from "@/lib/ComplexInput";
import { evaluateFormula, indexToColRef } from "@/lib/formulas";
import { formatDynamicColumnName } from "@/lib/dynamic-column";
import { useToast } from "@/hooks/use-toast";
import type { MappingConnections } from "@/app/mapper/page";
import * as XLSX from "xlsx";

interface MappingVersion {
  id: string;
  versionNumber: number;
  schemaFingerprint?: unknown;
  connections?: MappingConnections & {
    meta?: {
      smartsheetRowPolicy?: {
        autoSkipFormulaRows?: boolean;
        autoSkipParentRows?: boolean;
        excludedRowNumbers?: number[];
        excludedKeywords?: string[];
      };
      inputProcessing?: {
        complexFormattingEnabled?: boolean;
      };
      workbookOptions?: {
        inputSheetName?: string;
        outputSheetName?: string;
        excelNewSheetEnabled?: boolean;
        excelNewSheetKeepExistingData?: boolean;
      };
      dynamicTargetColumn?: {
        enabled?: boolean;
        sourceLabel?: string;
        sourceNodeId?: string;
        nameTemplate?: string;
        columnPosition?: "start" | "end" | "custom";
        customColumnNumber?: number;
      };
      targetMode?: "excel" | "smartsheet";
      detailMappingEnabled?: boolean;
      detailStore?: Record<string, { nodeId: string; ranges: { id: string; start: number; end: number }[] }>;
    };
  };
  formulas?: Record<string, string>;
}

interface RunMappingDialogProps {
  mapping: {
    id: string;
    name: string;
    autoPush?: boolean;
    smartsheetSheetId?: string | null;
    currentVersionId: string | null;
    versions: MappingVersion[];
  };
  open: boolean;
  onClose: () => void;
}

type Step = "upload" | "validate" | "approve" | "running" | "done";
type RunMode = "excel_to_excel" | "excel_to_ss";

interface ChangeCell {
  row: number;
  column: string;
  from: unknown;
  to: unknown;
}

interface ChangePreview {
  changes: ChangeCell[];
  changedRows: number[];
  changedColumns: string[];
  changedCellCount: number;
  changedRowCount: number;
}

interface RunOptions {
  hierarchyAware: boolean;
  protectFormulaCells: boolean;
  protectParentSummaryRows: boolean;
  excludedRowPatterns: string[];
  columnExcludedRowPatterns: Record<string, string[]>;
  excludedRowNumbers: number[];
  skipTopRows: number;
}

function toSheetName(raw: string, fallback: string): string {
  const sanitized = raw.replace(/[\\/*?:\[\]]/g, " ").trim();
  if (!sanitized) return fallback;
  return sanitized.slice(0, 31);
}

export function RunMappingDialog({ mapping, open, onClose }: RunMappingDialogProps) {
  const { toast } = useToast();
  const [step, setStep] = useState<Step>("upload");
  const [file, setFile] = useState<File | null>(null);
  const [parsedSheet, setParsedSheet] = useState<ExcelSheet | null>(null);
  const [validation, setValidation] = useState<ReturnType<typeof validateExcelAgainstFingerprint> | null>(null);
  const [loading, setLoading] = useState(false);
  const [outputPreviewRows, setOutputPreviewRows] = useState<Record<string, unknown>[]>([]);
  const [changePreview, setChangePreview] = useState<ChangePreview | null>(null);
  const [runOptions, setRunOptions] = useState<RunOptions>({
    hierarchyAware: true,
    protectFormulaCells: true,
    protectParentSummaryRows: true,
    excludedRowPatterns: ["total", "subtotal", "summary"],
    columnExcludedRowPatterns: {},
    excludedRowNumbers: [],
    skipTopRows: 0,
  });
  const [excludedPatternInput, setExcludedPatternInput] = useState("total, subtotal, summary");
  const [excludedRowNumberInput, setExcludedRowNumberInput] = useState("");
  const [columnExclusionInput, setColumnExclusionInput] = useState("");
  const [skipTopRowsInput, setSkipTopRowsInput] = useState("");
  const [lastPolicyDiagnostics, setLastPolicyDiagnostics] = useState<{
    totalOutputRows?: number;
    skippedCells?: number;
    skipByReason?: Record<string, number>;
    parentRowsDetected?: number;
    formulaRowsDetected?: number;
  } | null>(null);

  const isExcelToExcel = !mapping.smartsheetSheetId;
  const runMode: RunMode = isExcelToExcel ? "excel_to_excel" : "excel_to_ss";

  const currentVersion = mapping.versions.find((v) => v.id === mapping.currentVersionId) ?? mapping.versions.at(-1);
  const fingerprint = currentVersion?.schemaFingerprint as SchemaFingerprint | undefined;

  const savedRowPolicy = currentVersion?.connections?.meta?.smartsheetRowPolicy;
  const complexFormattingEnabled = currentVersion?.connections?.meta?.inputProcessing?.complexFormattingEnabled === true;
  const workbookOptions = currentVersion?.connections?.meta?.workbookOptions;
  const dynamicTargetColumn = currentVersion?.connections?.meta?.dynamicTargetColumn;

  useEffect(() => {
    const policy = savedRowPolicy;
    if (!policy) return;
    const excludedKeywords = Array.isArray(policy.excludedKeywords)
      ? policy.excludedKeywords.filter((k) => typeof k === "string" && k.trim() !== "")
      : ["total", "subtotal", "summary"];
    const excludedRowNumbers = Array.isArray(policy.excludedRowNumbers)
      ? policy.excludedRowNumbers.filter((n) => Number.isFinite(n) && n > 0).map((n) => Math.floor(n))
      : [];

    setRunOptions((prev) => ({
      ...prev,
      protectFormulaCells: policy.autoSkipFormulaRows !== false,
      protectParentSummaryRows: policy.autoSkipParentRows !== false,
      excludedRowPatterns: excludedKeywords,
      excludedRowNumbers,
    }));
    setExcludedPatternInput(excludedKeywords.join(", "));
    setExcludedRowNumberInput(excludedRowNumbers.join(", "));
  }, [savedRowPolicy]);

  const onDrop = async (acceptedFiles: File[]) => {
    const f = acceptedFiles[0];
    if (!f) return;
    setFile(f);
    setStep("validate");

    const buf = await f.arrayBuffer();
    const parsed = complexFormattingEnabled
      ? parseExcelFileComplex(buf, f.name, f.size)
      : parseExcelFile(buf, f.name, f.size);
    const sheet = parsed.sheets[0];

    if (!sheet) {
      toast({ title: "No data found in Excel file", variant: "destructive" });
      setStep("upload");
      return;
    }

    setParsedSheet(sheet);
    const previewRows = applyMappingToSheet(sheet);
    setOutputPreviewRows(previewRows);
    setChangePreview(buildChangePreview(sheet.data as Record<string, unknown>[], previewRows));

    if (!fingerprint) {
      setValidation({ status: "exact", missingColumns: [], remappedColumns: [], extraColumns: [] });
      setStep("approve");
      if (!isExcelToExcel && mapping.autoPush) {
        setTimeout(() => {
          void handleConfirm();
        }, 0);
      }
      return;
    }

    const result = validateExcelAgainstFingerprint(sheet, fingerprint);
    setValidation(result);
    if (result.status === "blocked") {
      setStep("validate");
    } else {
      setStep("approve");
      if (!isExcelToExcel && mapping.autoPush) {
        setTimeout(() => {
          void handleConfirm();
        }, 0);
      }
    }
  };

  const { getRootProps, getInputProps, isDragActive } = useDropzone({
    onDrop,
    accept: {
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"],
      "application/vnd.ms-excel": [".xls"],
      "text/csv": [".csv"],
    },
    maxFiles: 1,
  });

  function applyMappingToSheet(sheet: ExcelSheet): Record<string, unknown>[] {
    const conns = currentVersion?.connections;
    if (!conns?.nodes?.length) return sheet.data;

    const nodes = conns.nodes ?? [];
    const edges = conns.edges ?? [];
    const targetNodes = nodes.filter((n) => n.type === "ssCol");
    const formulaNodesOrdered = nodes
      .filter((n) => n.type === "formula")
      .slice()
      .sort((a, b) => {
        const aNum = Number(String(a.id).replace(/^formula_/, ""));
        const bNum = Number(String(b.id).replace(/^formula_/, ""));
        if (Number.isFinite(aNum) && Number.isFinite(bNum)) return aNum - bNum;
        return String(a.id).localeCompare(String(b.id));
      });
    const formulaRefById = new Map<string, string>();
    formulaNodesOrdered.forEach((node, index) => {
      formulaRefById.set(String(node.id), `F${indexToColRef(index).toUpperCase()}`);
    });
    const dynamicColEnabled = dynamicTargetColumn?.enabled === true;
    const dynamicSourceLabel = String(dynamicTargetColumn?.sourceLabel ?? "").trim();
    const dynamicSourceNodeId = String(dynamicTargetColumn?.sourceNodeId ?? "").trim();
    const dynamicTargetLabel = dynamicColEnabled
      ? formatDynamicColumnName(String(dynamicTargetColumn?.nameTemplate ?? "Enrollments {{DATE}}"), {
          ensureUnique: false,
        })
      : "";

    function normalizeRuntimeValue(value: unknown): string | number | null {
      if (value === null || value === undefined) return null;
      if (typeof value === "number") return Number.isFinite(value) ? value : null;
      if (typeof value === "boolean") return value ? 1 : 0;
      const textValue = String(value);
      const numeric = Number(textValue);
      return Number.isFinite(numeric) && textValue.trim() !== "" ? numeric : textValue;
    }

    function resolveNodeValue(
      nodeId: string,
      row: Record<string, string | number | boolean | null>,
      stack: Set<string> = new Set()
    ): unknown {
      if (stack.has(nodeId)) return null;
      const node = nodes.find((n) => n.id === nodeId);
      if (!node) return null;
      if (node.type === "excelCol") {
        const sourceField = String(node.data.sourceField ?? "").trim();
        const label = String(node.data.label ?? "").trim();
        if (sourceField && sourceField in row) return row[sourceField] ?? null;
        if (label && label in row) return row[label] ?? null;
        return null;
      }
      if (node.type === "ssCol") {
        const inEdge = edges.find((e) => e.target === nodeId);
        if (inEdge?.source) {
          const nextStack = new Set(stack);
          nextStack.add(nodeId);
          return resolveNodeValue(inEdge.source, row, nextStack);
        }
        const sourceField = String(node.data.sourceField ?? "").trim();
        const label = String(node.data.label ?? "").trim();
        if (sourceField && sourceField in row) return row[sourceField] ?? null;
        if (label && label in row) return row[label] ?? null;
        return null;
      }
      if (node.type === "formula") {
        const formula = (node.data.formula as string) || "";
        if (!formula) return null;
        const runtimeContext: Record<string, string | number | null> = {};
        const nextStack = new Set(stack);
        nextStack.add(nodeId);
        const incomingEdges = edges.filter((edge) => edge.target === nodeId);
        const connectedInputs = [
          ...(((node.data.leftInputs as { id: string; label: string; colRef?: string }[] | undefined) ?? []).map((input, index) => ({
            ...input,
            side: "left" as const,
            handleId: `left-${index}`,
          }))),
          ...(((node.data.rightInputs as { id: string; label: string; colRef?: string }[] | undefined) ?? []).map((input, index) => ({
            ...input,
            side: "right" as const,
            handleId: `right-${index}`,
          }))),
        ];

        connectedInputs.forEach((input, index) => {
          const matchingEdge = incomingEdges.find((edge) => {
            const targetHandle = (edge as { targetHandle?: string }).targetHandle;
            return edge.source === input.id && targetHandle === input.handleId;
          })
            ?? incomingEdges.find((edge) => edge.source === input.id);
          const sourceNodeId = matchingEdge?.source ?? input.id;
          const sourceNode = nodes.find((candidate) => candidate.id === sourceNodeId);
          const resolvedValue = sourceNode
            ? resolveNodeValue(sourceNode.id, row, nextStack)
            : row[input.label] ?? null;
          const normalizedValue = normalizeRuntimeValue(resolvedValue);
          const explicitRef = String(input.colRef ?? sourceNode?.data?.colRef ?? "").trim();
          const sourceFormulaRef = sourceNode ? formulaRefById.get(String(sourceNode.id)) : undefined;
          const aliasRef = indexToColRef(index).toUpperCase();
          if (explicitRef) runtimeContext[explicitRef] = normalizedValue;
          if (sourceFormulaRef) runtimeContext[sourceFormulaRef] = normalizedValue;
          runtimeContext[aliasRef] = normalizedValue;
        });

        if (Object.keys(runtimeContext).length === 0) {
          nodes
            .filter((candidate) => candidate.type === "excelCol" || candidate.type === "ssCol")
            .forEach((candidate) => {
              const ref = String(candidate.data.colRef ?? "").trim();
              const label = String(candidate.data.label ?? "").trim();
              const sourceField = String(candidate.data.sourceField ?? "").trim();
              const fieldName = sourceField || label;
              if (!ref || !fieldName) return;
              runtimeContext[ref] = normalizeRuntimeValue(row[fieldName]);
            });
        }

        const selfFormulaRef = formulaRefById.get(String(node.id));
        if (selfFormulaRef && !(selfFormulaRef in runtimeContext)) {
          runtimeContext[selfFormulaRef] = 0;
        }

        return evaluateFormula(formula, runtimeContext);
      }
      return null;
    }

    const detailEnabled = currentVersion?.connections?.meta?.detailMappingEnabled === true;
    const detailStore = (currentVersion?.connections?.meta?.detailStore ?? {}) as Record<string, { nodeId: string; ranges: { id: string; start: number; end: number }[] }>;

    // Determine whether any edges use range handles (r_xxx ids)
    const isRangeHandleId = (h: string | undefined | null) => typeof h === "string" && h.startsWith("r_");
    const hasRangeEdges = detailEnabled && edges.some(
      (e) => isRangeHandleId(e.sourceHandle) || isRangeHandleId(e.targetHandle)
    );

    if (!hasRangeEdges) {
      // --- Original behaviour: 1-to-1 row mapping ---
      return sheet.data.map((row) => {
        const outRow: Record<string, unknown> = {};
        for (const tgtNode of targetNodes) {
          const tgtLabel = (tgtNode.data.label as string) ?? "";
          if (!tgtLabel) continue;
          const inEdge = edges.find((e) => e.target === tgtNode.id);
          if (!inEdge) continue;
          outRow[tgtLabel] = resolveNodeValue(inEdge.source, row);
        }
        if (dynamicColEnabled && dynamicTargetLabel) {
          let dynamicValue: unknown = null;
          if (dynamicSourceNodeId) {
            dynamicValue = resolveNodeValue(dynamicSourceNodeId, row);
          } else if (dynamicSourceLabel) {
            dynamicValue = row[dynamicSourceLabel] ?? null;
          }
          outRow[dynamicTargetLabel] = dynamicValue;
        }
        return outRow;
      });
    }

    // --- Detailed mapping: range-aware output ---
    // Helper: look up a range object by its id from detailStore
    function findRangeById(rangeId: string): { start: number; end: number } | null {
      for (const entry of Object.values(detailStore)) {
        const r = entry.ranges.find((rr) => rr.id === rangeId);
        if (r) return r;
      }
      return null;
    }

    // DEBUG — remove after diagnosis
    console.log("[DetailMapping] detailEnabled:", detailEnabled);
    console.log("[DetailMapping] detailStore:", JSON.stringify(detailStore, null, 2));
    console.log("[DetailMapping] hasRangeEdges:", hasRangeEdges);
    console.log("[DetailMapping] edges:", JSON.stringify(edges.map((e) => ({ id: e.id, source: e.source, target: e.target, sourceHandle: e.sourceHandle, targetHandle: e.targetHandle })), null, 2));

    // Determine output array size: max of target range ends, fallback to input length
    let outputSize = sheet.data.length;
    for (const edge of edges) {
      if (isRangeHandleId(edge.targetHandle)) {
        const r = findRangeById(edge.targetHandle!);
        if (r && r.end > outputSize) outputSize = r.end;
      }
    }

    // Build sparse output rows indexed 1-based (index 0 = output row 1)
    const outRows: Record<string, unknown>[] = Array.from({ length: outputSize }, () => ({}));

    for (const tgtNode of targetNodes) {
      const tgtLabel = (tgtNode.data.label as string) ?? "";
      if (!tgtLabel) continue;

      // Collect all edges going to this target node
      const inEdges = edges.filter((e) => e.target === tgtNode.id);

      // Walk back through formula chains to find the deepest ranged source edge.
      // This handles: excelCol[range] → formula → ssCol (the formula→ssCol edge
      // itself has no range handle, but its upstream feed edge does).
      const findUpstreamRange = (nodeId: string, directHandle: string | null | undefined): { start: number; end: number } | null => {
        if (isRangeHandleId(directHandle)) return findRangeById(directHandle!);
        const n = nodes.find((x) => x.id === nodeId);
        if (!n || n.type !== "formula") return null;
        for (const fe of edges.filter((fe) => fe.target === nodeId)) {
          if (isRangeHandleId(fe.sourceHandle)) return findRangeById(fe.sourceHandle!);
          const deeper = findUpstreamRange(fe.source, fe.sourceHandle);
          if (deeper) return deeper;
        }
        return null;
      };

      for (const inEdge of inEdges) {
        const srcHandle = inEdge.sourceHandle;
        const tgtHandle = inEdge.targetHandle;

        // Skip orphaned edges: handle id looks like a range id but was deleted from the store
        if (isRangeHandleId(srcHandle) && !findRangeById(srcHandle!)) continue;
        if (isRangeHandleId(tgtHandle) && !findRangeById(tgtHandle!)) continue;

        const tgtIsRange = isRangeHandleId(tgtHandle);
        // Walk the source side — handles direct range handles AND formula chains
        const srcRange = findUpstreamRange(inEdge.source, srcHandle);
        const tgtRange = tgtIsRange ? findRangeById(tgtHandle!) : null;

        const isRangeEdge = srcRange !== null || tgtRange !== null;

        if (isRangeEdge) {
          const srcStart = srcRange ? srcRange.start : 1;
          const srcEnd = srcRange ? srcRange.end : sheet.data.length;
          const srcLen = srcEnd - srcStart + 1;
          const dstStart = tgtRange ? tgtRange.start : 1;

          for (let i = 0; i < srcLen; i++) {
            const inputRowIdx = srcStart - 1 + i;
            const outputRowIdx = dstStart - 1 + i;
            const inputRow = sheet.data[inputRowIdx] as Record<string, string | number | boolean | null> | undefined;
            if (!inputRow) continue;
            if (outputRowIdx >= outRows.length) continue;
            outRows[outputRowIdx][tgtLabel] = resolveNodeValue(inEdge.source, inputRow);
          }
        } else {
          // Fully non-range edge — apply to all rows as normal
          sheet.data.forEach((row, idx) => {
            if (outRows[idx]) {
              outRows[idx][tgtLabel] = resolveNodeValue(inEdge.source, row as Record<string, string | number | boolean | null>);
            }
          });
        }
      }
    }

    if (dynamicColEnabled && dynamicTargetLabel) {
      sheet.data.forEach((row, idx) => {
        let dynamicValue: unknown = null;
        if (dynamicSourceNodeId) {
          dynamicValue = resolveNodeValue(dynamicSourceNodeId, row as Record<string, string | number | boolean | null>);
        } else if (dynamicSourceLabel) {
          dynamicValue = (row as Record<string, unknown>)[dynamicSourceLabel] ?? null;
        }
        if (outRows[idx]) outRows[idx][dynamicTargetLabel] = dynamicValue;
      });
    }

    return outRows;
  }

  function isValueDifferent(a: unknown, b: unknown): boolean {
    if (a === b) return false;
    const left = a === null || a === undefined ? "" : String(a).trim();
    const right = b === null || b === undefined ? "" : String(b).trim();
    return left !== right;
  }

  function buildChangePreview(inputRows: Record<string, unknown>[], outputRows: Record<string, unknown>[]): ChangePreview {
    const changes: ChangeCell[] = [];
    const changedRows = new Set<number>();
    const changedColumns = new Set<string>();

    outputRows.forEach((outputRow, rowIdx) => {
      const inputRow = inputRows[rowIdx] || {};
      Object.keys(outputRow).forEach((column) => {
        const from = (inputRow as Record<string, unknown>)[column] ?? null;
        const to = outputRow[column] ?? null;
        if (!isValueDifferent(from, to)) return;
        const rowNumber = rowIdx + 1;
        changes.push({ row: rowNumber, column, from, to });
        changedRows.add(rowNumber);
        changedColumns.add(column);
      });
    });

    return {
      changes,
      changedRows: Array.from(changedRows).sort((a, b) => a - b),
      changedColumns: Array.from(changedColumns),
      changedCellCount: changes.length,
      changedRowCount: changedRows.size,
    };
  }

  async function handleConfirm() {
    if (!file) return;
    setLoading(true);
    setStep("running");

    if (runMode === "excel_to_excel") {
      try {
        if (!parsedSheet) throw new Error("No sheet data");
        const mappedRows = outputPreviewRows.length ? outputPreviewRows : applyMappingToSheet(parsedSheet);
        const outputRows = workbookOptions?.excelNewSheetEnabled
          ? (() => {
              const targetHeaders = ((currentVersion?.connections?.nodes ?? []) as Array<{ type?: string; data?: { label?: string } }>)
                .filter((n) => n.type === "ssCol")
                .map((n) => String(n.data?.label ?? "").trim())
                .filter(Boolean);

              const baseRows = (parsedSheet.data as Record<string, unknown>[]).map((row) => {
                const base: Record<string, unknown> = {};
                targetHeaders.forEach((header) => {
                  base[header] = workbookOptions?.excelNewSheetKeepExistingData ? row[header] ?? null : null;
                });
                return base;
              });

              return mappedRows.map((row, idx) => ({ ...(baseRows[idx] ?? {}), ...row }));
            })()
          : mappedRows;
        const preview = changePreview ?? buildChangePreview(parsedSheet.data as Record<string, unknown>[], outputRows);

        const wb = XLSX.utils.book_new();
        const outputSheetName = toSheetName(String(workbookOptions?.outputSheetName ?? "Mapped"), "Mapped");
        const sourceSheetName = toSheetName(String(workbookOptions?.inputSheetName ?? "Source"), "Source");

        const outputWs = XLSX.utils.json_to_sheet(outputRows);
        XLSX.utils.book_append_sheet(wb, outputWs, outputSheetName);

        if ((parsedSheet?.data?.length ?? 0) > 0) {
          const inputWs = XLSX.utils.json_to_sheet(parsedSheet?.data ?? []);
          const inputName = sourceSheetName === outputSheetName ? `${sourceSheetName.slice(0, 28)}_in` : sourceSheetName;
          XLSX.utils.book_append_sheet(wb, inputWs, inputName);
        }

        const outName = `${mapping.name.replace(/\s+/g, "_")}_mapped.xlsx`;
        XLSX.writeFile(wb, outName);

        // Save the run to database
        const versionForRun = mapping.versions.find((v) => v.id === mapping.currentVersionId) ?? mapping.versions.at(-1);
        if (versionForRun) {
          const saveRunRes = await fetch(`/api/mappings/${mapping.id}/runs`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
              mappingVersionId: versionForRun.id,
              direction: "excel_to_excel",
              inputFileName: file.name,
              inputData: parsedSheet.data,
              outputData: outputRows,
              changeSet: preview.changes,
              changedColumns: preview.changedColumns,
              changedCellCount: preview.changedCellCount,
              changedRowCount: preview.changedRowCount,
              rowCount: outputRows.length,
            }),
          });
          if (!saveRunRes.ok) {
            const err = await saveRunRes.json().catch(() => ({ error: "Failed to save run history" }));
            throw new Error(err.error || "Failed to save run history");
          }
        }

        toast({ title: "Download ready!", description: `Saved as ${outName}. Run saved to history.` });
        setStep("done");
      } catch (e) {
        toast({ title: "Mapping failed", description: (e as Error).message, variant: "destructive" });
        setStep("approve");
      } finally {
        setLoading(false);
      }
    } else {
      try {
        const outputRows = outputPreviewRows.length ? outputPreviewRows : (parsedSheet ? applyMappingToSheet(parsedSheet) : []);

        const res = await fetch(`/api/mappings/${mapping.id}/staging`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            direction: "excel_to_ss",
            excelData: parsedSheet?.data ?? null,
            outputRows,
            runOptions,
          }),
        });
        if (!res.ok) throw new Error((await res.json()).error);
        const stagingRun = await res.json();
        const diffRows = Array.isArray(stagingRun?.diffResult) ? stagingRun.diffResult : [];
        const diagnostics =
          stagingRun?.stagingExcelData && typeof stagingRun.stagingExcelData === "object"
            ? (stagingRun.stagingExcelData as {
                _rowPolicyDiagnostics?: {
                  totalOutputRows?: number;
                  skippedCells?: number;
                  skipByReason?: Record<string, number>;
                  parentRowsDetected?: number;
                  formulaRowsDetected?: number;
                };
              })._rowPolicyDiagnostics
            : null;
        setLastPolicyDiagnostics(diagnostics ?? null);

        if (mapping.autoPush && stagingRun?.id) {
          const autoRes = await fetch(`/api/mappings/${mapping.id}/staging/${stagingRun.id}`, {
            method: "PUT",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
              status: "merged",
              mergeResolution: diffRows,
              rowsChanged: diffRows.filter((c: { action?: string; resolution?: string }) => c?.action !== "skip" && c?.resolution !== "keep_production").length,
              conflictCount: diffRows.filter((c: { isConflict?: boolean }) => c?.isConflict).length,
              mergedAt: new Date().toISOString(),
            }),
          });
          if (!autoRes.ok) throw new Error((await autoRes.json()).error || "Auto push failed");
          const skipped = diagnostics?.skippedCells ?? 0;
          toast({
            title: "Auto push complete",
            description: skipped > 0
              ? `Changes pushed. Skipped ${skipped} protected/excluded cell updates.`
              : "Changes were pushed to production Smartsheet.",
          });
          onClose();
          return;
        }

        const skipped = diagnostics?.skippedCells ?? 0;
        toast({
          title: "Staging run started",
          description: skipped > 0
            ? `Review and merge when ready. ${skipped} cell updates are marked skipped by row policy.`
            : "Review and merge when ready.",
        });
        onClose();
      } catch (e) {
        toast({ title: "Failed to start run", description: (e as Error).message, variant: "destructive" });
        setStep("approve");
      } finally {
        setLoading(false);
      }
    }
  }

  function reset() {
    setStep("upload");
    setFile(null);
    setParsedSheet(null);
    setValidation(null);
    setOutputPreviewRows([]);
    setChangePreview(null);
    setLastPolicyDiagnostics(null);
    setRunOptions({
      hierarchyAware: true,
      protectFormulaCells: true,
      protectParentSummaryRows: true,
      excludedRowPatterns: ["total", "subtotal", "summary"],
      columnExcludedRowPatterns: {},
      excludedRowNumbers: [],
      skipTopRows: 0,
    });
    setExcludedPatternInput("total, subtotal, summary");
    setExcludedRowNumberInput("");
    setColumnExclusionInput("");
    setSkipTopRowsInput("");
  }

  return (
    <Dialog open={open} onOpenChange={(o) => { if (!o) { reset(); onClose(); } }}>
      <DialogContent className="max-h-[90vh] w-[95vw] max-w-lg overflow-y-auto">
        <DialogHeader>
          <DialogTitle>Run mapping: {mapping.name}</DialogTitle>
        </DialogHeader>

        {/* Mode badge */}
        <div className="flex items-center gap-2 rounded-lg border border-[#E5E5E5] bg-[#F9F9F9] px-3 py-2">
          {runMode === "excel_to_excel" ? (
            <><Table2 className="h-4 w-4 text-[#6B6B6B]" /><span className="text-xs text-[#6B6B6B]">Excel → Excel (download transformed file)</span></>
          ) : (
            <><Sheet className="h-4 w-4 text-[#6B6B6B]" /><span className="text-xs text-[#6B6B6B]">Excel → Smartsheet (staged review, no sheet copy)</span></>
          )}
        </div>

        {step === "upload" && (
          <div className="space-y-3">
            <div
              {...getRootProps()}
              className={`flex flex-col items-center justify-center rounded-lg border-2 border-dashed p-10 text-center cursor-pointer transition-colors ${
                isDragActive ? "border-black bg-[#F5F5F5]" : "border-[#E5E5E5] hover:border-[#A1A1A1]"
              }`}
            >
              <input {...getInputProps()} />
              <Upload className="mb-3 h-8 w-8 text-[#A1A1A1]" />
              <p className="text-sm font-medium">Drop your source Excel file here</p>
              <p className="mt-1 text-xs text-[#6B6B6B]">or click to browse — .xlsx, .xls, .csv</p>
            </div>
            <div className="flex items-center justify-between rounded-lg border border-[#E5E5E5] bg-[#F9F9F9] px-3 py-2">
              <span className="text-xs text-[#6B6B6B]">Complex formatting</span>
              <Switch checked={complexFormattingEnabled} disabled />
            </div>
          </div>
        )}

        {step === "validate" && validation?.status === "blocked" && (
          <div className="space-y-4">
            <div className="flex items-start gap-3 rounded-lg bg-red-50 p-4">
              <AlertCircle className="mt-0.5 h-5 w-5 shrink-0 text-red-500" />
              <div>
                <p className="text-sm font-medium text-red-800">File structure mismatch</p>
                <p className="mt-0.5 text-xs text-red-600">Required columns missing:</p>
                <ul className="mt-2 space-y-1">
                  {validation.missingColumns.map((col) => (
                    <li key={col} className="flex items-center gap-2 text-xs text-red-700">
                      <span className="h-1 w-1 rounded-full bg-red-500" />
                      {col}
                    </li>
                  ))}
                </ul>
              </div>
            </div>
            <Button variant="outline" onClick={reset} className="w-full">Upload a different file</Button>
          </div>
        )}

        {step === "approve" && validation && (
          <div className="space-y-4">
            <div className="flex items-start gap-3 rounded-lg bg-[#F5F5F5] p-4">
              {validation.status === "exact" ? (
                <CheckCircle2 className="mt-0.5 h-5 w-5 shrink-0 text-green-600" />
              ) : (
                <AlertCircle className="mt-0.5 h-5 w-5 shrink-0 text-amber-500" />
              )}
              <div>
                <p className="text-sm font-medium">
                  {validation.status === "exact" ? "Structure matches perfectly" : "Smart remap applied"}
                </p>
                <p className="mt-0.5 text-xs text-[#6B6B6B]">
                  {validation.status === "exact"
                    ? "All columns match the saved mapping schema."
                    : `${validation.remappedColumns.length} column(s) were auto-remapped by name.`}
                </p>
              </div>
            </div>

            {validation.remappedColumns.length > 0 && (
              <div className="rounded-lg border border-[#E5E5E5]">
                <div className="grid grid-cols-3 gap-2 border-b border-[#E5E5E5] bg-[#F9F9F9] px-3 py-2 text-xs font-medium text-[#6B6B6B]">
                  <span>Expected</span>
                  <span className="text-center"><ArrowRight className="inline h-3 w-3" /></span>
                  <span>Found as</span>
                </div>
                {validation.remappedColumns.map((r) => (
                  <div key={r.expected} className="grid grid-cols-3 gap-2 border-b border-[#F5F5F5] px-3 py-2 text-xs last:border-0">
                    <span className="font-mono">{r.expected}</span>
                    <span />
                    <span className="font-mono text-[#6B6B6B]">{r.found}</span>
                  </div>
                ))}
              </div>
            )}

            {validation.extraColumns.length > 0 && (
              <p className="text-xs text-[#6B6B6B]">
                {validation.extraColumns.length} extra column(s) ignored: {validation.extraColumns.join(", ")}
              </p>
            )}

            {changePreview && (
              <div className="space-y-3 rounded-lg border border-[#E5E5E5] p-3">
                <p className="text-xs font-medium">Run impact preview</p>
                <div className="grid grid-cols-3 gap-2 text-xs">
                  <div className="rounded-md bg-amber-50 px-2 py-1 text-amber-800">{changePreview.changedCellCount} cells changing</div>
                  <div className="rounded-md bg-blue-50 px-2 py-1 text-blue-800">{changePreview.changedRowCount} rows impacted</div>
                  <div className="rounded-md bg-emerald-50 px-2 py-1 text-emerald-800">{changePreview.changedColumns.length} columns impacted</div>
                </div>
                {changePreview.changedCellCount > 0 ? (
                  <div className="max-h-48 overflow-auto rounded-md border border-[#F0F0F0]">
                    <table className="w-full text-xs">
                      <thead className="sticky top-0 bg-[#F9F9F9]">
                        <tr className="border-b border-[#E5E5E5]">
                          <th className="px-2 py-1 text-left text-[#6B6B6B]">Row</th>
                          <th className="px-2 py-1 text-left text-[#6B6B6B]">Column</th>
                          <th className="px-2 py-1 text-left text-[#6B6B6B]">Current</th>
                          <th className="px-2 py-1 text-center text-[#6B6B6B]"><ArrowRight className="inline h-3 w-3" /></th>
                          <th className="px-2 py-1 text-left text-[#6B6B6B]">New</th>
                        </tr>
                      </thead>
                      <tbody>
                        {changePreview.changes.slice(0, 200).map((c, i) => (
                          <tr key={`${c.row}-${c.column}-${i}`} className="border-b border-[#F5F5F5] bg-amber-50/40">
                            <td className="px-2 py-1 font-medium">{c.row}</td>
                            <td className="px-2 py-1">{c.column}</td>
                            <td className="px-2 py-1 font-mono text-[#6B6B6B]">{String(c.from ?? "")}</td>
                            <td className="px-2 py-1 text-center text-[#A1A1A1]">→</td>
                            <td className="px-2 py-1 font-mono font-semibold">{String(c.to ?? "")}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                ) : (
                  <p className="text-xs text-[#6B6B6B]">No cell changes detected for this run.</p>
                )}
              </div>
            )}

            {runMode === "excel_to_ss" && (
              <div className="space-y-3 rounded-lg border border-[#E5E5E5] p-3">
                <p className="text-xs font-medium">Smartsheet safety options</p>
                {lastPolicyDiagnostics && (
                  <div className="rounded-md border border-[#EFEFEF] bg-[#FAFAFA] px-3 py-2 text-[11px] text-[#5A5A5A]">
                    <div>Last run diagnostics: analyzed <span className="font-medium text-black">{lastPolicyDiagnostics.totalOutputRows ?? 0}</span> rows, skipped <span className="font-medium text-black">{lastPolicyDiagnostics.skippedCells ?? 0}</span> cell updates.</div>
                    <div>Detected parent rows: <span className="font-medium text-black">{lastPolicyDiagnostics.parentRowsDetected ?? 0}</span>, formula rows: <span className="font-medium text-black">{lastPolicyDiagnostics.formulaRowsDetected ?? 0}</span>.</div>
                  </div>
                )}
                <div className="space-y-2">
                  <div className="flex items-center justify-between rounded-md border border-[#EFEFEF] px-3 py-2">
                    <div>
                      <p className="text-xs font-medium">Hierarchy-aware mode</p>
                      <p className="text-[11px] text-[#6B6B6B]">Detect parent/child sections and include section context in diff.</p>
                    </div>
                    <Switch
                      checked={runOptions.hierarchyAware}
                      onCheckedChange={(checked) => setRunOptions((prev) => ({ ...prev, hierarchyAware: checked }))}
                    />
                  </div>
                  <div className="flex items-center justify-between rounded-md border border-[#EFEFEF] px-3 py-2">
                    <div>
                      <p className="text-xs font-medium">Protect formula cells</p>
                      <p className="text-[11px] text-[#6B6B6B]">Skip cells that already contain Smartsheet formulas.</p>
                    </div>
                    <Switch
                      checked={runOptions.protectFormulaCells}
                      onCheckedChange={(checked) => setRunOptions((prev) => ({ ...prev, protectFormulaCells: checked }))}
                    />
                  </div>
                  <div className="flex items-center justify-between rounded-md border border-[#EFEFEF] px-3 py-2">
                    <div>
                      <p className="text-xs font-medium">Protect parent/summary rows</p>
                      <p className="text-[11px] text-[#6B6B6B]">Skip rows that act as rollups/section headers.</p>
                    </div>
                    <Switch
                      checked={runOptions.protectParentSummaryRows}
                      onCheckedChange={(checked) => setRunOptions((prev) => ({ ...prev, protectParentSummaryRows: checked }))}
                    />
                  </div>
                  <div className="space-y-2 rounded-md border border-[#EFEFEF] px-3 py-2">
                    <p className="text-xs font-medium">Skip top N rows</p>
                    <p className="text-[11px] text-[#6B6B6B]">Always skip the first N rows of the destination sheet (e.g. 2 for two header/formula rows at the top).</p>
                    <Input
                      value={skipTopRowsInput}
                      onChange={(e: ChangeEvent<HTMLInputElement>) => {
                        const raw = e.target.value;
                        setSkipTopRowsInput(raw);
                        const n = Math.max(0, Math.floor(Number(raw)));
                        setRunOptions((prev) => ({ ...prev, skipTopRows: Number.isFinite(n) ? n : 0 }));
                      }}
                      placeholder="2"
                      type="number"
                      min="0"
                    />
                  </div>
                  <div className="space-y-2 rounded-md border border-[#EFEFEF] px-3 py-2">
                    <p className="text-xs font-medium">Exclude rows by keyword</p>
                    <p className="text-[11px] text-[#6B6B6B]">Skip destination rows whose path/title contains these tokens.</p>
                    <Input
                      value={excludedPatternInput}
                      onChange={(e: ChangeEvent<HTMLInputElement>) => {
                        const raw = e.target.value;
                        setExcludedPatternInput(raw);
                        const patterns = raw
                          .split(",")
                          .map((s: string) => s.trim().toLowerCase())
                          .filter(Boolean);
                        setRunOptions((prev) => ({ ...prev, excludedRowPatterns: patterns }));
                      }}
                      placeholder="total, subtotal, summary"
                    />
                  </div>
                  <div className="space-y-2 rounded-md border border-[#EFEFEF] px-3 py-2">
                    <p className="text-xs font-medium">Exclude destination row numbers</p>
                    <p className="text-[11px] text-[#6B6B6B]">Rows are skipped in-place; source order is preserved.</p>
                    <Input
                      value={excludedRowNumberInput}
                      onChange={(e: ChangeEvent<HTMLInputElement>) => {
                        const raw = e.target.value;
                        setExcludedRowNumberInput(raw);
                        const rows = raw
                          .split(",")
                          .map((s: string) => Number(s.trim()))
                          .filter((n) => Number.isFinite(n) && n > 0)
                          .map((n) => Math.floor(n));
                        setRunOptions((prev) => ({ ...prev, excludedRowNumbers: rows }));
                      }}
                      placeholder="8, 18, 32"
                    />
                  </div>
                  <div className="space-y-2 rounded-md border border-[#EFEFEF] px-3 py-2">
                    <p className="text-xs font-medium">Column-specific row exclusion</p>
                    <p className="text-[11px] text-[#6B6B6B]">Format: <code>Column A=total|summary; Column B=capstone</code></p>
                    <Input
                      value={columnExclusionInput}
                      onChange={(e: ChangeEvent<HTMLInputElement>) => {
                        const raw = e.target.value;
                        setColumnExclusionInput(raw);
                        const mapping: Record<string, string[]> = {};
                        raw
                          .split(";")
                          .map((s: string) => s.trim())
                          .filter(Boolean)
                          .forEach((rule: string) => {
                            const [colPart, patternPart] = rule.split("=");
                            const colName = (colPart || "").trim();
                            const patterns = (patternPart || "")
                              .split("|")
                              .map((p: string) => p.trim().toLowerCase())
                              .filter(Boolean);
                            if (colName && patterns.length > 0) {
                              mapping[colName] = patterns;
                            }
                          });
                        setRunOptions((prev) => ({ ...prev, columnExcludedRowPatterns: mapping }));
                      }}
                      placeholder="2023 Totals=total|summary; April 2024=subtotal"
                    />
                  </div>
                </div>
              </div>
            )}

            <div className="flex gap-2">
              <Button variant="outline" onClick={reset} className="flex-1">Change file</Button>
              <Button onClick={handleConfirm} className="flex-1" disabled={loading}>
                {loading ? <Loader2 className="mr-2 h-4 w-4 animate-spin" /> : null}
                {runMode === "excel_to_excel" ? "Apply & Download" : "Confirm & Apply"}
              </Button>
            </div>
          </div>
        )}

        {step === "running" && (
          <div className="flex flex-col items-center gap-3 py-8">
            <Loader2 className="h-8 w-8 animate-spin text-[#A1A1A1]" />
            <p className="text-sm font-medium">
              {runMode === "excel_to_excel" ? "Applying mapping…" : "Creating staging copy…"}
            </p>
            <p className="text-xs text-[#6B6B6B]">Just a moment</p>
          </div>
        )}

        {step === "done" && (
          <div className="flex flex-col items-center gap-3 py-8">
            <div className="flex h-12 w-12 items-center justify-center rounded-full bg-green-50">
              <Download className="h-6 w-6 text-green-600" />
            </div>
            <p className="text-sm font-medium">Mapping applied &amp; downloaded!</p>
            <p className="text-xs text-[#6B6B6B]">Check your <strong>Downloads</strong> folder — the file is saved as <code className="font-mono text-[10px] bg-[#F5F5F5] px-1 rounded">{mapping.name.replace(/\s+/g, "_")}_mapped.xlsx</code>.</p>
            <div className="mt-2 flex gap-2">
              <Button variant="outline" onClick={reset}>Run again</Button>
              <Button onClick={onClose}>Done</Button>
            </div>
          </div>
        )}
      </DialogContent>
    </Dialog>
  );
}
