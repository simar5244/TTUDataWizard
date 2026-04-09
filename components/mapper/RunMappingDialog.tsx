"use client";

import { useState, useCallback } from "react";
import { useDropzone } from "react-dropzone";
import { Upload, CheckCircle2, AlertCircle, ArrowRight, Loader2, Download, Table2, Sheet } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Dialog, DialogContent, DialogHeader, DialogTitle } from "@/components/ui/dialog";
import { parseExcelFile, validateExcelAgainstFingerprint, type SchemaFingerprint, type ExcelSheet } from "@/lib/excel";
import { useToast } from "@/hooks/use-toast";
import type { MappingConnections } from "@/app/mapper/page";
import * as XLSX from "xlsx";
import { evaluate } from "mathjs";

interface MappingVersion {
  id: string;
  versionNumber: number;
  schemaFingerprint?: unknown;
  connections?: MappingConnections;
  formulas?: Record<string, string>;
}

interface RunMappingDialogProps {
  mapping: {
    id: string;
    name: string;
    smartsheetSheetId?: string | null;
    currentVersionId: string | null;
    versions: MappingVersion[];
  };
  open: boolean;
  onClose: () => void;
}

type Step = "upload" | "validate" | "approve" | "running" | "done";
type RunMode = "excel_to_excel" | "excel_to_ss";

export function RunMappingDialog({ mapping, open, onClose }: RunMappingDialogProps) {
  const { toast } = useToast();
  const [step, setStep] = useState<Step>("upload");
  const [file, setFile] = useState<File | null>(null);
  const [parsedSheet, setParsedSheet] = useState<ExcelSheet | null>(null);
  const [validation, setValidation] = useState<ReturnType<typeof validateExcelAgainstFingerprint> | null>(null);
  const [loading, setLoading] = useState(false);

  const isExcelToExcel = !mapping.smartsheetSheetId;
  const runMode: RunMode = isExcelToExcel ? "excel_to_excel" : "excel_to_ss";

  const currentVersion = mapping.versions.find((v) => v.id === mapping.currentVersionId) ?? mapping.versions.at(-1);
  const fingerprint = currentVersion?.schemaFingerprint as SchemaFingerprint | undefined;

  const onDrop = useCallback(async (acceptedFiles: File[]) => {
    const f = acceptedFiles[0];
    if (!f) return;
    setFile(f);
    setStep("validate");

    const buf = await f.arrayBuffer();
    const parsed = parseExcelFile(buf, f.name, f.size);
    const sheet = parsed.sheets[0];

    if (!sheet) {
      toast({ title: "No data found in Excel file", variant: "destructive" });
      setStep("upload");
      return;
    }

    setParsedSheet(sheet);

    if (!fingerprint) {
      setValidation({ status: "exact", missingColumns: [], remappedColumns: [], extraColumns: [] });
      setStep("approve");
      return;
    }

    const result = validateExcelAgainstFingerprint(sheet, fingerprint);
    setValidation(result);
    if (result.status === "blocked") setStep("validate");
    else setStep("approve");
  }, [fingerprint, toast]);

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
    if (!conns?.edges?.length || !conns?.nodes?.length) {
      return sheet.data;
    }

    const nodes = conns.nodes ?? [];
    const edges = conns.edges ?? [];

    return sheet.data.map((row) => {
      const outRow: Record<string, unknown> = {};

      for (const edge of edges) {
        const srcNode = nodes.find((n) => n.id === edge.source);
        const tgtNode = nodes.find((n) => n.id === edge.target);
        if (!srcNode || !tgtNode) continue;

        const srcLabel = srcNode.data.label ?? "";
        const tgtLabel = tgtNode.data.label ?? "";
        const rawVal = row[srcLabel];
        const formula = edge.data?.formula ?? "";

        if (!formula) {
          outRow[tgtLabel] = rawVal;
        } else {
          try {
            let expr = formula;
            nodes
              .filter((n) => n.type === "excelCol")
              .forEach((n) => {
                const colVal = row[n.data.label ?? ""] ?? 0;
                const safe = String(colVal).replace(/[^0-9.\-]/g, "") || "0";
                expr = expr.replace(new RegExp(`\\b${n.data.colKey ?? n.data.label}\\b`, "g"), safe);
              });
            outRow[tgtLabel] = evaluate(expr);
          } catch {
            outRow[tgtLabel] = rawVal;
          }
        }
      }
      return outRow;
    });
  }

  async function handleConfirm() {
    if (!file) return;
    setLoading(true);
    setStep("running");

    if (runMode === "excel_to_excel") {
      try {
        if (!parsedSheet) throw new Error("No sheet data");
        const outputRows = applyMappingToSheet(parsedSheet);

        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.json_to_sheet(outputRows);
        XLSX.utils.book_append_sheet(wb, ws, "Mapped");
        const outName = `${mapping.name.replace(/\s+/g, "_")}_mapped.xlsx`;
        XLSX.writeFile(wb, outName);

        toast({ title: "Download ready!", description: `Saved as ${outName}` });
        setStep("done");
      } catch (e) {
        toast({ title: "Mapping failed", description: (e as Error).message, variant: "destructive" });
        setStep("approve");
      } finally {
        setLoading(false);
      }
    } else {
      try {
        const res = await fetch(`/api/mappings/${mapping.id}/staging`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ direction: "excel_to_ss" }),
        });
        if (!res.ok) throw new Error((await res.json()).error);
        toast({ title: "Staging run started", description: "Review and merge when ready." });
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
  }

  return (
    <Dialog open={open} onOpenChange={(o) => { if (!o) { reset(); onClose(); } }}>
      <DialogContent className="max-w-lg">
        <DialogHeader>
          <DialogTitle>Run mapping: {mapping.name}</DialogTitle>
        </DialogHeader>

        {/* Mode badge */}
        <div className="flex items-center gap-2 rounded-lg border border-[#E5E5E5] bg-[#F9F9F9] px-3 py-2">
          {runMode === "excel_to_excel" ? (
            <><Table2 className="h-4 w-4 text-[#6B6B6B]" /><span className="text-xs text-[#6B6B6B]">Excel → Excel (download transformed file)</span></>
          ) : (
            <><Sheet className="h-4 w-4 text-[#6B6B6B]" /><span className="text-xs text-[#6B6B6B]">Excel → Smartsheet (creates staging copy)</span></>
          )}
        </div>

        {step === "upload" && (
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
            <p className="text-sm font-medium">Mapping applied successfully!</p>
            <p className="text-xs text-[#6B6B6B]">Your file was downloaded to your default downloads folder.</p>
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
