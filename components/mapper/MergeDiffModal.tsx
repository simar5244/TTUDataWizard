"use client";

import { useState, useEffect } from "react";
import { CheckCircle2, Loader2, ArrowRight } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Dialog, DialogContent, DialogHeader, DialogTitle } from "@/components/ui/dialog";
import { Badge } from "@/components/ui/badge";
import { useToast } from "@/hooks/use-toast";

interface DiffCell {
  row: number;
  rowId?: number | null;
  rowPath?: string | null;
  rowKind?: "detail" | "parent_or_summary" | "new_row";
  column: string;
  productionValue: unknown;
  stagingValue: unknown;
  isConflict: boolean;
  resolution?: "keep_production" | "use_staging" | "overwrite";
  overwriteValue?: string;
  action?: "update" | "append" | "skip";
  skipReason?: "protected_formula" | "protected_parent_summary" | "locked";
}

interface MergeDiffModalProps {
  run: {
    id: string;
    status: string;
    diffResult: unknown;
    stagingExcelData?: unknown;
    conflictCount: number | null;
    rowsChanged: number | null;
  };
  mappingId: string;
  open: boolean;
  onClose: () => void;
  onMerged: () => void;
}

export function MergeDiffModal({ run, mappingId, open, onClose, onMerged }: MergeDiffModalProps) {
  const { toast } = useToast();
  const [cells, setCells] = useState<DiffCell[]>([]);
  const [filter, setFilter] = useState<"all" | "conflicts">("all");
  const [sectionedView, setSectionedView] = useState(true);
  const [merging, setMerging] = useState(false);

  const diagnostics =
    run.stagingExcelData && typeof run.stagingExcelData === "object"
      ? (run.stagingExcelData as { _rowPolicyDiagnostics?: {
          totalOutputRows?: number;
          totalDiffCells?: number;
          skippedCells?: number;
          updatedCells?: number;
          appendedCells?: number;
          skipByReason?: Record<string, number>;
          parentRowsDetected?: number;
          formulaRowsDetected?: number;
          hierarchyRowsDetected?: number;
        } })._rowPolicyDiagnostics
      : undefined;
  useEffect(() => {
    if (run.diffResult) {
      setCells((run.diffResult as DiffCell[]) || []);
    }
  }, [run.diffResult]);

  const conflicts = cells.filter((c) => c.isConflict);
  const unresolvedConflicts = conflicts.filter((c) => !c.resolution);
  const displayCells = filter === "conflicts" ? conflicts : cells;
  const groupedBySection = displayCells.reduce((acc, cell) => {
    const key = cell.rowPath?.trim() || "Ungrouped";
    if (!acc.has(key)) acc.set(key, [] as DiffCell[]);
    acc.get(key)?.push(cell);
    return acc;
  }, new Map<string, DiffCell[]>());

  function resolveCell(idx: number, resolution: DiffCell["resolution"], overwriteValue?: string) {
    setCells((prev) => {
      const updated = [...prev];
      const globalIdx = prev.indexOf(displayCells[idx]);
      updated[globalIdx] = { ...updated[globalIdx], resolution, overwriteValue };
      return updated;
    });
  }

  function acceptAllStaging() {
    setCells((prev) => prev.map((c) => c.isConflict ? { ...c, resolution: "use_staging" } : c));
  }

  function keepAllProduction() {
    setCells((prev) => prev.map((c) => c.isConflict ? { ...c, resolution: "keep_production" } : c));
  }

  async function handleMerge() {
    if (unresolvedConflicts.length > 0) {
      toast({ title: "Resolve all conflicts before merging", variant: "destructive" });
      return;
    }
    setMerging(true);
    try {
      const res = await fetch(`/api/mappings/${mappingId}/staging/${run.id}`, {
        method: "PUT",
        headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            status: "merged",
            mergeResolution: cells,
            rowsChanged: cells.filter((c) => c.resolution !== "keep_production" && c.action !== "skip").length,
            conflictCount: conflicts.length,
            mergedAt: new Date().toISOString(),
          }),
      });
      if (!res.ok) throw new Error((await res.json()).error);
      const merged = await res.json();
      const applied = typeof merged.appliedCellCount === "number" ? merged.appliedCellCount : cells.length;
      const appended = typeof merged.appendedRowCount === "number" ? merged.appendedRowCount : 0;
      toast({ title: "Merged successfully", description: `Applied ${applied} cell updates${appended ? `, appended ${appended} rows` : ""}.` });
      onMerged();
    } catch (e) {
      toast({ title: "Merge failed", description: (e as Error).message, variant: "destructive" });
    } finally {
      setMerging(false);
    }
  }

  const readOnly = run.status === "merged";

  return (
    <Dialog open={open} onOpenChange={(o) => { if (!o) onClose(); }}>
      <DialogContent className="max-w-4xl max-h-[85vh] flex flex-col">
        <DialogHeader>
          <DialogTitle className="flex items-center gap-3">
            <span>{readOnly ? "Merge diff (read-only)" : "Review & Merge"}</span>
            <Badge variant="outline" className="text-xs">
              {conflicts.length} conflict{conflicts.length !== 1 ? "s" : ""} · {cells.filter((c) => !c.isConflict).length} auto-merged
            </Badge>
            {!readOnly && unresolvedConflicts.length > 0 && (
              <Badge variant="destructive" className="text-xs">
                {unresolvedConflicts.length} unresolved
              </Badge>
            )}
          </DialogTitle>
        </DialogHeader>

        {diagnostics && (
          <div className="grid grid-cols-2 gap-2 rounded-md border border-[#E5E5E5] bg-[#FAFAFA] p-2 text-[11px]">
            <div className="text-[#5A5A5A]">Rows analyzed: <span className="font-medium text-black">{diagnostics.totalOutputRows ?? 0}</span></div>
            <div className="text-[#5A5A5A]">Cells skipped: <span className="font-medium text-black">{diagnostics.skippedCells ?? 0}</span></div>
            <div className="text-[#5A5A5A]">Parent rows detected: <span className="font-medium text-black">{diagnostics.parentRowsDetected ?? 0}</span></div>
            <div className="text-[#5A5A5A]">Formula rows detected: <span className="font-medium text-black">{diagnostics.formulaRowsDetected ?? 0}</span></div>
          </div>
        )}

        {/* Controls */}
        {conflicts.length > 0 && (
          <div className="flex items-center justify-between border-b border-[#E5E5E5] pb-3">
            <div className="flex gap-1 rounded-lg bg-[#F5F5F5] p-1 text-xs">
              <button
                onClick={() => setFilter("conflicts")}
                className={`rounded-md px-3 py-1 transition-colors ${filter === "conflicts" ? "bg-white shadow" : "text-[#6B6B6B] hover:text-black"}`}
              >
                Conflicts only ({conflicts.length})
              </button>
              <button
                onClick={() => setFilter("all")}
                className={`rounded-md px-3 py-1 transition-colors ${filter === "all" ? "bg-white shadow" : "text-[#6B6B6B] hover:text-black"}`}
              >
                All changes ({cells.length})
              </button>
            </div>
            {!readOnly && (
              <div className="flex gap-2">
                <Button size="sm" variant="outline" onClick={() => setSectionedView((v) => !v)}>
                  {sectionedView ? "Flat preview" : "Sectioned preview"}
                </Button>
                <Button size="sm" variant="outline" onClick={keepAllProduction}>Keep all production</Button>
                <Button size="sm" variant="outline" onClick={acceptAllStaging}>Accept all staging</Button>
              </div>
            )}
          </div>
        )}

        {/* Diff table */}
        <div className="flex-1 overflow-auto">
          {displayCells.length === 0 ? (
            <div className="flex flex-col items-center justify-center py-12">
              <CheckCircle2 className="mb-2 h-8 w-8 text-green-500" />
              <p className="text-sm font-medium">No conflicts to resolve</p>
              <p className="text-xs text-[#6B6B6B]">All changes auto-merged</p>
            </div>
          ) : sectionedView ? (
            <div className="space-y-4">
              {Array.from(groupedBySection.entries()).map(([section, sectionCells]) => (
                <div key={section} className="rounded-lg border border-[#EAEAEA]">
                  <div className="flex items-center justify-between border-b border-[#EAEAEA] bg-[#F9F9F9] px-3 py-2">
                    <p className="text-xs font-medium">{section}</p>
                    <Badge variant="outline" className="text-[10px]">{sectionCells.length} changes</Badge>
                  </div>
                  <table className="w-full text-xs">
                    <thead className="bg-[#FCFCFC]">
                      <tr className="border-b border-[#EFEFEF]">
                        <th className="px-3 py-2 text-left font-medium text-[#6B6B6B]">Row</th>
                        <th className="px-3 py-2 text-left font-medium text-[#6B6B6B]">Column</th>
                        <th className="px-3 py-2 text-left font-medium text-[#6B6B6B]">Production</th>
                        <th className="px-3 py-2 text-center text-[#6B6B6B]"><ArrowRight className="inline h-3 w-3" /></th>
                        <th className="px-3 py-2 text-left font-medium text-[#6B6B6B]">Staging</th>
                        {!readOnly && <th className="px-3 py-2 text-left font-medium text-[#6B6B6B]">Choice</th>}
                      </tr>
                    </thead>
                    <tbody>
                      {sectionCells.map((cell, i) => (
                        <tr key={`${section}-${i}`} className={`border-b border-[#F5F5F5] ${cell.isConflict ? "bg-amber-50/50" : ""}`}>
                          <td className="px-3 py-2 text-[#6B6B6B]">{cell.row}</td>
                          <td className="px-3 py-2 font-medium">
                            {cell.column}
                            {cell.action === "skip" && (
                              <span className="ml-2 rounded bg-slate-100 px-1.5 py-0.5 text-[9px] text-slate-600">skipped</span>
                            )}
                          </td>
                          <td className={`px-3 py-2 font-mono ${cell.resolution === "keep_production" ? "font-bold text-green-700" : ""}`}>
                            {String(cell.productionValue ?? "")}
                          </td>
                          <td className="px-3 py-2 text-center text-[#A1A1A1]">→</td>
                          <td className={`px-3 py-2 font-mono ${cell.resolution === "use_staging" ? "font-bold text-green-700" : ""}`}>
                            {String(cell.stagingValue ?? "")}
                          </td>
                          {!readOnly && cell.isConflict && (
                            <td className="px-3 py-2">
                              <div className="flex items-center gap-1">
                                <button
                                  onClick={() => resolveCell(displayCells.indexOf(cell), "keep_production")}
                                  className={`rounded px-2 py-0.5 text-[10px] transition-colors ${cell.resolution === "keep_production" ? "bg-black text-white" : "border border-[#E5E5E5] hover:bg-[#F5F5F5]"}`}
                                >
                                  Keep prod
                                </button>
                                <button
                                  onClick={() => resolveCell(displayCells.indexOf(cell), "use_staging")}
                                  className={`rounded px-2 py-0.5 text-[10px] transition-colors ${cell.resolution === "use_staging" ? "bg-black text-white" : "border border-[#E5E5E5] hover:bg-[#F5F5F5]"}`}
                                >
                                  Use staging
                                </button>
                              </div>
                            </td>
                          )}
                          {!readOnly && !cell.isConflict && (
                            <td className="px-3 py-2">
                              {cell.action === "skip"
                                ? <span className="text-[10px] text-slate-600">{cell.skipReason?.replaceAll("_", " ")}</span>
                                : <span className="text-[10px] text-green-600">Auto-merged</span>}
                            </td>
                          )}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              ))}
            </div>
          ) : (
            <table className="w-full text-xs">
              <thead className="sticky top-0 bg-[#F9F9F9]">
                <tr className="border-b border-[#E5E5E5]">
                  <th className="px-3 py-2 text-left font-medium text-[#6B6B6B]">Row</th>
                  <th className="px-3 py-2 text-left font-medium text-[#6B6B6B]">Column</th>
                  <th className="px-3 py-2 text-left font-medium text-[#6B6B6B]">Production</th>
                  <th className="px-3 py-2 text-center text-[#6B6B6B]"><ArrowRight className="inline h-3 w-3" /></th>
                  <th className="px-3 py-2 text-left font-medium text-[#6B6B6B]">Staging</th>
                  {!readOnly && <th className="px-3 py-2 text-left font-medium text-[#6B6B6B]">Choice</th>}
                </tr>
              </thead>
              <tbody>
                {displayCells.map((cell, i) => (
                  <tr key={i} className={`border-b border-[#F5F5F5] ${cell.isConflict ? "bg-amber-50/50" : ""}`}>
                    <td className="px-3 py-2 text-[#6B6B6B]">{cell.row}</td>
                    <td className="px-3 py-2 font-medium">{cell.column}</td>
                    <td className={`px-3 py-2 font-mono ${cell.resolution === "keep_production" ? "font-bold text-green-700" : ""}`}>
                      {String(cell.productionValue ?? "")}
                    </td>
                    <td className="px-3 py-2 text-center text-[#A1A1A1]">→</td>
                    <td className={`px-3 py-2 font-mono ${cell.resolution === "use_staging" ? "font-bold text-green-700" : ""}`}>
                      {String(cell.stagingValue ?? "")}
                    </td>
                    {!readOnly && cell.isConflict && (
                      <td className="px-3 py-2">
                        <div className="flex items-center gap-1">
                          <button
                            onClick={() => resolveCell(i, "keep_production")}
                            className={`rounded px-2 py-0.5 text-[10px] transition-colors ${cell.resolution === "keep_production" ? "bg-black text-white" : "border border-[#E5E5E5] hover:bg-[#F5F5F5]"}`}
                          >
                            Keep prod
                          </button>
                          <button
                            onClick={() => resolveCell(i, "use_staging")}
                            className={`rounded px-2 py-0.5 text-[10px] transition-colors ${cell.resolution === "use_staging" ? "bg-black text-white" : "border border-[#E5E5E5] hover:bg-[#F5F5F5]"}`}
                          >
                            Use staging
                          </button>
                        </div>
                      </td>
                    )}
                    {!readOnly && !cell.isConflict && <td className="px-3 py-2"><span className="text-[10px] text-green-600">Auto-merged</span></td>}
                  </tr>
                ))}
              </tbody>
            </table>
          )}
        </div>

        {/* Footer */}
        {!readOnly && (
          <div className="flex items-center justify-between border-t border-[#E5E5E5] pt-3">
            <p className="text-xs text-[#6B6B6B]">
              {unresolvedConflicts.length === 0
                ? "All conflicts resolved — ready to merge"
                : `${unresolvedConflicts.length} conflict${unresolvedConflicts.length !== 1 ? "s" : ""} remaining`}
            </p>
            <div className="flex gap-2">
              <Button variant="outline" onClick={onClose}>Cancel</Button>
              <Button onClick={handleMerge} disabled={merging || unresolvedConflicts.length > 0}>
                {merging ? <Loader2 className="mr-2 h-4 w-4 animate-spin" /> : <CheckCircle2 className="mr-2 h-4 w-4" />}
                Merge to production
              </Button>
            </div>
          </div>
        )}
      </DialogContent>
    </Dialog>
  );
}
