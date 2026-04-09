"use client";

import { useEffect, useState } from "react";
import { useParams } from "next/navigation";
import Link from "next/link";
import { ArrowLeft, GitMerge, Clock, CheckCircle2, AlertCircle, Loader2, GitBranch } from "lucide-react";
import { AppShell } from "@/components/AppShell";
import { Button } from "@/components/ui/button";
import { Badge } from "@/components/ui/badge";
import { formatDateTime } from "@/lib/utils";
import { MergeDiffModal } from "@/components/mapper/MergeDiffModal";
import { Dialog, DialogContent, DialogHeader, DialogTitle } from "@/components/ui/dialog";

interface StagingRun {
  id: string;
  direction: string;
  status: string;
  rowsChanged: number | null;
  conflictCount: number | null;
  createdAt: string;
  mergedAt: string | null;
  mappingVersion: { versionNumber: number; changeSummary: string | null };
  diffResult: unknown;
  mergeResolution: unknown;
}

interface MappingRun {
  id: string;
  direction?: string;
  rowCount: number;
  changedCellCount?: number | null;
  changedRowCount?: number | null;
  changedColumns?: string[] | null;
  changeSet?: Array<{ row: number; column: string; from: unknown; to: unknown }> | null;
  createdAt: string;
  inputFileName: string | null;
}

export default function MappingHistoryPage() {
  const { id } = useParams<{ id: string }>();
  const [runs, setRuns] = useState<StagingRun[]>([]);
  const [mappingName, setMappingName] = useState("");
  const [loading, setLoading] = useState(true);
  const [selectedRun, setSelectedRun] = useState<StagingRun | null>(null);
  const [excelRuns, setExcelRuns] = useState<MappingRun[]>([]);
  const [selectedExcelRun, setSelectedExcelRun] = useState<MappingRun | null>(null);

  useEffect(() => {
    Promise.all([
      fetch(`/api/mappings/${id}`).then((r) => r.json()),
      fetch(`/api/mappings/${id}/staging`).then((r) => r.json()),
      fetch(`/api/mappings/${id}/runs`).then((r) => r.ok ? r.json() : []),
    ]).then(([m, r, x]) => {
      setMappingName(m.name || "");
      setRuns(Array.isArray(r) ? r : []);
      setExcelRuns(Array.isArray(x) ? x : []);
      setLoading(false);
    });
  }, [id]);

  const statusBadge = (status: string) => {
    switch (status) {
      case "merged": return <Badge className="bg-green-100 text-green-800 border-green-200"><CheckCircle2 className="mr-1 h-3 w-3" />Merged</Badge>;
      case "open": return <Badge variant="secondary"><GitBranch className="mr-1 h-3 w-3" />Open</Badge>;
      case "discarded": return <Badge variant="outline" className="text-[#6B6B6B]">Discarded</Badge>;
      default: return <Badge variant="outline">{status}</Badge>;
    }
  };

  return (
    <AppShell>
      <div className="p-6 lg:p-8">
        <div className="mb-6 flex items-center gap-3">
          <Button asChild variant="ghost" size="sm">
            <Link href={`/mapper/${id}`}><ArrowLeft className="mr-1.5 h-4 w-4" />Back</Link>
          </Button>
          <div>
            <h1 className="text-xl font-semibold">{mappingName}</h1>
            <p className="text-sm text-[#6B6B6B]">Staging run history</p>
          </div>
        </div>

        {loading ? (
          <div className="flex justify-center py-20"><Loader2 className="h-6 w-6 animate-spin text-[#A1A1A1]" /></div>
        ) : runs.length === 0 && excelRuns.length === 0 ? (
          <div className="flex flex-col items-center justify-center rounded-xl border border-dashed border-[#E5E5E5] py-20">
            <GitMerge className="mb-3 h-8 w-8 text-[#D1D1D1]" />
            <p className="text-sm font-medium">No runs yet</p>
            <p className="text-xs text-[#6B6B6B]">Run this mapping to start creating a history</p>
          </div>
        ) : (
          <div className="relative">
            {/* Timeline line */}
            <div className="absolute left-4 top-0 h-full w-0.5 bg-[#E5E5E5]" />
            <div className="space-y-4 pl-12">
              {excelRuns.map((run) => (
                <div key={`excel_${run.id}`} className="relative rounded-xl border border-[#E5E5E5] bg-white p-5 shadow-sm">
                  <div className="absolute -left-9 top-5 flex h-5 w-5 items-center justify-center rounded-full border-2 border-[#E5E5E5] bg-white">
                    <CheckCircle2 className="h-3 w-3 text-blue-600" />
                  </div>
                  <div className="flex items-start justify-between gap-4">
                    <div>
                      <div className="flex items-center gap-2">
                        <Badge className="bg-blue-100 text-blue-800 border-blue-200">Excel Run</Badge>
                        <span className="text-xs text-[#6B6B6B]">Excel → Excel</span>
                      </div>
                      <div className="mt-2 flex items-center gap-4 text-xs text-[#6B6B6B]">
                        <span className="flex items-center gap-1">
                          <Clock className="h-3 w-3" />
                          Ran {formatDateTime(run.createdAt)}
                        </span>
                        <span>{run.rowCount} rows output</span>
                        {typeof run.changedCellCount === "number" && <span>{run.changedCellCount} cells changed</span>}
                        {typeof run.changedRowCount === "number" && <span>{run.changedRowCount} rows impacted</span>}
                        {run.inputFileName && <span>Source: {run.inputFileName}</span>}
                      </div>
                    </div>
                    <Button size="sm" variant="outline" onClick={() => setSelectedExcelRun(run)}>
                      View details
                    </Button>
                  </div>
                </div>
              ))}

              {runs.map((run) => (
                <div key={run.id} className="relative rounded-xl border border-[#E5E5E5] bg-white p-5 shadow-sm">
                  {/* Timeline dot */}
                  <div className="absolute -left-9 top-5 flex h-5 w-5 items-center justify-center rounded-full border-2 border-[#E5E5E5] bg-white">
                    {run.status === "merged" ? (
                      <CheckCircle2 className="h-3 w-3 text-green-600" />
                    ) : run.status === "open" ? (
                      <AlertCircle className="h-3 w-3 text-amber-500" />
                    ) : (
                      <div className="h-2 w-2 rounded-full bg-[#D1D1D1]" />
                    )}
                  </div>

                  <div className="flex items-start justify-between gap-4">
                    <div>
                      <div className="flex items-center gap-2">
                        {statusBadge(run.status)}
                        <span className="text-xs text-[#6B6B6B]">
                          v{run.mappingVersion?.versionNumber} · {run.direction === "excel_to_ss" ? "Excel → Smartsheet" : "Smartsheet → Excel"}
                        </span>
                      </div>
                      <div className="mt-2 flex items-center gap-4 text-xs text-[#6B6B6B]">
                        <span className="flex items-center gap-1">
                          <Clock className="h-3 w-3" />
                          Started {formatDateTime(run.createdAt)}
                        </span>
                        {run.mergedAt && <span>Merged {formatDateTime(run.mergedAt)}</span>}
                        {run.rowsChanged !== null && <span>{run.rowsChanged} rows changed</span>}
                        {run.conflictCount !== null && run.conflictCount > 0 && (
                          <span className="text-amber-600">{run.conflictCount} conflicts resolved</span>
                        )}
                      </div>
                    </div>
                    {(run.diffResult || run.status === "open") && (
                      <Button size="sm" variant="outline" onClick={() => setSelectedRun(run)}>
                        {run.status === "open" ? "Review & Merge" : "View diff"}
                      </Button>
                    )}
                  </div>
                </div>
              ))}
            </div>
          </div>
        )}
      </div>

      {selectedRun && (
        <MergeDiffModal
          run={selectedRun}
          mappingId={id}
          open={!!selectedRun}
          onClose={() => setSelectedRun(null)}
          onMerged={() => {
            setRuns((prev) =>
              prev.map((r) => r.id === selectedRun.id ? { ...r, status: "merged", mergedAt: new Date().toISOString() } : r)
            );
            setSelectedRun(null);
          }}
        />
      )}

      {selectedExcelRun && (
        <Dialog open={!!selectedExcelRun} onOpenChange={(o) => { if (!o) setSelectedExcelRun(null); }}>
          <DialogContent className="max-w-4xl max-h-[85vh] flex flex-col">
            <DialogHeader>
              <DialogTitle className="flex items-center gap-2">
                <span>Run details</span>
                <Badge variant="outline" className="text-xs">Excel → Excel</Badge>
              </DialogTitle>
            </DialogHeader>

            <div className="grid grid-cols-3 gap-2 text-xs">
              <div className="rounded-md bg-amber-50 px-2 py-1 text-amber-800">{selectedExcelRun.changedCellCount ?? 0} cells changed</div>
              <div className="rounded-md bg-blue-50 px-2 py-1 text-blue-800">{selectedExcelRun.changedRowCount ?? 0} rows impacted</div>
              <div className="rounded-md bg-emerald-50 px-2 py-1 text-emerald-800">{Array.isArray(selectedExcelRun.changedColumns) ? selectedExcelRun.changedColumns.length : 0} columns impacted</div>
            </div>

            <div className="flex-1 overflow-auto rounded-md border border-[#E5E5E5]">
              {Array.isArray(selectedExcelRun.changeSet) && selectedExcelRun.changeSet.length > 0 ? (
                <table className="w-full text-xs">
                  <thead className="sticky top-0 bg-[#F9F9F9]">
                    <tr className="border-b border-[#E5E5E5]">
                      <th className="px-3 py-2 text-left font-medium text-[#6B6B6B]">Row</th>
                      <th className="px-3 py-2 text-left font-medium text-[#6B6B6B]">Column</th>
                      <th className="px-3 py-2 text-left font-medium text-[#6B6B6B]">Before</th>
                      <th className="px-3 py-2 text-left font-medium text-[#6B6B6B]">After</th>
                    </tr>
                  </thead>
                  <tbody>
                    {selectedExcelRun.changeSet.map((cell, i) => (
                      <tr key={`${cell.row}-${cell.column}-${i}`} className="border-b border-[#F5F5F5] bg-amber-50/40">
                        <td className="px-3 py-2">{cell.row}</td>
                        <td className="px-3 py-2">{cell.column}</td>
                        <td className="px-3 py-2 font-mono text-[#6B6B6B]">{String(cell.from ?? "")}</td>
                        <td className="px-3 py-2 font-mono font-semibold">{String(cell.to ?? "")}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              ) : (
                <div className="flex h-full items-center justify-center py-12">
                  <p className="text-xs text-[#6B6B6B]">No cell-level changes were captured for this run.</p>
                </div>
              )}
            </div>
          </DialogContent>
        </Dialog>
      )}
    </AppShell>
  );
}
