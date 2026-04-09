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

export default function MappingHistoryPage() {
  const { id } = useParams<{ id: string }>();
  const [runs, setRuns] = useState<StagingRun[]>([]);
  const [mappingName, setMappingName] = useState("");
  const [loading, setLoading] = useState(true);
  const [selectedRun, setSelectedRun] = useState<StagingRun | null>(null);

  useEffect(() => {
    Promise.all([
      fetch(`/api/mappings/${id}`).then((r) => r.json()),
      fetch(`/api/mappings/${id}/staging`).then((r) => r.json()),
    ]).then(([m, r]) => {
      setMappingName(m.name || "");
      setRuns(Array.isArray(r) ? r : []);
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
        ) : runs.length === 0 ? (
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
    </AppShell>
  );
}
