"use client";

import { useEffect, useState } from "react";
import Link from "next/link";
import { Plus, GitMerge, Clock, Play, Pencil, History, Share2, Trash2, Zap } from "lucide-react";
import { AppShell } from "@/components/AppShell";
import { Button } from "@/components/ui/button";
import { Badge } from "@/components/ui/badge";
import { formatDate } from "@/lib/utils";
import { useToast } from "@/hooks/use-toast";
import { RunMappingDialog } from "@/components/mapper/RunMappingDialog";

export interface MappingNode {
  id: string;
  type: string;
  data: {
    label?: string;
    sourceField?: string;
    colKey?: string;
    colId?: string;
    colRef?: string;
    formula?: string;
    leftInputs?: { id: string; label: string; colRef?: string }[];
    rightInputs?: { id: string; label: string; colRef?: string }[];
  };
}
export interface MappingEdge {
  id: string;
  source: string;
  target: string;
  data?: { formula?: string };
}
export interface MappingConnections {
  nodes?: MappingNode[];
  edges?: MappingEdge[];
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
    };
    dynamicTargetColumn?: {
      enabled?: boolean;
      sourceLabel?: string;
      sourceNodeId?: string;
      nameTemplate?: string;
      columnPosition?: "start" | "end" | "custom";
      customColumnNumber?: number;
    };
  };
}

interface MappingItem {
  id: string;
  name: string;
  slug: string;
  autoPush?: boolean;
  updatedAt: string;
  createdAt: string;
  smartsheetSheetId: string | null;
  smartsheetSheetName: string | null;
  currentVersionId: string | null;
  versions: {
    id: string;
    versionNumber: number;
    changeSummary: string | null;
    connections?: MappingConnections;
    formulas?: Record<string, string>;
    schemaFingerprint?: unknown;
  }[];
  stagingRuns: { id: string; status: string; createdAt: string }[];
  mappingRunCount?: number;
}

export default function MapperPage() {
  const { toast } = useToast();
  const [mappings, setMappings] = useState<MappingItem[]>([]);
  const [loading, setLoading] = useState(true);
  const [runTarget, setRunTarget] = useState<MappingItem | null>(null);

  useEffect(() => {
    fetch("/api/mappings")
      .then((r) => r.json())
      .then((d) => {
        setMappings(
          Array.isArray(d)
            ? d.map((m) => ({
                ...m,
                autoPush: typeof m.autoPush === "boolean" ? m.autoPush : false,
              }))
            : []
        );
        setLoading(false);
      });
  }, []);

  async function deleteMapping(id: string, name: string) {
    if (!confirm(`Delete mapping "${name}"? This cannot be undone.`)) return;
    const res = await fetch(`/api/mappings/${id}`, { method: "DELETE" });
    if (!res.ok) {
      const err = await res.json().catch(() => ({ error: "Delete failed" }));
      toast({ title: "Delete blocked", description: err.error || "Delete failed", variant: "destructive" });
      return;
    }
    setMappings((prev) => prev.filter((m) => m.id !== id));
    toast({ title: "Mapping deleted" });
  }

  function copyShareUrl(slug: string) {
    const url = `${window.location.origin}/share/${slug}?type=mapping`;
    navigator.clipboard.writeText(url);
    toast({ title: "Share link copied!" });
  }

  return (
    <AppShell>
      <div className="p-6 lg:p-8">
        <div className="mb-6 flex items-center justify-between">
          <div>
            <h1 className="text-xl font-semibold tracking-tight">Mappings</h1>
            <p className="mt-0.5 text-sm text-[#6B6B6B]">Excel → Smartsheet automation templates</p>
          </div>
          <Button asChild>
            <Link href="/mapper/new">
              <Plus className="mr-1.5 h-4 w-4" />
              New mapping
            </Link>
          </Button>
        </div>

        {loading ? (
          <div className="space-y-3">
            {[1, 2, 3].map((i) => <div key={i} className="h-20 animate-pulse rounded-xl bg-[#F5F5F5]" />)}
          </div>
        ) : mappings.length === 0 ? (
          <div className="flex flex-col items-center justify-center rounded-xl border border-dashed border-[#E5E5E5] bg-white py-20">
            <GitMerge className="mb-4 h-10 w-10 text-[#D1D1D1]" />
            <p className="text-sm font-medium">No mappings yet</p>
            <p className="mt-1 text-xs text-[#6B6B6B]">Connect Excel columns to Smartsheet — once, reuse forever</p>
            <Button asChild className="mt-5">
              <Link href="/mapper/new">Create your first mapping</Link>
            </Button>
          </div>
        ) : (
          <div className="space-y-3">
            {mappings.map((m) => {
              const currentVersion = m.versions.find((v) => v.id === m.currentVersionId) || m.versions[m.versions.length - 1];
              const openRun = m.stagingRuns.find((r) => r.status === "open");
              return (
                <div
                  key={m.id}
                  className="flex items-center justify-between gap-4 rounded-xl border border-[#E5E5E5] bg-white px-5 py-4 shadow-sm"
                >
                  <div className="min-w-0 flex-1">
                    <div className="flex items-center gap-2">
                      <span className="font-medium">{m.name}</span>
                      {currentVersion && (
                        <Badge variant="outline" className="text-xs">
                          v{currentVersion.versionNumber}
                        </Badge>
                      )}
                      {openRun && (
                        <Badge variant="secondary" className="text-xs">
                          <Zap className="mr-1 h-2.5 w-2.5" />
                          Run open
                        </Badge>
                      )}
                    </div>
                    <div className="mt-1 flex items-center gap-3 text-xs text-[#6B6B6B]">
                      {m.smartsheetSheetId ? (
                        <span>→ {m.smartsheetSheetName ?? "Smartsheet"}</span>
                      ) : (
                        <span className="text-blue-600">Excel → Excel</span>
                      )}
                      {m.autoPush && <span className="text-emerald-700">Auto Push</span>}
                      <span className="flex items-center gap-1">
                        <Clock className="h-3 w-3" />
                        Updated {formatDate(m.updatedAt)}
                      </span>
                      <span>{m.stagingRuns.length + (m.mappingRunCount ?? 0)} runs</span>
                    </div>
                  </div>
                  <div className="flex shrink-0 items-center gap-1.5">
                    <Button size="sm" variant="outline" onClick={() => setRunTarget(m)}>
                      <Play className="mr-1.5 h-3.5 w-3.5" />
                      Run
                    </Button>
                    <Button asChild size="sm" variant="outline">
                      <Link href={`/mapper/${m.id}`}>
                        <Pencil className="mr-1.5 h-3.5 w-3.5" />
                        Edit
                      </Link>
                    </Button>
                    <Button asChild size="sm" variant="ghost">
                      <Link href={`/mapper/${m.id}/history`}>
                        <History className="h-3.5 w-3.5" />
                      </Link>
                    </Button>
                    <Button size="sm" variant="ghost" onClick={() => copyShareUrl(m.slug)}>
                      <Share2 className="h-3.5 w-3.5" />
                    </Button>
                    <Button
                      size="sm"
                      variant="ghost"
                      className="text-red-500 hover:text-red-600"
                      onClick={() => deleteMapping(m.id, m.name)}
                    >
                      <Trash2 className="h-3.5 w-3.5" />
                    </Button>
                  </div>
                </div>
              );
            })}
          </div>
        )}
      </div>

      {runTarget && (
        <RunMappingDialog
          mapping={runTarget}
          open={!!runTarget}
          onClose={() => setRunTarget(null)}
        />
      )}
    </AppShell>
  );
}
