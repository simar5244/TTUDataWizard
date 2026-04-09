"use client";

import { useEffect, useState } from "react";
import { useParams } from "next/navigation";
import Link from "next/link";
import { ArrowLeft, Clock, RotateCcw, Loader2, GitCommit, CheckCircle2 } from "lucide-react";
import { AppShell } from "@/components/AppShell";
import { Button } from "@/components/ui/button";
import { Badge } from "@/components/ui/badge";
import { formatDateTime } from "@/lib/utils";
import { useToast } from "@/hooks/use-toast";

interface MappingVersion {
  id: string;
  versionNumber: number;
  changeSummary: string | null;
  createdAt: string;
  user: { name: string | null; email: string };
}

interface MappingData {
  id: string;
  name: string;
  currentVersionId: string | null;
}

export default function VersionsPage() {
  const { id } = useParams<{ id: string }>();
  const { toast } = useToast();
  const [mapping, setMapping] = useState<MappingData | null>(null);
  const [versions, setVersions] = useState<MappingVersion[]>([]);
  const [loading, setLoading] = useState(true);
  const [settingVersion, setSettingVersion] = useState<string | null>(null);

  useEffect(() => {
    Promise.all([
      fetch(`/api/mappings/${id}`).then((r) => r.json()),
      fetch(`/api/mappings/${id}/versions`).then((r) => r.json()),
    ]).then(([m, v]) => {
      setMapping(m);
      setVersions(Array.isArray(v) ? v : []);
      setLoading(false);
    });
  }, [id]);

  async function setCurrentVersion(versionId: string) {
    setSettingVersion(versionId);
    try {
      const res = await fetch(`/api/mappings/${id}/versions`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ versionId }),
      });
      if (!res.ok) throw new Error((await res.json()).error);
      setMapping((prev) => prev ? { ...prev, currentVersionId: versionId } : prev);
      toast({ title: "Version restored" });
    } catch (e) {
      toast({ title: "Failed to restore version", description: (e as Error).message, variant: "destructive" });
    } finally {
      setSettingVersion(null);
    }
  }

  return (
    <AppShell>
      <div className="p-6 lg:p-8">
        <div className="mb-6 flex items-center gap-3">
          <Button asChild variant="ghost" size="sm">
            <Link href={`/mapper/${id}`}>
              <ArrowLeft className="mr-1.5 h-4 w-4" />Back
            </Link>
          </Button>
          <div>
            <h1 className="text-xl font-semibold">{mapping?.name || "…"}</h1>
            <p className="text-sm text-[#6B6B6B]">Version history</p>
          </div>
        </div>

        {loading ? (
          <div className="flex justify-center py-20">
            <Loader2 className="h-6 w-6 animate-spin text-[#A1A1A1]" />
          </div>
        ) : (
          <div className="relative">
            <div className="absolute left-4 top-0 h-full w-0.5 bg-[#E5E5E5]" />
            <div className="space-y-3 pl-12">
              {versions.map((v) => {
                const isCurrent = v.id === mapping?.currentVersionId;
                return (
                  <div
                    key={v.id}
                    className={`relative rounded-xl border bg-white p-5 shadow-sm transition-all ${
                      isCurrent ? "border-black" : "border-[#E5E5E5]"
                    }`}
                  >
                    {/* Timeline dot */}
                    <div className="absolute -left-9 top-5 flex h-5 w-5 items-center justify-center rounded-full border-2 border-white bg-white shadow">
                      {isCurrent ? (
                        <div className="h-3 w-3 rounded-full bg-black" />
                      ) : (
                        <GitCommit className="h-3.5 w-3.5 text-[#A1A1A1]" />
                      )}
                    </div>

                    <div className="flex items-start justify-between gap-4">
                      <div>
                        <div className="flex items-center gap-2">
                          <span className="font-semibold">v{v.versionNumber}</span>
                          {isCurrent && (
                            <Badge className="text-xs">
                              <CheckCircle2 className="mr-1 h-3 w-3" />Current
                            </Badge>
                          )}
                        </div>
                        <p className="mt-1 text-sm text-[#6B6B6B]">
                          {v.changeSummary || "No description"}
                        </p>
                        <div className="mt-2 flex items-center gap-3 text-xs text-[#A1A1A1]">
                          <span className="flex items-center gap-1">
                            <Clock className="h-3 w-3" />
                            {formatDateTime(v.createdAt)}
                          </span>
                          {v.user && (
                            <span>by {v.user.name || v.user.email}</span>
                          )}
                        </div>
                      </div>
                      {!isCurrent && (
                        <Button
                          size="sm"
                          variant="outline"
                          disabled={settingVersion === v.id}
                          onClick={() => setCurrentVersion(v.id)}
                        >
                          {settingVersion === v.id ? (
                            <Loader2 className="mr-1.5 h-3.5 w-3.5 animate-spin" />
                          ) : (
                            <RotateCcw className="mr-1.5 h-3.5 w-3.5" />
                          )}
                          Restore
                        </Button>
                      )}
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        )}
      </div>
    </AppShell>
  );
}
