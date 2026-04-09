"use client";

import { useEffect, useState } from "react";
import { useParams, useRouter } from "next/navigation";
import { AppShell } from "@/components/AppShell";
import { MapperWorkspace } from "@/components/mapper/MapperWorkspace";
import { Loader2 } from "lucide-react";

interface MappingData {
  id: string;
  name: string;
  smartsheetSheetId: string | null;
  smartsheetSheetName: string | null;
  currentVersionId: string | null;
  versions: {
    id: string;
    versionNumber: number;
    connections: unknown;
    formulas: unknown;
    schemaFingerprint: unknown;
    changeSummary: string | null;
  }[];
}

export default function EditMapperPage() {
  const { id } = useParams<{ id: string }>();
  const router = useRouter();
  const [mapping, setMapping] = useState<MappingData | null>(null);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    fetch(`/api/mappings/${id}`)
      .then((r) => r.json())
      .then((d) => { setMapping(d); setLoading(false); });
  }, [id]);

  if (loading) {
    return (
      <AppShell>
        <div className="flex h-full items-center justify-center">
          <Loader2 className="h-6 w-6 animate-spin text-[#A1A1A1]" />
        </div>
      </AppShell>
    );
  }

  if (!mapping || (mapping as { error?: string }).error) {
    return (
      <AppShell>
        <div className="flex h-full items-center justify-center">
          <p className="text-sm text-[#6B6B6B]">Mapping not found.</p>
        </div>
      </AppShell>
    );
  }

  return (
    <AppShell>
      <MapperWorkspace
        initialMapping={mapping}
        onSaved={(savedId) => router.push(`/mapper/${savedId}`)}
      />
    </AppShell>
  );
}
