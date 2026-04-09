"use client";

import { useRouter } from "next/navigation";
import { ReactFlowProvider } from "@xyflow/react";
import { AppShell } from "@/components/AppShell";
import { MapperWorkspace } from "@/components/mapper/MapperWorkspace";

export default function CreateMapperPage() {
  const router = useRouter();

  function handleSaved(id: string) {
    router.push(`/mapper/${id}`);
  }

  return (
    <AppShell>
      <ReactFlowProvider>
        <MapperWorkspace onSaved={handleSaved} />
      </ReactFlowProvider>
    </AppShell>
  );
}
