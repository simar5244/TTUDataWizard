"use client";

import { useEffect, useState } from "react";
import { useParams, useRouter } from "next/navigation";
import { AppShell } from "@/components/AppShell";
import { DashboardBuilder, type DataSource } from "@/components/dashboard/DashboardBuilder";
import { Loader2 } from "lucide-react";

interface DashboardData {
  id: string;
  name: string;
  slug: string;
  userId: string;
  excelData: unknown;
  dataSources: DataSource[];
  charts: unknown[];
  layout: unknown[];
}

export default function EditDashboardPage() {
  const { id } = useParams<{ id: string }>();
  const router = useRouter();
  const [dashboard, setDashboard] = useState<DashboardData | null>(null);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    fetch(`/api/dashboards/${id}`)
      .then((r) => r.json())
      .then((d) => { setDashboard(d); setLoading(false); });
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

  return (
    <AppShell>
      <DashboardBuilder
        initialDashboard={dashboard || undefined}
        onSaved={(savedId: string) => router.push(`/dashboards/${savedId}`)}
      />
    </AppShell>
  );
}
