"use client";

import { useRouter } from "next/navigation";
import { AppShell } from "@/components/AppShell";
import { DashboardBuilder } from "@/components/dashboard/DashboardBuilder";

export default function NewDashboardPage() {
  const router = useRouter();
  return (
    <AppShell>
      <DashboardBuilder onSaved={(id) => router.push(`/dashboards/${id}`)} />
    </AppShell>
  );
}
