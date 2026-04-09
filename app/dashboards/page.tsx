"use client";

import { useEffect, useState } from "react";
import Link from "next/link";
import { Plus, BarChart2, Clock, Trash2, Share2, Pencil } from "lucide-react";
import { AppShell } from "@/components/AppShell";
import { Button } from "@/components/ui/button";
import { formatDate } from "@/lib/utils";
import { useToast } from "@/hooks/use-toast";

interface DashboardItem {
  id: string;
  name: string;
  slug: string;
  updatedAt: string;
  charts: unknown[];
}

export default function DashboardsPage() {
  const { toast } = useToast();
  const [dashboards, setDashboards] = useState<DashboardItem[]>([]);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    fetch("/api/dashboards")
      .then((r) => r.json())
      .then((d) => { setDashboards(Array.isArray(d) ? d : []); setLoading(false); });
  }, []);

  async function deleteDashboard(id: string, name: string) {
    if (!confirm(`Delete dashboard "${name}"?`)) return;
    const res = await fetch(`/api/dashboards/${id}`, { method: "DELETE" });
    if (!res.ok) {
      const err = await res.json().catch(() => ({ error: "Delete failed" }));
      toast({ title: "Delete blocked", description: err.error || "Delete failed", variant: "destructive" });
      return;
    }
    setDashboards((prev) => prev.filter((d) => d.id !== id));
    toast({ title: "Dashboard deleted" });
  }

  function copyShareUrl(slug: string) {
    navigator.clipboard.writeText(`${window.location.origin}/share/${slug}?type=dashboard`);
    toast({ title: "Share link copied!" });
  }

  return (
    <AppShell>
      <div className="p-6 lg:p-8">
        <div className="mb-6 flex items-center justify-between">
          <div>
            <h1 className="text-xl font-semibold tracking-tight">Dashboards</h1>
            <p className="mt-0.5 text-sm text-[#6B6B6B]">Charts and visualizations from your Excel data</p>
          </div>
          <Button asChild>
            <Link href="/dashboards/new">
              <Plus className="mr-1.5 h-4 w-4" />New dashboard
            </Link>
          </Button>
        </div>

        {loading ? (
          <div className="grid gap-4 sm:grid-cols-2 lg:grid-cols-3">
            {[1, 2, 3].map((i) => <div key={i} className="h-32 animate-pulse rounded-xl bg-[#F5F5F5]" />)}
          </div>
        ) : dashboards.length === 0 ? (
          <div className="flex flex-col items-center justify-center rounded-xl border border-dashed border-[#E5E5E5] bg-white py-20">
            <BarChart2 className="mb-4 h-10 w-10 text-[#D1D1D1]" />
            <p className="text-sm font-medium">No dashboards yet</p>
            <p className="mt-1 text-xs text-[#6B6B6B]">Upload Excel data and build charts in minutes</p>
            <Button asChild className="mt-5">
              <Link href="/dashboards/new">Create your first dashboard</Link>
            </Button>
          </div>
        ) : (
          <div className="grid gap-4 sm:grid-cols-2 lg:grid-cols-3">
            {dashboards.map((d) => (
              <div key={d.id} className="group rounded-xl border border-[#E5E5E5] bg-white p-5 shadow-sm transition-all hover:shadow-md">
                <div className="mb-3 flex items-start justify-between gap-2">
                  <h3 className="font-medium leading-tight">{d.name}</h3>
                  <span className="shrink-0 rounded-full bg-[#F5F5F5] px-2 py-0.5 text-xs text-[#6B6B6B]">
                    {Array.isArray(d.charts) ? d.charts.length : 0} charts
                  </span>
                </div>
                <p className="flex items-center gap-1 text-xs text-[#A1A1A1]">
                  <Clock className="h-3 w-3" />
                  {formatDate(d.updatedAt)}
                </p>
                <div className="mt-4 flex gap-1.5 opacity-0 transition-opacity group-hover:opacity-100">
                  <Button asChild size="sm" className="flex-1">
                    <Link href={`/dashboards/${d.id}`}>
                      <Pencil className="mr-1.5 h-3.5 w-3.5" />Open
                    </Link>
                  </Button>
                  <Button size="sm" variant="outline" onClick={() => copyShareUrl(d.slug)}>
                    <Share2 className="h-3.5 w-3.5" />
                  </Button>
                  <Button size="sm" variant="ghost" className="text-red-500 hover:text-red-600" onClick={() => deleteDashboard(d.id, d.name)}>
                    <Trash2 className="h-3.5 w-3.5" />
                  </Button>
                </div>
              </div>
            ))}
          </div>
        )}
      </div>
    </AppShell>
  );
}
