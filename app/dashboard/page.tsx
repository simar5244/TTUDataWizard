"use client";

import { useEffect, useState } from "react";
import Link from "next/link";
import { useSession } from "next-auth/react";
import { GitMerge, BarChart2, Plus, ArrowRight, Clock, Zap } from "lucide-react";
import { AppShell } from "@/components/AppShell";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle, CardDescription } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import { formatDate } from "@/lib/utils";

interface MappingItem {
  id: string;
  name: string;
  updatedAt: string;
  smartsheetSheetName: string | null;
  stagingRuns: { status: string; createdAt: string }[];
}

interface DashboardItem {
  id: string;
  name: string;
  updatedAt: string;
  charts: unknown[];
}

export default function DashboardPage() {
  const { data: session } = useSession();
  const [mappings, setMappings] = useState<MappingItem[]>([]);
  const [dashboards, setDashboards] = useState<DashboardItem[]>([]);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    Promise.all([
      fetch("/api/mappings").then((r) => r.json()),
      fetch("/api/dashboards").then((r) => r.json()),
    ]).then(([m, d]) => {
      setMappings(Array.isArray(m) ? m : []);
      setDashboards(Array.isArray(d) ? d : []);
      setLoading(false);
    });
  }, []);

  const firstName = session?.user?.name?.split(" ")[0] ?? "there";

  return (
    <AppShell>
      <div className="p-6 lg:p-8">
        {/* Header */}
        <div className="mb-8">
          <h1 className="text-2xl font-semibold tracking-tight">Good morning, {firstName}</h1>
          <p className="mt-1 text-sm text-[#6B6B6B]">Here&apos;s what&apos;s happening with your data automations.</p>
        </div>

        {/* Stats row */}
        <div className="mb-8 grid grid-cols-2 gap-4 sm:grid-cols-4">
          {[
            { label: "Mappings", value: mappings.length, icon: GitMerge },
            { label: "Dashboards", value: dashboards.length, icon: BarChart2 },
            {
              label: "Active Runs",
              value: mappings.flatMap((m) => m.stagingRuns).filter((r) => r.status === "open").length,
              icon: Zap,
            },
            {
              label: "Total Charts",
              value: dashboards.reduce((a, d) => a + (Array.isArray(d.charts) ? d.charts.length : 0), 0),
              icon: BarChart2,
            },
          ].map(({ label, value, icon: Icon }) => (
            <div key={label} className="rounded-xl border border-[#E5E5E5] bg-white p-4 shadow-sm">
              <div className="flex items-center justify-between">
                <p className="text-sm text-[#6B6B6B]">{label}</p>
                <Icon className="h-4 w-4 text-[#A1A1A1]" />
              </div>
              <p className="mt-2 text-2xl font-semibold">{loading ? "—" : value}</p>
            </div>
          ))}
        </div>

        {/* Mappings */}
        <div className="mb-8">
          <div className="mb-4 flex items-center justify-between">
            <h2 className="text-base font-semibold">Recent Mappings</h2>
            <div className="flex gap-2">
              <Button asChild size="sm" variant="outline">
                <Link href="/mapper">View all</Link>
              </Button>
              <Button asChild size="sm">
                <Link href="/mapper/new">
                  <Plus className="mr-1.5 h-3.5 w-3.5" />
                  New mapping
                </Link>
              </Button>
            </div>
          </div>

          {loading ? (
            <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-3">
              {[1, 2, 3].map((i) => (
                <div key={i} className="h-28 animate-pulse rounded-xl bg-[#F5F5F5]" />
              ))}
            </div>
          ) : mappings.length === 0 ? (
            <div className="flex flex-col items-center justify-center rounded-xl border border-dashed border-[#E5E5E5] bg-white py-12">
              <GitMerge className="mb-3 h-8 w-8 text-[#D1D1D1]" />
              <p className="text-sm font-medium">No mappings yet</p>
              <p className="mt-1 text-xs text-[#6B6B6B]">Create your first Excel → Smartsheet mapping</p>
              <Button asChild className="mt-4" size="sm">
                <Link href="/mapper/new">Create mapping</Link>
              </Button>
            </div>
          ) : (
            <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-3">
              {mappings.slice(0, 6).map((m) => (
                <Link key={m.id} href={`/mapper/${m.id}`}>
                  <Card className="cursor-pointer transition-shadow hover:shadow-md">
                    <CardHeader className="pb-2">
                      <div className="flex items-start justify-between gap-2">
                        <CardTitle className="text-sm font-medium">{m.name}</CardTitle>
                        {m.stagingRuns[0]?.status === "open" && (
                          <Badge variant="secondary" className="shrink-0 text-xs">
                            Open run
                          </Badge>
                        )}
                      </div>
                      {m.smartsheetSheetName && (
                        <CardDescription className="text-xs">→ {m.smartsheetSheetName}</CardDescription>
                      )}
                    </CardHeader>
                    <CardContent>
                      <div className="flex items-center justify-between">
                        <span className="flex items-center gap-1 text-xs text-[#6B6B6B]">
                          <Clock className="h-3 w-3" />
                          {formatDate(m.updatedAt)}
                        </span>
                        <ArrowRight className="h-3.5 w-3.5 text-[#A1A1A1]" />
                      </div>
                    </CardContent>
                  </Card>
                </Link>
              ))}
            </div>
          )}
        </div>

        {/* Dashboards */}
        <div>
          <div className="mb-4 flex items-center justify-between">
            <h2 className="text-base font-semibold">Recent Dashboards</h2>
            <div className="flex gap-2">
              <Button asChild size="sm" variant="outline">
                <Link href="/dashboards">View all</Link>
              </Button>
              <Button asChild size="sm">
                <Link href="/dashboards/new">
                  <Plus className="mr-1.5 h-3.5 w-3.5" />
                  New dashboard
                </Link>
              </Button>
            </div>
          </div>

          {loading ? (
            <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-3">
              {[1, 2].map((i) => (
                <div key={i} className="h-28 animate-pulse rounded-xl bg-[#F5F5F5]" />
              ))}
            </div>
          ) : dashboards.length === 0 ? (
            <div className="flex flex-col items-center justify-center rounded-xl border border-dashed border-[#E5E5E5] bg-white py-12">
              <BarChart2 className="mb-3 h-8 w-8 text-[#D1D1D1]" />
              <p className="text-sm font-medium">No dashboards yet</p>
              <p className="mt-1 text-xs text-[#6B6B6B]">Build charts directly from your Excel data</p>
              <Button asChild className="mt-4" size="sm">
                <Link href="/dashboards/new">Create dashboard</Link>
              </Button>
            </div>
          ) : (
            <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-3">
              {dashboards.slice(0, 6).map((d) => (
                <Link key={d.id} href={`/dashboards/${d.id}`}>
                  <Card className="cursor-pointer transition-shadow hover:shadow-md">
                    <CardHeader className="pb-2">
                      <CardTitle className="text-sm font-medium">{d.name}</CardTitle>
                      <CardDescription className="text-xs">
                        {Array.isArray(d.charts) ? d.charts.length : 0} charts
                      </CardDescription>
                    </CardHeader>
                    <CardContent>
                      <div className="flex items-center justify-between">
                        <span className="flex items-center gap-1 text-xs text-[#6B6B6B]">
                          <Clock className="h-3 w-3" />
                          {formatDate(d.updatedAt)}
                        </span>
                        <ArrowRight className="h-3.5 w-3.5 text-[#A1A1A1]" />
                      </div>
                    </CardContent>
                  </Card>
                </Link>
              ))}
            </div>
          )}
        </div>
      </div>
    </AppShell>
  );
}
