"use client";

import { useEffect, useState } from "react";
import { useParams, useRouter } from "next/navigation";
import { useSession } from "next-auth/react";
import { Loader2, AlertCircle, Pencil, ExternalLink } from "lucide-react";
import { DashboardBuilder } from "@/components/dashboard/DashboardBuilder";
import { Button } from "@/components/ui/button";

interface DashboardData {
  id: string;
  name: string;
  slug: string;
  userId: string;
  excelData: unknown;
  charts: unknown[];
  layout: unknown[];
}

export default function SharePage() {
  const { slug } = useParams<{ slug: string }>();
  const router = useRouter();
  const { data: session } = useSession();
  const [type, setType] = useState<"dashboard" | "mapping" | null>(null);
  const [dashboard, setDashboard] = useState<DashboardData | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    fetch(`/api/share/${slug}?type=dashboard`)
      .then(async (r) => {
        if (!r.ok) {
          const e = await r.json();
          setError(e.error ?? "Not found");
          setLoading(false);
          return;
        }
        const d = await r.json();
        setType(d.type);
        if (d.type === "dashboard") setDashboard(d.data);
        setLoading(false);
      })
      .catch(() => { setError("Failed to load"); setLoading(false); });
  }, [slug]);

  const currentUserId = (session?.user as { id?: string })?.id;
  const isOwner = dashboard && currentUserId && dashboard.userId === currentUserId;

  if (loading) {
    return (
      <div className="flex h-screen items-center justify-center bg-[#F9F9F9]">
        <Loader2 className="h-6 w-6 animate-spin text-[#A1A1A1]" />
      </div>
    );
  }

  if (error) {
    return (
      <div className="flex h-screen flex-col items-center justify-center gap-3 bg-[#F9F9F9]">
        <AlertCircle className="h-10 w-10 text-[#D1D1D1]" />
        <p className="text-sm font-medium text-[#6B6B6B]">{error}</p>
        <Button variant="outline" size="sm" onClick={() => router.push("/")}>
          Go home
        </Button>
      </div>
    );
  }

  if (type === "dashboard" && dashboard) {
    return (
      <div className="flex h-screen flex-col bg-white">
        {/* Minimal header bar */}
        <div className="flex shrink-0 items-center gap-3 border-b border-[#E5E5E5] bg-white px-5 py-2">
          <div className="flex items-center gap-2">
            <div className="flex h-6 w-6 items-center justify-center rounded bg-black">
              <ExternalLink className="h-3.5 w-3.5 text-white" />
            </div>
            <span className="text-sm font-semibold">DataWizard</span>
          </div>
          <span className="text-xs text-[#A1A1A1]">·</span>
          <span className="text-xs text-[#6B6B6B]">Shared dashboard</span>
          <div className="ml-auto flex items-center gap-2">
            {isOwner && (
              <Button
                size="sm"
                variant="outline"
                onClick={() => router.push(`/dashboards/${dashboard.id}`)}
              >
                <Pencil className="mr-1.5 h-3.5 w-3.5" />
                Edit dashboard
              </Button>
            )}
            {!session && (
              <Button size="sm" variant="ghost" onClick={() => router.push("/login")}>
                Sign in to edit
              </Button>
            )}
          </div>
        </div>

        {/* Dashboard in view-only mode — no AppShell wrapper */}
        <div className="flex-1 overflow-hidden">
          <DashboardBuilder
            initialDashboard={dashboard}
            onSaved={() => {}}
            initialViewMode={true}
          />
        </div>
      </div>
    );
  }

  return (
    <div className="flex h-screen items-center justify-center bg-[#F9F9F9]">
      <p className="text-sm text-[#6B6B6B]">This content type cannot be displayed.</p>
    </div>
  );
}
