"use client";

import Link from "next/link";
import { usePathname } from "next/navigation";
import { signOut, useSession } from "next-auth/react";
import { LayoutDashboard, GitMerge, BarChart2, Settings, LogOut, Zap, Menu, X, ChevronsLeft, ChevronsRight } from "lucide-react";
import { cn } from "@/lib/utils";
import { Button } from "@/components/ui/button";
import { useState } from "react";

const navItems = [
  { href: "/dashboard", label: "Dashboard", icon: LayoutDashboard },
  { href: "/mapper", label: "Mapper", icon: GitMerge },
  { href: "/dashboards", label: "Dashboards", icon: BarChart2 },
  { href: "/settings", label: "Settings", icon: Settings },
];

export function AppShell({ children }: { children: React.ReactNode }) {
  const pathname = usePathname();
  const { data: session } = useSession();
  const [mobileOpen, setMobileOpen] = useState(false);
  const [collapsed, setCollapsed] = useState(false);

  return (
    <div className="flex h-screen bg-[#F9F9F9] overflow-hidden">
      {/* Sidebar */}
      <aside
        className={cn(
          "fixed inset-y-0 left-0 z-40 flex flex-col border-r border-[#E5E5E5] bg-white transition-all duration-200",
          collapsed ? "w-16" : "w-60",
          mobileOpen ? "translate-x-0" : "-translate-x-full md:translate-x-0"
        )}
      >
        {/* Logo */}
        <div className={cn("flex h-16 items-center gap-2 border-b border-[#E5E5E5]", collapsed ? "px-4 justify-between" : "px-5")}>
          <div className="flex shrink-0 items-center gap-2">
            <div className="flex h-8 w-8 shrink-0 items-center justify-center rounded-lg bg-black">
              <Zap className="h-4 w-4 text-white" />
            </div>
            {!collapsed && <span className="text-[15px] font-semibold tracking-tight">DataWizard</span>}
          </div>
          <button
            onClick={() => setCollapsed((c) => !c)}
            className="hidden rounded p-1 hover:bg-[#F5F5F5] md:flex"
            title={collapsed ? "Expand sidebar" : "Collapse sidebar"}
          >
            {collapsed ? <ChevronsRight className="h-4 w-4 text-[#A1A1A1]" /> : <ChevronsLeft className="h-4 w-4 text-[#A1A1A1]" />}
          </button>
        </div>

        {/* Nav */}
        <nav className="flex flex-1 flex-col gap-1 p-3">
          {navItems.map(({ href, label, icon: Icon }) => {
            const active = pathname === href || pathname.startsWith(href + "/");
            return (
              <Link
                key={href}
                href={href}
                onClick={() => setMobileOpen(false)}
                title={collapsed ? label : undefined}
                className={cn(
                  "flex items-center rounded-lg px-3 py-2 text-sm font-medium transition-colors",
                  collapsed ? "justify-center gap-0" : "gap-3",
                  active
                    ? "bg-black text-white"
                    : "text-[#6B6B6B] hover:bg-[#F5F5F5] hover:text-black"
                )}
              >
                <Icon className="h-4 w-4 shrink-0" />
                {!collapsed && label}
              </Link>
            );
          })}
        </nav>

        {/* User footer */}
        <div className="border-t border-[#E5E5E5] p-3">
          <div className={cn("flex items-center rounded-lg px-3 py-2", collapsed ? "justify-center gap-0" : "gap-3")}>
            <div
              className="flex h-7 w-7 shrink-0 items-center justify-center rounded-full bg-[#1A1A1A] text-xs font-semibold text-white"
              title={collapsed ? (session?.user?.name || session?.user?.email || "User") : undefined}
            >
              {session?.user?.name?.[0]?.toUpperCase() ?? session?.user?.email?.[0]?.toUpperCase() ?? "U"}
            </div>
            {!collapsed && (
              <>
                <div className="min-w-0 flex-1">
                  <p className="truncate text-xs font-medium">{session?.user?.name || session?.user?.email}</p>
                </div>
                <Button
                  variant="ghost"
                  size="icon"
                  className="h-7 w-7 shrink-0"
                  onClick={() => signOut({ callbackUrl: "/login" })}
                >
                  <LogOut className="h-3.5 w-3.5" />
                </Button>
              </>
            )}
          </div>
        </div>
      </aside>

      {/* Mobile overlay */}
      {mobileOpen && (
        <div className="fixed inset-0 z-30 bg-black/20 md:hidden" onClick={() => setMobileOpen(false)} />
      )}

      {/* Main content */}
      <div className={cn("flex flex-1 flex-col transition-all duration-200", collapsed ? "md:pl-16" : "md:pl-60")}>
        {/* Mobile header */}
        <header className="flex h-16 items-center gap-3 border-b border-[#E5E5E5] bg-white px-4 md:hidden">
          <Button variant="ghost" size="icon" onClick={() => setMobileOpen(true)}>
            {mobileOpen ? <X className="h-5 w-5" /> : <Menu className="h-5 w-5" />}
          </Button>
          <span className="text-sm font-semibold">DataWizard</span>
        </header>

        <main className="flex-1 overflow-auto">{children}</main>
      </div>
    </div>
  );
}
