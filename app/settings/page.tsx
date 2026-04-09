"use client";

import { useState, useEffect } from "react";
import { CheckCircle2, AlertCircle, Loader2, Trash2, Link2 } from "lucide-react";
import { AppShell } from "@/components/AppShell";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Switch } from "@/components/ui/switch";
import { Card, CardContent, CardHeader, CardTitle, CardDescription } from "@/components/ui/card";
import { Separator } from "@/components/ui/separator";
import { Badge } from "@/components/ui/badge";
import { useToast } from "@/hooks/use-toast";

export default function SettingsPage() {
  const { toast } = useToast();
  const [ssToken, setSsToken] = useState("");
  const [ssConnected, setSsConnected] = useState(false);
  const [ssUser, setSsUser] = useState<{ name?: string; email?: string } | null>(null);
  const [ssLoading, setSsLoading] = useState(false);
  const [checkingStatus, setCheckingStatus] = useState(true);
  const [allowEdits, setAllowEdits] = useState(true);
  const [allowDeletes, setAllowDeletes] = useState(true);
  const [securitySaving, setSecuritySaving] = useState(false);

  useEffect(() => {
    fetch("/api/settings/smartsheet")
      .then((r) => r.json())
      .then((d) => {
        setSsConnected(d.connected);
        setAllowEdits(typeof d.allowEdits === "boolean" ? d.allowEdits : true);
        setAllowDeletes(typeof d.allowDeletes === "boolean" ? d.allowDeletes : true);
        setCheckingStatus(false);
      });
  }, []);

  async function updateSecurity(next: { allowEdits?: boolean; allowDeletes?: boolean }) {
    setSecuritySaving(true);
    try {
      const res = await fetch("/api/settings/smartsheet", {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(next),
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || "Failed to update security settings");
      setAllowEdits(typeof data.allowEdits === "boolean" ? data.allowEdits : allowEdits);
      setAllowDeletes(typeof data.allowDeletes === "boolean" ? data.allowDeletes : allowDeletes);
      toast({ title: "Security settings updated" });
    } catch (e) {
      toast({ title: "Update failed", description: (e as Error).message, variant: "destructive" });
    } finally {
      setSecuritySaving(false);
    }
  }

  async function connectSmartsheet() {
    if (!ssToken.trim()) return;
    setSsLoading(true);
    try {
      const res = await fetch("/api/settings/smartsheet", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ token: ssToken }),
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error);
      setSsConnected(true);
      setSsUser({ name: data.name, email: data.email });
      setSsToken("");
      toast({ title: "Smartsheet connected!", description: `Connected as ${data.name || data.email}` });
    } catch (e) {
      toast({ title: "Connection failed", description: (e as Error).message, variant: "destructive" });
    } finally {
      setSsLoading(false);
    }
  }

  async function disconnectSmartsheet() {
    if (!confirm("Disconnect Smartsheet? Existing mappings will keep their configuration.")) return;
    await fetch("/api/settings/smartsheet", { method: "DELETE" });
    setSsConnected(false);
    setSsUser(null);
    toast({ title: "Smartsheet disconnected" });
  }

  return (
    <AppShell>
      <div className="mx-auto max-w-2xl p-6 lg:p-8">
        <div className="mb-8">
          <h1 className="text-xl font-semibold tracking-tight">Settings</h1>
          <p className="mt-0.5 text-sm text-[#6B6B6B]">Manage your account and integrations</p>
        </div>

        {/* Smartsheet Integration */}
        <Card>
          <CardHeader>
            <div className="flex items-center justify-between">
              <div>
                <CardTitle className="text-base">Smartsheet Integration</CardTitle>
                <CardDescription className="mt-1">
                  Connect your Smartsheet account to enable data syncing
                </CardDescription>
              </div>
              {checkingStatus ? (
                <Loader2 className="h-4 w-4 animate-spin text-[#A1A1A1]" />
              ) : ssConnected ? (
                <Badge className="bg-green-100 text-green-800 border-green-200">
                  <CheckCircle2 className="mr-1 h-3 w-3" />Connected
                </Badge>
              ) : (
                <Badge variant="outline" className="text-[#6B6B6B]">
                  <AlertCircle className="mr-1 h-3 w-3" />Not connected
                </Badge>
              )}
            </div>
          </CardHeader>
          <CardContent className="space-y-4">
            {ssConnected ? (
              <div className="space-y-4">
                {ssUser && (
                  <div className="rounded-lg border border-[#E5E5E5] bg-[#F9F9F9] px-4 py-3">
                    <p className="text-sm font-medium">{ssUser.name}</p>
                    {ssUser.email && <p className="text-xs text-[#6B6B6B]">{ssUser.email}</p>}
                  </div>
                )}
                <Button variant="outline" size="sm" className="text-red-500 hover:text-red-600 border-red-200" onClick={disconnectSmartsheet}>
                  <Trash2 className="mr-1.5 h-3.5 w-3.5" />
                  Disconnect
                </Button>
              </div>
            ) : (
              <div className="space-y-4">
                <div className="rounded-lg border border-[#E5E5E5] bg-[#F5F5F5] p-4 text-sm text-[#6B6B6B]">
                  <p className="font-medium text-black">How to get your API token:</p>
                  <ol className="mt-2 space-y-1 text-xs">
                    <li>1. Log in to <a href="https://app.smartsheet.com" target="_blank" className="underline">Smartsheet</a></li>
                    <li>2. Click your profile icon → Account → Personal Settings</li>
                    <li>3. Go to API Access → Generate new access token</li>
                    <li>4. Copy and paste the token below</li>
                  </ol>
                </div>
                <div className="space-y-2">
                  <Label>API Token</Label>
                  <div className="flex gap-2">
                    <Input
                      type="password"
                      placeholder="Paste your Smartsheet API token…"
                      value={ssToken}
                      onChange={(e) => setSsToken(e.target.value)}
                      onKeyDown={(e) => e.key === "Enter" && connectSmartsheet()}
                    />
                    <Button onClick={connectSmartsheet} disabled={ssLoading || !ssToken.trim()}>
                      {ssLoading ? <Loader2 className="mr-1.5 h-4 w-4 animate-spin" /> : <Link2 className="mr-1.5 h-4 w-4" />}
                      Connect
                    </Button>
                  </div>
                </div>
              </div>
            )}
          </CardContent>
        </Card>

        <Separator className="my-6" />

        <Card>
          <CardHeader>
            <CardTitle className="text-base">Security Controls</CardTitle>
            <CardDescription className="mt-1">
              Hard-stop controls for data operations across the app
            </CardDescription>
          </CardHeader>
          <CardContent className="space-y-5">
            <div className="flex items-center justify-between rounded-lg border border-[#E5E5E5] px-4 py-3">
              <div>
                <p className="text-sm font-medium">Allow edits</p>
                <p className="text-xs text-[#6B6B6B]">If off, create/update/run actions are blocked system-wide.</p>
              </div>
              <Switch
                checked={allowEdits}
                disabled={securitySaving}
                onCheckedChange={(checked) => {
                  setAllowEdits(checked);
                  void updateSecurity({ allowEdits: checked });
                }}
              />
            </div>
            <div className="flex items-center justify-between rounded-lg border border-[#E5E5E5] px-4 py-3">
              <div>
                <p className="text-sm font-medium">Allow deletes</p>
                <p className="text-xs text-[#6B6B6B]">If off, delete actions are blocked system-wide.</p>
              </div>
              <Switch
                checked={allowDeletes}
                disabled={securitySaving}
                onCheckedChange={(checked) => {
                  setAllowDeletes(checked);
                  void updateSecurity({ allowDeletes: checked });
                }}
              />
            </div>

          </CardContent>
        </Card>



      </div>
    </AppShell>
  );
}
