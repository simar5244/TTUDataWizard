"use client";

import { Handle, Position, type NodeProps } from "@xyflow/react";
import { Badge } from "@/components/ui/badge";

const TYPE_COLORS: Record<string, string> = {
  number: "bg-indigo-50 border-indigo-200 text-indigo-700",
  string: "bg-emerald-50 border-emerald-200 text-emerald-700",
  date: "bg-amber-50 border-amber-200 text-amber-700",
  boolean: "bg-rose-50 border-rose-200 text-rose-700",
  empty: "bg-slate-50 border-slate-200 text-slate-600",
};

export function ExcelColumnNode({ data, selected }: NodeProps) {
  const { label, dataType, sampleValues, colKey, colRef } = data as {
    label: string;
    dataType: string;
    sampleValues: unknown[];
    colKey: string;
    colRef?: string;
  };

  return (
    <div
      className={`relative min-w-[170px] rounded-xl border-2 bg-white shadow-sm transition-all ${
        selected ? "border-indigo-500 ring-2 ring-indigo-100" : "border-slate-200"
      } hover:border-indigo-300`}
    >
      <div className="px-3 py-2.5">
        <div className="flex items-center justify-between gap-2">
          <span className="truncate text-xs font-semibold text-slate-700">{label}</span>
          <Badge
            variant="outline"
            className={`shrink-0 text-[9px] capitalize ${TYPE_COLORS[dataType] || ""}`}
          >
            {dataType}
          </Badge>
        </div>
        {colRef && (
          <code className="mt-1 block text-[9px] text-indigo-500 font-mono">{colRef}</code>
        )}
      </div>
      <Handle
        type="source"
        position={Position.Right}
        id={colKey}
        className="!h-3 !w-3 !border-2 !border-indigo-500 !bg-white transition-colors hover:!bg-indigo-500"
      />
    </div>
  );
}
