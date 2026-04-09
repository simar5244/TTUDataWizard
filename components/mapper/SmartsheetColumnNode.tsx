"use client";

import { Handle, Position, type NodeProps } from "@xyflow/react";

const TYPE_COLORS: Record<string, string> = {
  TEXT_NUMBER: "bg-emerald-50 border-emerald-200 text-emerald-700",
  DATE: "bg-amber-50 border-amber-200 text-amber-700",
  CHECKBOX: "bg-rose-50 border-rose-200 text-rose-700",
  PICKLIST: "bg-indigo-50 border-indigo-200 text-indigo-700",
  CONTACT_LIST: "bg-blue-50 border-blue-200 text-blue-700",
};

export function SmartsheetColumnNode({ data, selected }: NodeProps) {
  const { label, colType, colId, colRef } = data as { label: string; colType: string; colId: number; colRef?: string };

  return (
    <div
      className={`relative min-w-[170px] rounded-xl border-2 bg-white shadow-sm transition-all ${
        selected ? "border-emerald-500 ring-2 ring-emerald-100" : "border-slate-200"
      } hover:border-emerald-300`}
    >
      {/* Target handle - for receiving from formula output */}
      <Handle
        type="target"
        position={Position.Left}
        id={`in-${colId}`}
        className="!h-3 !w-3 !border-2 !border-emerald-500 !bg-white transition-colors hover:!bg-emerald-500"
        style={{ left: '-6px' }}
      />
      {/* Source handle - for sending to formula input */}
      <Handle
        type="source"
        position={Position.Left}
        id={`out-${colId}`}
        className="!h-3 !w-3 !border-2 !border-emerald-500 !bg-white transition-colors hover:!bg-emerald-500"
        style={{ left: '-6px' }}
      />
      <div className="px-3 py-2.5">
        <div className="flex items-center justify-between gap-2">
          <span className="truncate text-xs font-semibold text-slate-700">{label}</span>
          <span className={`shrink-0 rounded border px-1.5 py-0.5 text-[9px] ${TYPE_COLORS[colType] || "bg-slate-50 border-slate-200 text-slate-600"}`}>
            {colType?.toLowerCase().replace("_", " ") || "text"}
          </span>
        </div>
        {colRef && (
          <code className="mt-1 block text-[9px] text-emerald-600 font-mono">{colRef}</code>
        )}
      </div>
    </div>
  );
}
