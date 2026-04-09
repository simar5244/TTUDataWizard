"use client";

import { Handle, Position, type NodeProps } from "@xyflow/react";

export function SmartsheetColumnNode({ data, selected }: NodeProps) {
  const { label, colType, colId } = data as { label: string; colType: string; colId: number };

  return (
    <div
      className={`relative min-w-[160px] rounded-lg border-2 bg-white shadow-sm transition-all ${
        selected ? "border-black" : "border-[#E5E5E5]"
      }`}
    >
      <Handle
        type="target"
        position={Position.Left}
        id={String(colId)}
        className="!h-3 !w-3 !border-2 !border-black !bg-white"
      />
      <div className="px-3 py-2">
        <div className="flex items-center justify-between gap-2">
          <span className="truncate text-xs font-semibold">{label}</span>
          <span className="shrink-0 rounded border border-[#E5E5E5] bg-[#F5F5F5] px-1.5 py-0.5 text-[10px] text-[#6B6B6B]">
            {colType?.toLowerCase() || "text"}
          </span>
        </div>
      </div>
    </div>
  );
}
