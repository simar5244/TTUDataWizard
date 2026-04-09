"use client";

import { Handle, Position, type NodeProps } from "@xyflow/react";
import { Badge } from "@/components/ui/badge";

const TYPE_COLORS: Record<string, string> = {
  number: "bg-blue-50 border-blue-200",
  string: "bg-green-50 border-green-200",
  date: "bg-purple-50 border-purple-200",
  boolean: "bg-orange-50 border-orange-200",
  empty: "bg-gray-50 border-gray-200",
};

export function ExcelColumnNode({ data, selected }: NodeProps) {
  const { label, dataType, sampleValues, colKey } = data as {
    label: string;
    dataType: string;
    sampleValues: unknown[];
    colKey: string;
  };

  return (
    <div
      className={`relative min-w-[160px] rounded-lg border-2 bg-white shadow-sm transition-all ${
        selected ? "border-black" : "border-[#E5E5E5]"
      }`}
    >
      <div className="px-3 py-2">
        <div className="flex items-center justify-between gap-2">
          <span className="truncate text-xs font-semibold">{label}</span>
          <Badge
            variant="outline"
            className={`shrink-0 text-[10px] capitalize ${TYPE_COLORS[dataType] || ""}`}
          >
            {dataType}
          </Badge>
        </div>
        {sampleValues && sampleValues.length > 0 && (
          <p className="mt-1 truncate text-[10px] text-[#A1A1A1]">
            e.g. {String(sampleValues[0])}
            {sampleValues[1] !== undefined ? `, ${String(sampleValues[1])}` : ""}
          </p>
        )}
      </div>
      <Handle
        type="source"
        position={Position.Right}
        id={colKey}
        className="!h-3 !w-3 !border-2 !border-black !bg-white"
      />
    </div>
  );
}
