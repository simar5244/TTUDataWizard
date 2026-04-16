"use client";

import { useEffect, useState, type MouseEvent } from "react";
import { Handle, Position, useReactFlow, type NodeProps } from "@xyflow/react";
import { Badge } from "@/components/ui/badge";
import { Trash2 } from "lucide-react";
import { useNodeActions } from "@/components/mapper/NodeActionsContext";
import { type RowRange, rangeLabel } from "@/components/mapper/DetailView";

const TYPE_COLORS: Record<string, string> = {
  number: "bg-indigo-50 border-indigo-200 text-indigo-700",
  string: "bg-emerald-50 border-emerald-200 text-emerald-700",
  date: "bg-amber-50 border-amber-200 text-amber-700",
  boolean: "bg-rose-50 border-rose-200 text-rose-700",
  empty: "bg-slate-50 border-slate-200 text-slate-600",
};

export function ExcelColumnNode({ id, data, selected }: NodeProps) {
  const { deleteNode } = useNodeActions();
  const { deleteElements } = useReactFlow();
  const { label, dataType, colKey, colRef, onLabelChange, onTypeChange, detailRanges } = data as {
    label: string;
    dataType: string;
    colKey: string;
    colRef?: string;
    onLabelChange?: (nodeId: string, label: string) => void;
    onTypeChange?: (nodeId: string, type: string) => void;
    detailRanges?: RowRange[];
  };
  const hasRanges = Array.isArray(detailRanges) && detailRanges.length > 0;
  const [editing, setEditing] = useState(false);
  const [draftLabel, setDraftLabel] = useState(label);

  useEffect(() => {
    setDraftLabel(label);
  }, [label]);

  const nodeId = String(id || colKey || "");

  function commitLabel() {
    const trimmed = draftLabel.trim();
    if (!trimmed || !onLabelChange || !nodeId) {
      setEditing(false);
      return;
    }
    onLabelChange(nodeId, trimmed);
    setEditing(false);
  }

  async function handleDeleteClick(e: MouseEvent<HTMLButtonElement>) {
    e.preventDefault();
    e.stopPropagation();
    if (!nodeId) return;
    await deleteElements({ nodes: [{ id: nodeId }] });
    deleteNode(nodeId);
  }

  return (
    <div
      onDoubleClick={() => { if (!hasRanges) setEditing(true); }}
      className={`group relative min-w-[170px] rounded-xl border-2 bg-white shadow-sm transition-all ${
        selected ? "border-indigo-500 ring-2 ring-indigo-100" : "border-slate-200"
      } hover:border-indigo-300`}
    >
      <button
        onClick={(e) => { void handleDeleteClick(e); }}
        onMouseDown={(e) => e.stopPropagation()}
        onPointerDown={(e) => e.stopPropagation()}
        className="nodrag nopan absolute -right-2 -top-2 z-10 hidden h-4 w-4 items-center justify-center rounded-full bg-red-500 text-white shadow group-hover:flex"
        title="Delete"
      >
        <Trash2 className="h-2.5 w-2.5" />
      </button>
      <div className="px-3 py-2.5">
        <div className="flex items-center justify-between gap-2">
          {editing ? (
            <input
              value={draftLabel}
              onChange={(e) => setDraftLabel(e.target.value)}
              onBlur={commitLabel}
              onKeyDown={(e) => {
                if (e.key === "Enter") commitLabel();
                if (e.key === "Escape") {
                  setDraftLabel(label);
                  setEditing(false);
                }
              }}
              autoFocus
              className="h-6 w-full rounded border border-slate-200 px-1.5 text-xs font-semibold text-slate-700 outline-none"
            />
          ) : (
            <span className="truncate text-xs font-semibold text-slate-700">{label}</span>
          )}
          <Badge
            variant="outline"
            className={`shrink-0 text-[9px] capitalize ${TYPE_COLORS[dataType] || ""}`}
          >
            {editing ? (
              <select
                value={dataType}
                onChange={(e) => {
                  if (!onTypeChange || !nodeId) return;
                  onTypeChange(nodeId, e.target.value);
                }}
                className="bg-transparent text-[9px] outline-none"
              >
                <option value="string">string</option>
                <option value="number">number</option>
                <option value="date">date</option>
                <option value="boolean">boolean</option>
                <option value="empty">empty</option>
              </select>
            ) : dataType}
          </Badge>
        </div>
        {colRef && (
          <code className="mt-1 block text-[9px] text-indigo-500 font-mono">{colRef}</code>
        )}
        {/* Per-range rows with individual handles */}
        {hasRanges && (
          <div className="mt-2 flex flex-col gap-1 border-t border-slate-100 pt-1.5">
            {detailRanges!.map((r) => (
              <div key={r.id} className="relative flex items-center justify-between pr-5">
                <span className="font-mono text-[10px] text-indigo-600">{rangeLabel(r)}</span>
                <span className="text-[9px] text-slate-400">{r.end - r.start + 1}r</span>
                <Handle
                  type="source"
                  position={Position.Right}
                  id={r.id}
                  className="!absolute !-right-4 !h-2.5 !w-2.5 !border-2 !border-indigo-500 !bg-white transition-colors hover:!bg-indigo-500"
                />
              </div>
            ))}
          </div>
        )}
      </div>
      {/* Default handle — hidden when ranges are defined so only range handles are connectable */}
      {!hasRanges && (
        <Handle
          type="source"
          position={Position.Right}
          id={colKey}
          className="!h-3 !w-3 !border-2 !border-indigo-500 !bg-white transition-colors hover:!bg-indigo-500"
        />
      )}
    </div>
  );
}
