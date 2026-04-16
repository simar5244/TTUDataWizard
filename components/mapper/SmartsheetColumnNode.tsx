"use client";

import { useEffect, useState, type MouseEvent } from "react";
import { Handle, Position, useReactFlow, type NodeProps } from "@xyflow/react";
import { Trash2 } from "lucide-react";
import { useNodeActions } from "@/components/mapper/NodeActionsContext";
import { type RowRange, rangeLabel } from "@/components/mapper/DetailView";

const TYPE_COLORS: Record<string, string> = {
  TEXT_NUMBER: "bg-emerald-50 border-emerald-200 text-emerald-700",
  DATE: "bg-amber-50 border-amber-200 text-amber-700",
  CHECKBOX: "bg-rose-50 border-rose-200 text-rose-700",
  PICKLIST: "bg-indigo-50 border-indigo-200 text-indigo-700",
  CONTACT_LIST: "bg-blue-50 border-blue-200 text-blue-700",
};

export function SmartsheetColumnNode({ id, data, selected }: NodeProps) {
  const { deleteNode } = useNodeActions();
  const { deleteElements } = useReactFlow();
  const { label, colType, colId, colRef, onLabelChange, onTypeChange, detailRanges } = data as {
    label: string;
    colType: string;
    colId: number | string;
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

  const nodeId = String(id || `ss_${String(colId)}`);

  function commitLabel() {
    const trimmed = draftLabel.trim();
    if (!trimmed || !onLabelChange) {
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
        selected ? "border-emerald-500 ring-2 ring-emerald-100" : "border-slate-200"
      } hover:border-emerald-300`}
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
      {/* Default handles — visible only when no ranges are defined */}
      {!hasRanges && (
        <>
          <Handle
            type="target"
            position={Position.Left}
            id={`in-${colId}`}
            className="!h-3 !w-3 !border-2 !border-emerald-500 !bg-white transition-colors hover:!bg-emerald-500"
            style={{ left: '-6px' }}
          />
          <Handle
            type="source"
            position={Position.Left}
            id={`out-${colId}`}
            className="!h-3 !w-3 !border-2 !border-emerald-500 !bg-white transition-colors hover:!bg-emerald-500"
            style={{ left: '-6px' }}
          />
        </>
      )}
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
          <span className={`shrink-0 rounded border px-1.5 py-0.5 text-[9px] ${TYPE_COLORS[colType] || "bg-slate-50 border-slate-200 text-slate-600"}`}>
            {editing ? (
              <select
                value={colType || "TEXT_NUMBER"}
                onChange={(e) => {
                  if (!onTypeChange) return;
                  onTypeChange(nodeId, e.target.value);
                }}
                className="bg-transparent text-[9px] outline-none"
              >
                <option value="TEXT_NUMBER">text number</option>
                <option value="DATE">date</option>
                <option value="CHECKBOX">checkbox</option>
                <option value="PICKLIST">picklist</option>
                <option value="CONTACT_LIST">contact list</option>
              </select>
            ) : (colType?.toLowerCase().replace("_", " ") || "text")}
          </span>
        </div>
        {colRef && (
          <code className="mt-1 block text-[9px] text-emerald-600 font-mono">{colRef}</code>
        )}
        {/* Per-range rows with individual target handles on the left */}
        {hasRanges && (
          <div className="mt-2 flex flex-col gap-1 border-t border-slate-100 pt-1.5">
            {detailRanges!.map((r) => (
              <div key={r.id} className="relative flex items-center justify-between pl-5">
                <Handle
                  type="target"
                  position={Position.Left}
                  id={r.id}
                  className="!absolute !-left-4 !h-2.5 !w-2.5 !border-2 !border-emerald-500 !bg-white transition-colors hover:!bg-emerald-500"
                />
                <span className="font-mono text-[10px] text-emerald-700">{rangeLabel(r)}</span>
                <span className="text-[9px] text-slate-400">{r.end - r.start + 1}r</span>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}
