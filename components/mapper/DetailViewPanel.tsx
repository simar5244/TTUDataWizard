"use client";

/**
 * DetailViewPanel.tsx
 *
 * Self-contained panel that renders when a user double-clicks a node
 * while "Detailed Mapping" mode is active.
 *
 * Props are intentionally simple — it only needs the store slice for
 * the open node plus callback functions. Zero coupling to MapperWorkspace internals.
 */

import { useState, useRef, useEffect } from "react";
import { X, Plus, Circle } from "lucide-react";
import {
  type NodeDetailMapping,
  type DetailMappingStore,
  parseRangeInput,
  addRange,
  removeRange,
  rangeLabel,
  totalRows,
} from "@/components/mapper/DetailView";

// ─── Sub-components ───────────────────────────────────────────────────────────

interface RangeInputRowProps {
  onAdd: (start: number, end: number) => void;
}

function RangeInputRow({ onAdd }: RangeInputRowProps) {
  const [value, setValue] = useState("");
  const [error, setError] = useState<string | null>(null);
  const inputRef = useRef<HTMLInputElement>(null);

  function commit() {
    const parsed = parseRangeInput(value);
    if (!parsed) {
      setError("Use format: 1-99 or 5");
      return;
    }
    setError(null);
    setValue("");
    onAdd(parsed.start, parsed.end);
    inputRef.current?.focus();
  }

  return (
    <div className="mt-1.5">
      <div className="flex items-center gap-1.5">
        <input
          ref={inputRef}
          value={value}
          onChange={(e) => { setValue(e.target.value); setError(null); }}
          onKeyDown={(e) => {
            if (e.key === "Enter") { e.preventDefault(); commit(); }
          }}
          placeholder="e.g. 1-99 or 107"
          className="h-7 flex-1 rounded border border-slate-200 bg-white px-2 text-[11px] text-slate-700 outline-none placeholder:text-slate-400 focus:border-indigo-400"
        />
        <button
          type="button"
          onClick={commit}
          className="flex h-7 w-7 items-center justify-center rounded border border-slate-200 bg-slate-50 text-slate-600 transition-colors hover:bg-indigo-50 hover:border-indigo-300 hover:text-indigo-600"
          title="Add range"
        >
          <Plus className="h-3.5 w-3.5" />
        </button>
      </div>
      {error && <p className="mt-0.5 text-[10px] text-red-500">{error}</p>}
    </div>
  );
}

interface RangeListProps {
  mapping: NodeDetailMapping | undefined;
  onRemove: (rangeId: string) => void;
  accentColor: string;
}

function RangeList({ mapping, onRemove, accentColor }: RangeListProps) {
  if (!mapping || mapping.ranges.length === 0) {
    return (
      <p className="mt-1.5 text-[10px] text-slate-400 italic">No ranges defined — all rows included.</p>
    );
  }

  return (
    <div className="mt-1.5 flex flex-col gap-1">
      {mapping.ranges.map((r) => (
        <div
          key={r.id}
          className="flex items-center gap-1.5 rounded-md border border-slate-100 bg-slate-50 px-2 py-1"
        >
          <Circle className={`h-2 w-2 shrink-0 fill-current ${accentColor}`} />
          <span className="flex-1 font-mono text-[11px] text-slate-700">{rangeLabel(r)}</span>
          <span className="text-[10px] text-slate-400">{r.end - r.start + 1} row{r.end - r.start + 1 !== 1 ? "s" : ""}</span>
          <button
            type="button"
            onClick={() => onRemove(r.id)}
            className="ml-1 rounded p-0.5 text-slate-400 transition-colors hover:bg-red-50 hover:text-red-500"
            title="Remove range"
          >
            <X className="h-3 w-3" />
          </button>
        </div>
      ))}
      <p className="mt-0.5 text-[10px] text-slate-500 font-medium">
        Total: {totalRows(mapping.ranges)} rows
      </p>
    </div>
  );
}

// ─── Main Panel ───────────────────────────────────────────────────────────────

export interface DetailViewPanelProps {
  nodeId: string;
  nodeLabel: string;
  nodeType: "source" | "target";   // "source" = excel/left, "target" = ss/right
  store: DetailMappingStore;
  onStoreChange: (next: DetailMappingStore) => void;
  onClose: () => void;
  /** screen-space position hint (top-left of the panel) */
  x: number;
  y: number;
}

export function DetailViewPanel({
  nodeId,
  nodeLabel,
  nodeType,
  store,
  onStoreChange,
  onClose,
  x,
  y,
}: DetailViewPanelProps) {
  const panelRef = useRef<HTMLDivElement>(null);

  const isSource = nodeType === "source";
  const accentBorder = isSource ? "border-indigo-400" : "border-emerald-400";
  const accentHeader = isSource
    ? "bg-indigo-50 border-indigo-200 text-indigo-700"
    : "bg-emerald-50 border-emerald-200 text-emerald-700";
  const accentBlip = isSource ? "text-indigo-400" : "text-emerald-400";
  const accentFocus = isSource ? "focus:border-indigo-400" : "focus:border-emerald-400";

  // Close on outside click
  useEffect(() => {
    function handleOutside(e: MouseEvent) {
      if (panelRef.current && !panelRef.current.contains(e.target as Node)) {
        onClose();
      }
    }
    document.addEventListener("mousedown", handleOutside);
    return () => document.removeEventListener("mousedown", handleOutside);
  }, [onClose]);

  // Keep panel on-screen
  const PANEL_W = 260;
  const PANEL_H = 320;
  const clampedX = Math.min(x, window.innerWidth - PANEL_W - 16);
  const clampedY = Math.min(y, window.innerHeight - PANEL_H - 16);

  function handleAdd(start: number, end: number) {
    onStoreChange(addRange(store, nodeId, start, end));
  }

  function handleRemove(rangeId: string) {
    onStoreChange(removeRange(store, nodeId, rangeId));
  }

  const mapping = store[nodeId];

  return (
    <div
      ref={panelRef}
      className={`fixed z-50 w-[260px] rounded-xl border-2 bg-white shadow-xl ${accentBorder}`}
      style={{ left: clampedX, top: clampedY }}
      // Prevent ReactFlow from interpreting clicks/drags on this panel as canvas events
      onMouseDown={(e) => e.stopPropagation()}
      onPointerDown={(e) => e.stopPropagation()}
    >
      {/* Header */}
      <div className={`flex items-center justify-between rounded-t-lg border-b px-3 py-2 ${accentHeader}`}>
        <div className="flex flex-col gap-0.5">
          <span className="text-[11px] font-semibold uppercase tracking-wide opacity-70">
            {isSource ? "Input rows" : "Output rows"}
          </span>
          <span className="text-xs font-semibold truncate max-w-[180px]">{nodeLabel}</span>
        </div>
        <button
          type="button"
          onClick={onClose}
          className="rounded p-0.5 opacity-60 transition-opacity hover:opacity-100"
        >
          <X className="h-3.5 w-3.5" />
        </button>
      </div>

      {/* Body */}
      <div className="px-3 py-2.5">
        <p className="text-[10px] text-slate-500 leading-snug">
          Define which row ranges to include from this column.<br />
          Leave empty to use all rows.
        </p>

        <RangeList mapping={mapping} onRemove={handleRemove} accentColor={accentBlip} />

        <div className={`mt-2 border-t border-slate-100 pt-2`}>
          <p className="text-[10px] font-medium text-slate-500 mb-1">Add range</p>
          <RangeInputRow onAdd={handleAdd} />
        </div>
      </div>

      {/* Footer hint */}
      <div className="rounded-b-xl border-t border-slate-100 bg-slate-50 px-3 py-1.5">
        <p className="text-[10px] text-slate-400">
          Tip: connect matching row-count ranges across columns to map them precisely.
        </p>
      </div>
    </div>
  );
}
