"use client";

import { useState, useEffect } from "react";
import {
  BaseEdge,
  EdgeLabelRenderer,
  getStraightPath,
  useReactFlow,
  type EdgeProps,
} from "@xyflow/react";
import { ArrowLeftRight, Trash2 } from "lucide-react";

const OPS: { label: string; fn: (a: string, b: string) => string }[] = [
  { label: "Direct",    fn: (a)    => a },
  { label: "ADD",       fn: (a, b) => `${a} + ${b}` },
  { label: "SUBTRACT",  fn: (a, b) => `${a} - ${b}` },
  { label: "MULTIPLY",  fn: (a, b) => `${a} * ${b}` },
  { label: "DIVIDE",    fn: (a, b) => `${a} / ${b}` },
  { label: "CONCAT",    fn: (a, b) => `CONCAT(${a}, ${b})` },
  { label: "ROUND",     fn: (a)    => `ROUND(${a}, 2)` },
];

export function FormulaEdge({
  id, sourceX, sourceY, targetX, targetY, data,
}: EdgeProps) {
  const { setEdges } = useReactFlow();

  const formula    = (data?.formula     as string) || "";
  const srcLabel   = (data?.sourceLabel as string) || "Source";
  const tgtLabel   = (data?.targetLabel as string) || "Target";
  const isNew      = data?.isNew === true;

  const [editing,   setEditing]   = useState(isNew);
  const [draft,     setDraft]     = useState(formula);
  const [swapped,   setSwapped]   = useState(false);
  const [activeOp,  setActiveOp]  = useState<string | null>(null);

  const col1 = swapped ? tgtLabel : srcLabel;
  const col2 = swapped ? srcLabel : tgtLabel;

  useEffect(() => {
    if (isNew) {
      setEdges((eds) =>
        eds.map((e) => e.id === id ? { ...e, data: { ...e.data, isNew: false } } : e)
      );
    }
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  function applyOp(label: string) {
    setActiveOp(label);
    const op = OPS.find((o) => o.label === label);
    if (op) setDraft(op.fn(col1, col2));
  }

  function handleSwap() {
    const ns = !swapped;
    setSwapped(ns);
    if (activeOp) {
      const op = OPS.find((o) => o.label === activeOp);
      if (op) {
        const c1 = ns ? tgtLabel : srcLabel;
        const c2 = ns ? srcLabel : tgtLabel;
        setDraft(op.fn(c1, c2));
      }
    }
  }

  function saveFormula() {
    setEdges((eds) =>
      eds.map((e) => e.id === id ? { ...e, data: { ...e.data, formula: draft } } : e)
    );
    setEditing(false);
  }

  function deleteEdge(ev: React.MouseEvent) {
    ev.stopPropagation();
    setEdges((eds) => eds.filter((e) => e.id !== id));
  }

  const [edgePath, labelX, labelY] = getStraightPath({ sourceX, sourceY, targetX, targetY });

  return (
    <>
      <BaseEdge
        id={id}
        path={edgePath}
        style={{
          stroke: formula ? "#0A0A0A" : "#A1A1A1",
          strokeWidth: formula ? 2 : 1.5,
          strokeDasharray: formula ? "0" : "4 4",
        }}
      />
      <EdgeLabelRenderer>
        <div
          style={{ transform: `translate(-50%, -50%) translate(${labelX}px,${labelY}px)` }}
          className="nodrag nopan absolute"
        >
          {editing ? (
            <div
              className="flex flex-col gap-2 rounded-xl border border-[#E0E0E0] bg-white p-3 shadow-xl"
              style={{ minWidth: 288, zIndex: 1000 }}
            >
              {/* Header: source → target */}
              <div className="flex items-center gap-2 text-[11px]">
                <span className="rounded bg-[#F0F0F0] px-1.5 py-0.5 font-medium text-[#0A0A0A] max-w-[90px] truncate">
                  {srcLabel}
                </span>
                <span className="text-[#A1A1A1]">→</span>
                <span className="rounded bg-[#EEF2FF] px-1.5 py-0.5 font-medium text-[#3730A3] max-w-[90px] truncate">
                  {tgtLabel}
                </span>
                <button
                  className="ml-auto rounded p-0.5 hover:bg-[#F5F5F5]"
                  onClick={handleSwap}
                  title="Swap order"
                >
                  <ArrowLeftRight className="h-3 w-3 text-[#6B6B6B]" />
                </button>
              </div>

              {/* Operation buttons */}
              <div className="flex flex-wrap gap-1">
                {OPS.map((op) => (
                  <button
                    key={op.label}
                    onClick={() => applyOp(op.label)}
                    className={`rounded px-2 py-0.5 text-[10px] font-medium transition-colors ${
                      activeOp === op.label
                        ? "bg-black text-white"
                        : "bg-[#F5F5F5] text-[#555] hover:bg-[#EAEAEA]"
                    }`}
                  >
                    {op.label}
                  </button>
                ))}
              </div>

              {/* Formula input */}
              <div className="flex items-center gap-1 rounded border border-[#E5E5E5] bg-[#FAFAFA] px-2 py-1.5">
                <span className="text-[10px] text-[#A1A1A1] select-none">=</span>
                <input
                  autoFocus
                  className="flex-1 bg-transparent text-xs font-mono outline-none"
                  placeholder={`${srcLabel} + ${tgtLabel}`}
                  value={draft}
                  onChange={(e) => { setDraft(e.target.value); setActiveOp(null); }}
                  onKeyDown={(e) => {
                    if (e.key === "Enter") saveFormula();
                    if (e.key === "Escape") setEditing(false);
                  }}
                />
              </div>

              {/* Actions */}
              <div className="flex justify-end gap-1.5">
                <button
                  onClick={() => setEditing(false)}
                  className="rounded px-2.5 py-1 text-[10px] text-[#6B6B6B] hover:bg-[#F5F5F5]"
                >
                  Cancel
                </button>
                <button
                  onClick={saveFormula}
                  className="rounded bg-black px-2.5 py-1 text-[10px] text-white hover:bg-[#333]"
                >
                  Apply
                </button>
              </div>
            </div>
          ) : (
            <div className="group flex items-center gap-1 rounded-md border border-[#E5E5E5] bg-white px-2 py-1 text-[11px] shadow-sm hover:border-[#A1A1A1]">
              <button
                className="flex items-center gap-1 hover:opacity-70"
                onClick={() => { setDraft(formula); setEditing(true); }}
                title="Click to edit formula"
              >
                {formula ? (
                  <span className="font-mono font-medium text-[#0A0A0A] max-w-[120px] truncate">{formula}</span>
                ) : (
                  <span className="text-[#A1A1A1]">direct →</span>
                )}
              </button>
              <button
                onClick={deleteEdge}
                title="Remove connection"
                className="ml-1 rounded p-0.5 opacity-0 hover:bg-red-50 group-hover:opacity-100"
              >
                <Trash2 className="h-2.5 w-2.5 text-red-400" />
              </button>
            </div>
          )}
        </div>
      </EdgeLabelRenderer>
    </>
  );
}
