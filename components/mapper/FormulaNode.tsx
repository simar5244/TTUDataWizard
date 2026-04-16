"use client";

import { useState, useCallback, type MouseEvent } from "react";
import { Handle, Position, useReactFlow, type NodeProps } from "@xyflow/react";
import { Trash2, FunctionSquare, ArrowRight } from "lucide-react";
import { useNodeActions } from "@/components/mapper/NodeActionsContext";

// Formula templates with categories
const MATH_TEMPLATES = [
  { label: "Add", formula: "A + B", desc: "Sum two values", requires: ["number", "number"], output: "number" },
  { label: "Subtract", formula: "A - B", desc: "Difference", requires: ["number", "number"], output: "number" },
  { label: "Multiply", formula: "A * B", desc: "Product", requires: ["number", "number"], output: "number" },
  { label: "Divide", formula: "A / B", desc: "Quotient", requires: ["number", "number"], output: "number" },
  { label: "Round", formula: "round(A, 2)", desc: "Round to 2 decimals", requires: ["number"], output: "number" },
  { label: "Power", formula: "A ^ B", desc: "A to power B", requires: ["number", "number"], output: "number" },
  { label: "Modulo", formula: "A % B", desc: "Remainder", requires: ["number", "number"], output: "number" },
  { label: "Abs", formula: "abs(A)", desc: "Absolute value", requires: ["number"], output: "number" },
  { label: "Min", formula: "min(A, B)", desc: "Minimum of two", requires: ["number", "number"], output: "number" },
  { label: "Max", formula: "max(A, B)", desc: "Maximum of two", requires: ["number", "number"], output: "number" },
];

const STRING_TEMPLATES = [
  { label: "Trim", formula: "trim(A)", desc: "Remove leading/trailing spaces", requires: ["string"], output: "string" },
  { label: "Upper", formula: "upper(A)", desc: "UPPERCASE", requires: ["string"], output: "string" },
  { label: "Lower", formula: "lower(A)", desc: "lowercase", requires: ["string"], output: "string" },
  { label: "Concat", formula: "A & B", desc: "Join text", requires: ["string", "string"], output: "string" },
  { label: "Left", formula: "left(A, 5)", desc: "First 5 chars", requires: ["string"], output: "string" },
  { label: "Right", formula: "right(A, 5)", desc: "Last 5 chars", requires: ["string"], output: "string" },
  { label: "Mid", formula: "mid(A, 2, 5)", desc: "Substring from pos 2, len 5", requires: ["string"], output: "string" },
  { label: "Len", formula: "len(A)", desc: "Character count", requires: ["string"], output: "number" },
  { label: "Replace", formula: "replace(A, \"old\", \"new\")", desc: "Replace text", requires: ["string"], output: "string" },
  { label: "Remove", formula: "remove(A, \"x\")", desc: "Remove all 'x'", requires: ["string"], output: "string" },
  { label: "Clean", formula: "clean(A)", desc: "Remove non-printable chars", requires: ["string"], output: "string" },
];

const LOGIC_TEMPLATES = [
  { label: "If", formula: "A > 0 ? A : 0", desc: "Conditional", requires: ["any"], output: "any" },
  { label: "And", formula: "A > 0 && B > 0", desc: "Both true", requires: ["boolean", "boolean"], output: "boolean" },
  { label: "Or", formula: "A > 0 || B > 0", desc: "Either true", requires: ["boolean", "boolean"], output: "boolean" },
  { label: "Equal", formula: "A == B", desc: "Equals", requires: ["any", "any"], output: "boolean" },
  { label: "NotEqual", formula: "A != B", desc: "Not equals", requires: ["any", "any"], output: "boolean" },
  { label: "Greater", formula: "A > B", desc: "A greater than B", requires: ["number", "number"], output: "boolean" },
  { label: "Less", formula: "A < B", desc: "A less than B", requires: ["number", "number"], output: "boolean" },
];

const ALL_TEMPLATES = [...MATH_TEMPLATES, ...STRING_TEMPLATES, ...LOGIC_TEMPLATES];

// Detect input types from connected columns
function detectInputTypes(inputs: { id: string; label: string; dataType?: string }[]): string[] {
  return inputs.map(input => {
    if (input.dataType) return input.dataType;
    const label = input.label.toLowerCase();
    if (label.includes('date')) return 'date';
    if (label.includes('amount') || label.includes('price') || label.includes('cost') || label.includes('revenue') || label.includes('profit') || label.includes('number')) return 'number';
    if (label.includes('id') || label.includes('code') || label.includes('phone')) return 'string';
    return 'any';
  });
}

// Validate formula compatibility
function validateFormula(formula: string, inputs: { id: string; label: string; dataType?: string }[]): { valid: boolean; warning?: string } {
  const inputTypes = detectInputTypes(inputs);
  const template = ALL_TEMPLATES.find(t => formula.toLowerCase().includes(t.label.toLowerCase()) || formula.toLowerCase().includes(t.formula.split('(')[0].toLowerCase()));
  
  if (!template) return { valid: true };
  
  // Check math operations on strings
  if (template.output === 'number' && inputTypes.some(t => t === 'string')) {
    return { valid: true, warning: 'Math on text may fail' };
  }
  
  // Check division by zero risk
  if (template.label === 'Divide' && inputs.length > 1) {
    return { valid: true, warning: 'Ensure divisor is not zero' };
  }
  
  return { valid: true };
}

interface FormulaNodeData {
  label?: string;
  colRef?: string;
  formula: string;
  leftInputs: { id: string; label: string; colRef?: string }[];
  rightInputs: { id: string; label: string; colRef?: string }[];
  onFormulaChange?: (nodeId: string, newFormula: string) => void;
  onDelete?: (nodeId: string) => void;
  onLabelChange?: (nodeId: string, newLabel: string) => void;
}

export function FormulaNode({ id, data, selected }: NodeProps) {
  const { deleteNode } = useNodeActions();
  const { deleteElements } = useReactFlow();
  const nodeData = data as unknown as FormulaNodeData;
  const {
    label = "Formula",
    colRef = "",
    formula = "",
    leftInputs = [],
    rightInputs = [],
    onFormulaChange,
    onLabelChange,
  } = nodeData;

  const [isEditing, setIsEditing] = useState(false);
  const [draftFormula, setDraftFormula] = useState(formula);
  const [draftLabel, setDraftLabel] = useState(label);

  const handleSave = useCallback(() => {
    onFormulaChange?.(id, draftFormula);
    onLabelChange?.(id, draftLabel);
    setIsEditing(false);
  }, [id, draftFormula, draftLabel, onFormulaChange, onLabelChange]);

  const handleDeleteClick = useCallback(async (e: MouseEvent<HTMLButtonElement>) => {
    e.stopPropagation();
    e.preventDefault();
    await deleteElements({ nodes: [{ id: String(id) }] });
    deleteNode(String(id));
  }, [deleteElements, deleteNode, id]);

  const allInputs = [...leftInputs, ...rightInputs];
  const inputCount = allInputs.length;
  const validation = validateFormula(draftFormula || formula, allInputs);

  return (
    <div
      className={`group relative w-[260px] rounded-xl border-2 bg-white shadow-lg transition-all ${
        selected ? "border-indigo-500 ring-2 ring-indigo-100" : "border-slate-200"
      }`}
    >
      {/* LEFT input handles - always 1 spare: connected + 1 */}
      <div className="absolute -left-3 top-0 flex h-full flex-col justify-center gap-3 py-4">
        {Array.from({ length: leftInputs.length + 1 }).map((_, idx) => (
          <Handle
            key={`left-${idx}`}
            type="target"
            position={Position.Left}
            id={`left-${idx}`}
            className="!h-3 !w-3 !border-2 !border-indigo-500 !bg-white hover:!bg-indigo-500"
            style={{ top: `${35 + idx * 35}px` }}
          />
        ))}
      </div>

      {/* RIGHT input handles - always 1 spare: connected + 1 */}
      <div className="absolute -right-3 top-0 flex h-full flex-col justify-center gap-3 py-4">
        {Array.from({ length: rightInputs.length + 1 }).map((_, idx) => (
          <Handle
            key={`right-${idx}`}
            type="target"
            position={Position.Right}
            id={`right-${idx}`}
            className="!h-3 !w-3 !border-2 !border-emerald-500 !bg-white hover:!bg-emerald-500"
            style={{ top: `${35 + idx * 35}px` }}
          />
        ))}
      </div>

      {/* Header */}
      <div className="flex items-center justify-between gap-2 border-b border-slate-100 bg-slate-50/50 px-3 py-2 rounded-t-xl">
          <div className="flex items-center gap-1.5">
            <FunctionSquare className="h-3.5 w-3.5 text-indigo-500" />
            <span className="truncate text-xs font-semibold text-slate-700">{label}</span>
            {colRef && (
              <code className="rounded bg-indigo-100 px-1 py-0.5 text-[9px] font-mono font-semibold text-indigo-700">{colRef}</code>
            )}
          </div>
        <button
          onClick={(e) => { void handleDeleteClick(e); }}
          onMouseDown={(e) => e.stopPropagation()}
          onPointerDown={(e) => e.stopPropagation()}
          className="nodrag nopan rounded p-0.5 opacity-0 hover:bg-red-50 group-hover:opacity-100"
          title="Delete"
        >
          <Trash2 className="h-3 w-3 text-red-400" />
        </button>
      </div>

      {/* Body */}
      <div className="px-3 py-3">
        {/* Formula input area */}
        {isEditing ? (
          <div className="space-y-2">
            <input
              className="w-full rounded border border-slate-200 px-2 py-1 text-xs font-medium outline-none focus:border-indigo-500"
              value={draftLabel}
              onChange={(e) => setDraftLabel(e.target.value)}
              placeholder="Name..."
            />
            <textarea
              className="w-full resize-none rounded border border-slate-200 bg-slate-50 px-2 py-1.5 text-xs font-mono outline-none focus:border-indigo-500"
              rows={2}
              value={draftFormula}
              onChange={(e) => setDraftFormula(e.target.value)}
              placeholder="e.g., LA + LB or trim(LA)…"
            />
            <div className="flex gap-1">
              <button
                onClick={handleSave}
                className="flex-1 rounded bg-indigo-500 px-2 py-1 text-[10px] font-medium text-white hover:bg-indigo-600"
              >
                Apply
              </button>
              <button
                onClick={() => {
                  setDraftFormula(formula);
                  setDraftLabel(label);
                  setIsEditing(false);
                }}
                className="rounded bg-slate-100 px-2 py-1 text-[10px] text-slate-600 hover:bg-slate-200"
              >
                Cancel
              </button>
            </div>
          </div>
        ) : (
          <div
            onClick={() => setIsEditing(true)}
            className={`cursor-pointer rounded-lg border px-3 py-2 transition-colors ${
              formula.trim()
                ? validation.warning 
                  ? "border-amber-200 bg-amber-50/30"
                  : "border-indigo-200 bg-indigo-50/30"
                : "border-slate-200 bg-slate-50"
            }`}
          >
            <code className="block break-all text-[12px] font-mono text-slate-700">
              {formula.trim() ? `= ${formula}` : "Click to add formula..."}
            </code>
            {validation.warning && (
              <div className="mt-1 flex items-center gap-1 text-[9px] text-amber-600">
                <span className="h-1 w-1 rounded-full bg-amber-500" />
                {validation.warning}
              </div>
            )}
          </div>
        )}


        {/* Connected inputs summary - left inputs in indigo, right inputs in emerald */}
        {inputCount > 0 && (
          <div className="mt-3 flex flex-wrap gap-1 border-t border-slate-100 pt-2">
            {leftInputs.map((input, idx) => (
              <span
                key={`left-${input.id}-${idx}`}
                className="inline-flex items-center gap-1 rounded bg-indigo-50 px-1.5 py-0.5 text-[9px] text-indigo-700"
                title={input.label}
              >
                <span className="h-1.5 w-1.5 rounded-full bg-indigo-400" />
                {input.colRef ? (
                  <><code className="font-mono font-bold">{input.colRef}</code><span className="text-indigo-400">:{input.label}</span></>
                ) : input.label}
              </span>
            ))}
            {rightInputs.map((input, idx) => (
              <span
                key={`right-${input.id}-${idx}`}
                className="inline-flex items-center gap-1 rounded bg-emerald-50 px-1.5 py-0.5 text-[9px] text-emerald-700"
                title={input.label}
              >
                <span className="h-1.5 w-1.5 rounded-full bg-emerald-400" />
                {input.colRef ? (
                  <><code className="font-mono font-bold">{input.colRef}</code><span className="text-emerald-400">:{input.label}</span></>
                ) : input.label}
              </span>
            ))}
          </div>
        )}
      </div>

      {/* Output handle - bottom (to ssCol) */}
      <Handle
        type="source"
        position={Position.Bottom}
        id="output"
        className="!h-4 !w-4 !-bottom-2 !border-2 !border-emerald-500 !bg-white hover:!bg-emerald-500"
      />
      {/* Output handle - right (to chain into another formula) */}
      <Handle
        type="source"
        position={Position.Right}
        id="output-right"
        className="!h-3 !w-3 !border-2 !border-emerald-500 !bg-white hover:!bg-emerald-500"
        style={{ right: '-6px', top: '50%' }}
      />

      {/* Output indicator */}
      <div className="absolute -bottom-6 left-1/2 -translate-x-1/2 flex items-center gap-1 text-[10px] text-slate-400">
        <ArrowRight className="h-3 w-3" />
        <span>output</span>
      </div>
    </div>
  );
}
