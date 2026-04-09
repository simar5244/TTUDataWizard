"use client";

import { BaseEdge, EdgeLabelRenderer, getBezierPath, type EdgeProps } from "@xyflow/react";
import { useReactFlow } from "@xyflow/react";
import { Trash2 } from "lucide-react";

export function SimpleEdge({
  id,
  sourceX,
  sourceY,
  targetX,
  targetY,
  sourcePosition,
  targetPosition,
  data,
}: EdgeProps) {
  const { setEdges } = useReactFlow();

  const [edgePath, labelX, labelY] = getBezierPath({
    sourceX,
    sourceY,
    sourcePosition,
    targetX,
    targetY,
    targetPosition,
    curvature: 0.3,
  });

  const isDirectMap = data?.isDirectMap === true;
  const isFormulaOutput = data?.isFormulaOutput === true;

  function deleteEdge(ev: React.MouseEvent) {
    ev.stopPropagation();
    setEdges((eds) => eds.filter((e) => e.id !== id));
  }

  return (
    <>
      <BaseEdge
        id={id}
        path={edgePath}
        style={{
          stroke: isDirectMap ? "#6366f1" : isFormulaOutput ? "#10b981" : "#64748b",
          strokeWidth: isDirectMap || isFormulaOutput ? 2.5 : 1.5,
          strokeDasharray: isFormulaOutput ? "0" : isDirectMap ? "0" : "5 5",
        }}
      />
      <EdgeLabelRenderer>
        <div
          style={{
            transform: `translate(-50%, -50%) translate(${labelX}px,${labelY}px)`,
          }}
          className="nodrag nopan pointer-events-auto absolute"
        >
          <button
            onClick={deleteEdge}
            className="flex h-5 w-5 items-center justify-center rounded-full bg-white opacity-0 shadow-sm transition-opacity hover:bg-red-50 group-hover:opacity-100"
            style={{ opacity: 0.9 }}
            title="Delete connection"
          >
            <Trash2 className="h-3 w-3 text-red-400" />
          </button>
        </div>
      </EdgeLabelRenderer>
    </>
  );
}
