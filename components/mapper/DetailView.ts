/**
 * DetailView.ts
 *
 * Isolated module for "Detailed Mapping" feature.
 * Manages row-range definitions per node (both source/input and output).
 * Zero dependencies on MapperWorkspace internals — only plain data types.
 */

// ─── Types ────────────────────────────────────────────────────────────────────

/** A single contiguous row range, e.g. rows 1-99 */
export interface RowRange {
  id: string;        // unique within the node's list
  start: number;
  end: number;
}

/** All row-range definitions for a single node */
export interface NodeDetailMapping {
  nodeId: string;
  ranges: RowRange[];
}

/**
 * Top-level store: keyed by nodeId.
 * Serialisable to JSON (stored in mapping meta).
 */
export type DetailMappingStore = Record<string, NodeDetailMapping>;

// ─── Helpers ──────────────────────────────────────────────────────────────────

/** Generate a stable unique id for a new range */
export function newRangeId(): string {
  return `r_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`;
}

/**
 * Parse a user-typed range string like "1-99" or "107-112" or "5".
 * Returns null if the input is not a valid range.
 */
export function parseRangeInput(raw: string): { start: number; end: number } | null {
  const trimmed = raw.trim();
  const single = /^\d+$/.exec(trimmed);
  if (single) {
    const n = parseInt(trimmed, 10);
    if (!Number.isFinite(n) || n < 1) return null;
    return { start: n, end: n };
  }
  const rangeMatch = /^(\d+)\s*[-–]\s*(\d+)$/.exec(trimmed);
  if (rangeMatch) {
    const a = parseInt(rangeMatch[1], 10);
    const b = parseInt(rangeMatch[2], 10);
    if (!Number.isFinite(a) || !Number.isFinite(b) || a < 1 || b < a) return null;
    return { start: a, end: b };
  }
  return null;
}

/** Total row count for a list of ranges */
export function totalRows(ranges: RowRange[]): number {
  return ranges.reduce((sum, r) => sum + (r.end - r.start + 1), 0);
}

/** Human-readable label for a range, e.g. "1–99" */
export function rangeLabel(r: RowRange): string {
  return r.start === r.end ? `${r.start}` : `${r.start}–${r.end}`;
}

/**
 * Add a new range to a node's detail mapping.
 * Returns an updated copy of the store.
 */
export function addRange(
  store: DetailMappingStore,
  nodeId: string,
  start: number,
  end: number
): DetailMappingStore {
  const existing = store[nodeId] ?? { nodeId, ranges: [] };
  const updated: NodeDetailMapping = {
    ...existing,
    ranges: [
      ...existing.ranges,
      { id: newRangeId(), start, end },
    ],
  };
  return { ...store, [nodeId]: updated };
}

/**
 * Remove a range from a node's detail mapping.
 * Returns an updated copy of the store.
 */
export function removeRange(
  store: DetailMappingStore,
  nodeId: string,
  rangeId: string
): DetailMappingStore {
  const existing = store[nodeId];
  if (!existing) return store;
  const ranges = existing.ranges.filter((r) => r.id !== rangeId);
  return { ...store, [nodeId]: { ...existing, ranges } };
}

/**
 * Validate that a source→target connection has matching total row counts.
 * Returns null if valid, or an error string.
 */
export function validateRowCountMatch(
  store: DetailMappingStore,
  sourceNodeId: string,
  targetNodeId: string
): string | null {
  const src = store[sourceNodeId];
  const tgt = store[targetNodeId];
  if (!src && !tgt) return null;        // neither side has ranges — no constraint
  if (!src || src.ranges.length === 0) return null;
  if (!tgt || tgt.ranges.length === 0) return null;
  const srcCount = totalRows(src.ranges);
  const tgtCount = totalRows(tgt.ranges);
  if (srcCount !== tgtCount) {
    return `Row count mismatch: source has ${srcCount} row(s), target has ${tgtCount} row(s).`;
  }
  return null;
}

/**
 * Validate all edges in a detailed-mapping context.
 *
 * With per-range handles, each edge carries a specific sourceHandle / targetHandle
 * (both are range ids like "r_xxx"). Validation checks that the two individual
 * ranges have the same row count. Edges without range handles are skipped — they
 * are not constrained by detailed mapping.
 *
 * Falls back to the old node-level total comparison only for edges where BOTH
 * sides lack range handles but both nodes have ranges defined (legacy behaviour).
 */
export function validateAllEdges(
  store: DetailMappingStore,
  edges: { source: string; target: string; sourceHandle?: string | null; targetHandle?: string | null }[]
): string[] {
  const isRangeId = (h: string | null | undefined) => typeof h === "string" && h.startsWith("r_");

  function findRangeInStore(rangeId: string): RowRange | null {
    for (const entry of Object.values(store)) {
      const r = entry.ranges.find((rr) => rr.id === rangeId);
      if (r) return r;
    }
    return null;
  }

  const errors: string[] = [];
  for (const edge of edges) {
    const srcIsRange = isRangeId(edge.sourceHandle);
    const tgtIsRange = isRangeId(edge.targetHandle);

    if (srcIsRange && tgtIsRange) {
      // Both sides have explicit range handles — compare just those two ranges
      const srcRange = findRangeInStore(edge.sourceHandle!);
      const tgtRange = findRangeInStore(edge.targetHandle!);
      if (!srcRange || !tgtRange) continue; // unknown range id, skip
      const srcCount = srcRange.end - srcRange.start + 1;
      const tgtCount = tgtRange.end - tgtRange.start + 1;
      if (srcCount !== tgtCount) {
        errors.push(
          `${edge.source} → ${edge.target}: Row count mismatch: source range has ${srcCount} row(s), target range has ${tgtCount} row(s).`
        );
      }
    } else if (srcIsRange || tgtIsRange) {
      // Only one side is a range handle — the other side will expand/contract automatically, no error
      continue;
    }
    // Neither side is a range handle — no detailed-mapping constraint applies
  }
  return errors;
}
