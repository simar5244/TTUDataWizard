import { formatDynamicColumnName } from "@/lib/dynamic-column";

export interface DuplicateColumnNodeConfig {
  duplicateOnRun?: boolean;
  duplicateNameTemplate?: string;
  label?: string;
}

export interface ResolveDuplicateColumnNameOptions {
  template: string;
  existingTitles?: string[];
  now?: Date;
}

export interface DuplicateColumnNodeLike {
  type?: string;
  data?: {
    label?: string;
    colType?: string;
    columnPosition?: "start" | "end" | "custom";
    customColumnNumber?: number;
    duplicateOnRun?: boolean;
    duplicateNameTemplate?: string;
  };
}

export interface ExistingSheetColumnLike {
  id: number;
  title: string;
}

export interface DuplicateColumnResolvedTarget {
  resolvedName: string;
  resolvedId?: number;
}

export interface DuplicateColumnPreparationResult {
  byBaseLabel: Map<string, DuplicateColumnResolvedTarget>;
  resolvedLabels: Set<string>;
  refreshedTitleToId: Map<string, number>;
}

export interface DuplicateBaseLabelSets {
  raw: Set<string>;
  normalized: Set<string>;
}

export interface SmartsheetTargetNodeLite {
  id: string;
  type?: string;
  position?: { x: number; y: number };
  data?: Record<string, unknown>;
}

export interface SmartsheetTargetEdgeLite {
  source: string;
  target: string;
}

export interface SmartsheetRemoteColumnLite {
  id: number;
  title: string;
  type: string;
}

export interface SmartsheetTargetMergeResult {
  nodes: SmartsheetTargetNodeLite[];
  removedNodeIds: string[];
  remappedNodeIds: Map<string, string>;
}

function renderTemplateForDate(template: string, now: Date): string {
  const yyyy = String(now.getFullYear());
  const mm = String(now.getMonth() + 1).padStart(2, "0");
  const dd = String(now.getDate()).padStart(2, "0");
  return template
    .replace(/\{\{\s*YYYY-MM-DD\s*\}\}/gi, `${yyyy}-${mm}-${dd}`)
    .replace(/\{\{\s*YYYYMMDD\s*\}\}/gi, `${yyyy}${mm}${dd}`)
    .replace(/\{\{\s*MM\/DD\/YYYY\s*\}\}/gi, `${mm}/${dd}/${yyyy}`)
    .replace(/\{\{\s*DATE\s*\}\}/gi, `${mm}/${dd}/${yyyy}`)
    .trim();
}

function ensureUniqueName(candidate: string, existingTitles: string[]): string {
  if (!existingTitles.includes(candidate)) return candidate;
  let suffix = 2;
  let unique = `${candidate} (${suffix})`;
  while (existingTitles.includes(unique)) {
    suffix += 1;
    unique = `${candidate} (${suffix})`;
  }
  return unique;
}

function delay(ms: number): Promise<void> {
  return new Promise((resolve) => {
    setTimeout(resolve, ms);
  });
}

function isTemporaryPersistentTarget(node: SmartsheetTargetNodeLite): boolean {
  if (node.type !== "ssCol") return false;
  const data = (node.data ?? {}) as Record<string, unknown>;
  const rawColId = String(data.colId ?? "").trim();
  const synthetic = data.synthetic === true || rawColId.startsWith("ss_manual_");
  const duplicateOnRun = data.duplicateOnRun === true;
  return synthetic || duplicateOnRun;
}

function normalizeLabel(value: unknown): string {
  return String(value ?? "").trim().toLowerCase();
}

export function collectDuplicateBaseLabels(nodes: Array<{ type?: string; data?: Record<string, unknown> }>): DuplicateBaseLabelSets {
  const raw = new Set<string>();
  const normalized = new Set<string>();

  (Array.isArray(nodes) ? nodes : []).forEach((node) => {
    if (node?.type !== "ssCol") return;
    const duplicateOnRun = node?.data?.duplicateOnRun === true;
    if (!duplicateOnRun) return;
    const baseLabel = String(node?.data?.label ?? "").trim();
    if (!baseLabel) return;
    raw.add(baseLabel);
    normalized.add(normalizeLabel(baseLabel));
  });

  return { raw, normalized };
}

export function isDuplicateColumnOnRunEnabled(nodeData: unknown): boolean {
  if (!nodeData || typeof nodeData !== "object") return false;
  const data = nodeData as DuplicateColumnNodeConfig;
  return data.duplicateOnRun === true;
}

export function getDuplicateColumnTemplate(nodeData: unknown): string {
  if (!nodeData || typeof nodeData !== "object") return "";
  const data = nodeData as DuplicateColumnNodeConfig;
  const fromTemplate = String(data.duplicateNameTemplate ?? "").trim();
  if (fromTemplate) return fromTemplate;
  return String(data.label ?? "").trim();
}

export function resolveDuplicateColumnName(options: ResolveDuplicateColumnNameOptions): string {
  const template = String(options.template ?? "").trim();
  if (!template) return "";
  const existingTitles = (options.existingTitles ?? []).map((title) => String(title));

  if (!options.now) {
    return formatDynamicColumnName(template, {
      existingTitles,
      ensureUnique: true,
    });
  }

  const rendered = renderTemplateForDate(template, options.now) || template;
  return ensureUniqueName(rendered, existingTitles);
}

export async function prepareDuplicateColumnsForRun(options: {
  nodes: DuplicateColumnNodeLike[];
  existingColumns: ExistingSheetColumnLike[];
  createColumn: (args: { title: string; type: string; index: number }) => Promise<{ id: number; title: string }>;
  fetchColumns: () => Promise<ExistingSheetColumnLike[]>;
  now?: Date;
  fetchRetryCount?: number;
  fetchRetryDelayMs?: number;
}): Promise<DuplicateColumnPreparationResult> {
  const nodes = Array.isArray(options.nodes) ? options.nodes : [];
  const existingTitles = (options.existingColumns ?? []).map((column) => String(column.title));
  const byBaseLabel = new Map<string, DuplicateColumnResolvedTarget>();
  const resolvedLabels = new Set<string>();
  const now = options.now ?? new Date();

  for (let index = 0; index < nodes.length; index += 1) {
    const node = nodes[index];
    if (node?.type !== "ssCol") continue;
    const nodeData = node.data ?? {};
    if (!isDuplicateColumnOnRunEnabled(nodeData)) continue;

    const baseLabel = String(nodeData.label ?? "").trim();
    if (!baseLabel) continue;

    const duplicateTemplate = getDuplicateColumnTemplate(nodeData);
    const resolvedName = resolveDuplicateColumnName({
      template: duplicateTemplate || baseLabel,
      existingTitles,
      now,
    });
    if (!resolvedName) continue;

    const duplicatePosition = typeof nodeData.columnPosition === "string" ? nodeData.columnPosition : "end";
    const duplicateCustomNumber =
      typeof nodeData.customColumnNumber === "number" && nodeData.customColumnNumber > 0
        ? Math.floor(nodeData.customColumnNumber)
        : 1;
    const duplicateIndex =
      duplicatePosition === "start"
        ? 0
        : duplicatePosition === "custom"
          ? Math.max(0, Math.min(existingTitles.length, duplicateCustomNumber - 1))
          : existingTitles.length;

    const created = await options.createColumn({
      title: resolvedName,
      type: typeof nodeData.colType === "string" && nodeData.colType ? String(nodeData.colType) : "TEXT_NUMBER",
      index: duplicateIndex,
    });

    const createdId = Number(created.id);
    const safeCreatedId = Number.isFinite(createdId) ? createdId : undefined;
    byBaseLabel.set(baseLabel, {
      resolvedName: String(created.title ?? resolvedName),
      resolvedId: safeCreatedId,
    });
    resolvedLabels.add(String(created.title ?? resolvedName));
    existingTitles.push(String(created.title ?? resolvedName));
  }

  const refreshedTitleToId = new Map<string, number>();
  const retryCount =
    typeof options.fetchRetryCount === "number" && options.fetchRetryCount > 0
      ? Math.floor(options.fetchRetryCount)
      : 3;
  const retryDelayMs =
    typeof options.fetchRetryDelayMs === "number" && options.fetchRetryDelayMs > 0
      ? Math.floor(options.fetchRetryDelayMs)
      : 250;

  for (let attempt = 0; attempt < retryCount; attempt += 1) {
    const refreshedColumns = await options.fetchColumns();
    refreshedTitleToId.clear();
    refreshedColumns.forEach((column) => {
      const id = Number(column.id);
      if (!Number.isFinite(id)) return;
      refreshedTitleToId.set(String(column.title), id);
    });

    const unresolved = Array.from(byBaseLabel.entries()).filter(([, target]) => {
      if (typeof target.resolvedId === "number" && Number.isFinite(target.resolvedId)) return false;
      return !refreshedTitleToId.has(target.resolvedName);
    });
    if (unresolved.length === 0) break;
    if (attempt < retryCount - 1) {
      await delay(retryDelayMs);
    }
  }

  Array.from(byBaseLabel.entries()).forEach(([baseLabel, target]) => {
    const refreshedId = refreshedTitleToId.get(target.resolvedName);
    if (typeof refreshedId === "number" && Number.isFinite(refreshedId)) {
      byBaseLabel.set(baseLabel, {
        resolvedName: target.resolvedName,
        resolvedId: refreshedId,
      });
    }
  });

  return {
    byBaseLabel,
    resolvedLabels,
    refreshedTitleToId,
  };
}

export function mergeSmartsheetTargetNodes(options: {
  existingNodes: SmartsheetTargetNodeLite[];
  remoteColumns: SmartsheetRemoteColumnLite[];
  colRefForIndex: (index: number) => string;
  onLabelChange?: (nodeId: string, label: string) => void;
  onTypeChange?: (nodeId: string, type: string) => void;
}): SmartsheetTargetMergeResult {
  const existingNodes = Array.isArray(options.existingNodes) ? options.existingNodes : [];
  const remoteColumns = Array.isArray(options.remoteColumns) ? options.remoteColumns : [];
  const nonTargets = existingNodes.filter((node) => node.type !== "ssCol");
  const existingTargets = existingNodes.filter((node) => node.type === "ssCol");
  const remoteById = new Map<string, SmartsheetRemoteColumnLite>();
  remoteColumns.forEach((column) => {
    remoteById.set(String(column.id), column);
  });

  const keptTargets: SmartsheetTargetNodeLite[] = [];
  const removedNodeIds: string[] = [];
  const remappedNodeIds = new Map<string, string>();
  const seenRemoteIds = new Set<string>();

  for (let index = 0; index < existingTargets.length; index += 1) {
    const node = existingTargets[index];
    const data = (node.data ?? {}) as Record<string, unknown>;
    const rawColId = String(data.colId ?? "").trim();
    const remote = remoteById.get(rawColId);

    if (remote) {
      seenRemoteIds.add(String(remote.id));
      const currentNodeId = String(node.id || `ss_${String(remote.id)}`);
      const nextNodeId = `ss_${String(remote.id)}`;
      if (currentNodeId !== nextNodeId) {
        remappedNodeIds.set(currentNodeId, nextNodeId);
      }
      keptTargets.push({
        ...node,
        id: nextNodeId,
        data: {
          ...data,
          label: remote.title,
          colType: remote.type,
          colId: remote.id,
          colRef: `R${options.colRefForIndex(index)}`,
          nodeId: nextNodeId,
          onLabelChange: options.onLabelChange,
          onTypeChange: options.onTypeChange,
          _previousNodeId: currentNodeId,
        },
      });
      continue;
    }

    if (isTemporaryPersistentTarget(node)) {
      keptTargets.push(node);
      continue;
    }

    removedNodeIds.push(String(node.id));
  }

  const newRemoteTargets: SmartsheetTargetNodeLite[] = [];
  remoteColumns.forEach((column, index) => {
    const remoteId = String(column.id);
    if (seenRemoteIds.has(remoteId)) return;
    newRemoteTargets.push({
      id: `ss_${remoteId}`,
      type: "ssCol",
      position: { x: 800, y: 80 + index * 65 },
      data: {
        label: column.title,
        colType: column.type,
        colId: column.id,
        colRef: `R${options.colRefForIndex(index)}`,
        index,
        nodeId: `ss_${remoteId}`,
        onLabelChange: options.onLabelChange,
        onTypeChange: options.onTypeChange,
      },
    });
  });

  return {
    nodes: [...nonTargets, ...keptTargets, ...newRemoteTargets],
    removedNodeIds,
    remappedNodeIds,
  };
}
