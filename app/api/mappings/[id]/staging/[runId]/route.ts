import { NextRequest, NextResponse } from "next/server";
import { getServerSession } from "next-auth";
import { Prisma } from "@prisma/client";
import { authOptions } from "@/lib/auth";
import { prisma } from "@/lib/prisma";
import { assertEditAllowed, SecurityPolicyError } from "@/lib/security";
import { addSheetColumn, addSheetRows, deleteSheetColumn, getSheet, getSheetRows, updateSheetRows } from "@/lib/smartsheet";
import { formatDynamicColumnName } from "@/lib/dynamic-column";

interface DiffCell {
  row: number;
  rowId?: number | null;
  rowPath?: string | null;
  rowKind?: "detail" | "parent_or_summary" | "new_row";
  column: string;
  columnId?: number | null;
  productionValue: unknown;
  stagingValue: unknown;
  resolution?: "keep_production" | "use_staging";
  action?: "update" | "append" | "skip";
  skipReason?: "protected_formula" | "protected_parent_summary" | "locked" | "excluded_by_rule";
}

interface DynamicTargetColumnConfig {
  enabled?: boolean;
  sourceLabel?: string;
  sourceNodeId?: string;
  nameTemplate?: string;
  columnPosition?: "start" | "end" | "custom";
  customColumnNumber?: number;
}

export async function GET(req: NextRequest, { params }: { params: { id: string; runId: string } }) {
  const session = await getServerSession(authOptions);
  if (!session?.user) return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  const userId = (session.user as { id: string }).id;

  const run = await prisma.stagingRun.findFirst({
    where: { id: params.runId, mappingId: params.id, userId },
    include: { mappingVersion: true },
  });

  if (!run) return NextResponse.json({ error: "Not found" }, { status: 404 });
  return NextResponse.json(run);
}

export async function PUT(req: NextRequest, { params }: { params: { id: string; runId: string } }) {
  const session = await getServerSession(authOptions);
  if (!session?.user) return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  const userId = (session.user as { id: string }).id;

  try {
    await assertEditAllowed(userId);
  } catch (e) {
    if (e instanceof SecurityPolicyError) {
      return NextResponse.json({ error: e.message }, { status: e.status });
    }
    throw e;
  }

  const run = await prisma.stagingRun.findFirst({
    where: { id: params.runId, mappingId: params.id, userId },
  });
  if (!run) return NextResponse.json({ error: "Not found" }, { status: 404 });

  let body: Record<string, unknown> = {};
  try {
    body = await req.json();
  } catch {
    return NextResponse.json({ error: "Invalid JSON request body" }, { status: 400 });
  }
  const { status, diffResult, mergeResolution, rowsChanged, conflictCount, mergedAt } = body as {
    status?: string;
    diffResult?: unknown;
    mergeResolution?: unknown;
    rowsChanged?: number;
    conflictCount?: number;
    mergedAt?: string;
  };
  const nextDiffResult = typeof diffResult === "undefined"
    ? (run.diffResult as Prisma.InputJsonValue | Prisma.NullableJsonNullValueInput | undefined)
    : (diffResult as Prisma.InputJsonValue | Prisma.NullableJsonNullValueInput);
  const nextMergeResolution = typeof mergeResolution === "undefined"
    ? (run.mergeResolution as Prisma.InputJsonValue | Prisma.NullableJsonNullValueInput | undefined)
    : (mergeResolution as Prisma.InputJsonValue | Prisma.NullableJsonNullValueInput);
  let appliedCellCount = 0;
  let appendedRowCount = 0;

  if (status === "merged") {
    const mapping = await prisma.mapping.findFirst({ where: { id: params.id, userId } });
    if (!mapping?.smartsheetSheetId) {
      return NextResponse.json({ error: "Mapping is not connected to a Smartsheet target" }, { status: 400 });
    }

    const user = await prisma.user.findUnique({ where: { id: userId }, select: { smartsheetToken: true } });
    if (!user?.smartsheetToken) {
      return NextResponse.json({ error: "Smartsheet not connected" }, { status: 400 });
    }

    const cells = (Array.isArray(mergeResolution) ? mergeResolution : Array.isArray(diffResult) ? diffResult : []) as DiffCell[];
    appliedCellCount = cells.filter((c) => (c.resolution ?? "use_staging") !== "keep_production").length;

    const [sheetMeta, prodRows] = await Promise.all([
      getSheet(user.smartsheetToken, mapping.smartsheetSheetId),
      getSheetRows(user.smartsheetToken, mapping.smartsheetSheetId),
    ]);

    const colIdByTitle = new Map<string, number>();
    sheetMeta.columns.forEach((c) => colIdByTitle.set(String(c.title), Number(c.id)));

    const mappingVersion = await prisma.mappingVersion.findUnique({
      where: { id: run.mappingVersionId },
      select: { connections: true },
    });

    const previousStagingRun = await prisma.stagingRun.findFirst({
      where: { mappingId: params.id, userId, id: { not: params.runId } },
      orderBy: { createdAt: "desc" },
      select: { stagingExcelData: true },
    });
    const previousStagingData = (previousStagingRun?.stagingExcelData && typeof previousStagingRun.stagingExcelData === "object")
      ? (previousStagingRun.stagingExcelData as Record<string, unknown>)
      : null;
    const previousDynamicResolvedName = typeof previousStagingData?._dynamicTargetColumnResolvedName === "string"
      ? String(previousStagingData._dynamicTargetColumnResolvedName).trim()
      : "";
    const dynamicTargetCfg = (mappingVersion?.connections as {
      meta?: { dynamicTargetColumn?: DynamicTargetColumnConfig };
    } | null)?.meta?.dynamicTargetColumn;
    const dynamicColumnEnabled = dynamicTargetCfg?.enabled === true;
    const stagingData = (run.stagingExcelData && typeof run.stagingExcelData === "object")
      ? (run.stagingExcelData as Record<string, unknown>)
      : null;
    const activeTargetLabels = new Set(
      Array.isArray(stagingData?._activeTargetLabels)
        ? stagingData._activeTargetLabels
            .map((value) => String(value ?? "").trim())
            .filter(Boolean)
        : []
    );
    const activeTargetColumnIds = new Set(
      Array.isArray(stagingData?._activeTargetColumnIds)
        ? stagingData._activeTargetColumnIds
            .map((value) => Number(value))
            .filter((value) => Number.isFinite(value))
        : []
    );
    const resolvedFromStaging = typeof stagingData?._dynamicTargetColumnResolvedName === "string"
      ? String(stagingData?._dynamicTargetColumnResolvedName)
      : "";
    const expectedDynamicColumnName = dynamicColumnEnabled
      ? (resolvedFromStaging || formatDynamicColumnName(String(dynamicTargetCfg?.nameTemplate ?? "Enrollments {{DATE}}"), {
          existingTitles: sheetMeta.columns.map((col) => col.title),
          ensureUnique: true,
        }))
      : "";

    let dynamicColumnId: number | null = null;
    if (dynamicColumnEnabled) {
      const position = dynamicTargetCfg?.columnPosition ?? "end";
      const customIndex =
        typeof dynamicTargetCfg?.customColumnNumber === "number" && dynamicTargetCfg.customColumnNumber > 0
          ? Math.floor(dynamicTargetCfg.customColumnNumber) - 1
          : 0;
      const resolvedIndex =
        position === "start"
          ? 0
          : position === "custom"
            ? Math.max(0, Math.min(sheetMeta.columns.length, customIndex))
            : sheetMeta.columns.length;
      const created = await addSheetColumn(user.smartsheetToken, mapping.smartsheetSheetId, {
        title: expectedDynamicColumnName,
        type: "TEXT_NUMBER",
        index: resolvedIndex,
      });
      dynamicColumnId = created.id;
      colIdByTitle.set(expectedDynamicColumnName, dynamicColumnId);
    }

    const mappingConnections = (mappingVersion?.connections ?? null) as {
      nodes?: Array<{ id?: string; type?: string; data?: Record<string, unknown> }>;
      edges?: Array<Record<string, unknown>>;
      meta?: Record<string, unknown>;
    } | null;
    const latestNodes = Array.isArray(mappingConnections?.nodes) ? mappingConnections.nodes : [];
    const syntheticTargetColumnIds = new Map<string, number>();

    for (let index = 0; index < latestNodes.length; index += 1) {
      const node = latestNodes[index];
      if (node?.type !== "ssCol") continue;
      const nodeData = node.data ?? {};
      const label = String(nodeData.label ?? "").trim();
      const rawColId = nodeData.colId;
      const isSynthetic = nodeData.synthetic === true || (typeof rawColId === "string" && rawColId.startsWith("ss_manual_"));
      if (!label || !isSynthetic) continue;

      const existingColId = colIdByTitle.get(label);
      if (existingColId) {
        syntheticTargetColumnIds.set(label, existingColId);
        continue;
      }

      const syntheticPosition = typeof nodeData.columnPosition === "string" ? nodeData.columnPosition : "end";
      const syntheticCustomNumber = typeof nodeData.customColumnNumber === "number" && nodeData.customColumnNumber > 0
        ? Math.floor(nodeData.customColumnNumber)
        : 1;
      const nodeIndex =
        syntheticPosition === "start"
          ? 0
          : syntheticPosition === "custom"
            ? Math.max(0, Math.min(sheetMeta.columns.length, syntheticCustomNumber - 1))
            : sheetMeta.columns.length;
      const created = await addSheetColumn(user.smartsheetToken, mapping.smartsheetSheetId, {
        title: label,
        type: typeof nodeData.colType === "string" && nodeData.colType ? String(nodeData.colType) : "TEXT_NUMBER",
        index: nodeIndex,
      });
      syntheticTargetColumnIds.set(label, created.id);
      colIdByTitle.set(label, created.id);
    }

    const pendingDeletionColIds: number[] = [];
    if (dynamicColumnEnabled && previousDynamicResolvedName && previousDynamicResolvedName !== expectedDynamicColumnName) {
      const prevDynId = colIdByTitle.get(previousDynamicResolvedName);
      if (typeof prevDynId === "number" && Number.isFinite(prevDynId) && prevDynId !== dynamicColumnId) {
        pendingDeletionColIds.push(prevDynId);
      }
    }

    const rowIdByNumber = new Map<number, number>();
    prodRows.forEach((r, idx) => {
      rowIdByNumber.set(idx + 1, r.id);
      if (typeof r.rowNumber === "number") rowIdByNumber.set(r.rowNumber, r.id);
    });

    const updatesByRow = new Map<number, { id: number; cells: { columnId: number; value: string | number | boolean | null }[] }>();
    const appendsByRowIndex = new Map<number, { toBottom: boolean; cells: { columnId: number; value: string | number | boolean | null }[] }>();

    for (const cell of cells) {
      const shouldApply = (cell.resolution ?? "use_staging") !== "keep_production";
      if (!shouldApply) continue;

      if (cell.action === "skip") continue;

      const rowIdx = Number(cell.row) - 1;
      if (rowIdx < 0) continue;

      const cellColumnLabel = String(cell.column);
      const columnId =
        syntheticTargetColumnIds.get(cellColumnLabel) ??
        (dynamicColumnEnabled && cellColumnLabel === expectedDynamicColumnName ? (dynamicColumnId ?? undefined) : undefined) ??
        (typeof cell.columnId === "number" && Number.isFinite(cell.columnId) ? cell.columnId : undefined) ??
        colIdByTitle.get(cellColumnLabel);
      if (!columnId) continue;
      if (activeTargetLabels.size > 0 && !activeTargetLabels.has(cellColumnLabel) && cellColumnLabel !== expectedDynamicColumnName) {
        continue;
      }
      const isTrackedByLabel = activeTargetLabels.has(cellColumnLabel) || cellColumnLabel === expectedDynamicColumnName;
      if (activeTargetColumnIds.size > 0 && !isTrackedByLabel && !activeTargetColumnIds.has(columnId) && columnId !== dynamicColumnId) {
        continue;
      }

      const explicitRowId = typeof cell.rowId === "number" && Number.isFinite(cell.rowId) ? cell.rowId : undefined;
      const resolvedRowId =
        cell.action === "append"
          ? undefined
          : (explicitRowId ?? rowIdByNumber.get(Number(cell.row)));

      if (resolvedRowId) {
        const targetRow = prodRows.find((r) => r.id === resolvedRowId);
        if (!targetRow) continue;
        const existing = updatesByRow.get(targetRow.id) ?? { id: targetRow.id, cells: [] };
        existing.cells = existing.cells.filter((candidate) => candidate.columnId !== columnId);
        existing.cells.push({
          columnId,
          value: (cell.stagingValue as string | number | boolean | null) ?? null,
        });
        updatesByRow.set(targetRow.id, existing);
      } else {
        const append = appendsByRowIndex.get(rowIdx) ?? { toBottom: true, cells: [] };
        append.cells = append.cells.filter((candidate) => candidate.columnId !== columnId);
        append.cells.push({
          columnId,
          value: (cell.stagingValue as string | number | boolean | null) ?? null,
        });
        appendsByRowIndex.set(rowIdx, append);
      }
    }

    const updates = Array.from(updatesByRow.values()).filter((r) => r.cells.length > 0);
    if (updates.length > 0) {
      await updateSheetRows(user.smartsheetToken, mapping.smartsheetSheetId, updates);
    }

    const appends = Array.from(appendsByRowIndex.entries())
      .sort((a, b) => a[0] - b[0])
      .map(([, row]) => row)
      .filter((r) => r.cells.length > 0);
    appendedRowCount = appends.length;

    if (appends.length > 0) {
      await addSheetRows(user.smartsheetToken, mapping.smartsheetSheetId, appends);
    }

    for (const colId of pendingDeletionColIds) {
      try {
        await deleteSheetColumn(user.smartsheetToken, mapping.smartsheetSheetId, colId);
      } catch {
        // no-op
      }
    }
  }

  const updated = await prisma.stagingRun.update({
    where: { id: params.runId },
    data: {
      status: status ?? run.status,
      diffResult: nextDiffResult,
      mergeResolution: nextMergeResolution,
      rowsChanged: rowsChanged ?? run.rowsChanged,
      conflictCount: conflictCount ?? run.conflictCount,
      mergedAt: mergedAt ? new Date(mergedAt) : run.mergedAt,
    },
  });

  if (status === "merged") {
    return NextResponse.json({ ...updated, appliedCellCount, appendedRowCount });
  }

  return NextResponse.json(updated);
}
