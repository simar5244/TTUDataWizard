import { NextRequest, NextResponse } from "next/server";
import { getServerSession } from "next-auth";
import { authOptions } from "@/lib/auth";
import { prisma } from "@/lib/prisma";
import { getSheet, getSheetRows, type SmartsheetRow } from "@/lib/smartsheet";
import { formatDynamicColumnName } from "@/lib/dynamic-column";
import { assertEditAllowed, SecurityPolicyError } from "@/lib/security";
import { collectDuplicateBaseLabels } from "@/lib/duplicationmech";

interface RunOptions {
  hierarchyAware?: boolean;
  protectFormulaCells?: boolean;
  protectParentSummaryRows?: boolean;
  excludedRowPatterns?: string[];
  columnExcludedRowPatterns?: Record<string, string[]>;
  excludedRowNumbers?: number[];
  skipTopRows?: number;
}

interface DynamicTargetColumnConfig {
  enabled?: boolean;
  sourceLabel?: string;
  sourceNodeId?: string;
  nameTemplate?: string;
  columnPosition?: "start" | "end" | "custom";
  customColumnNumber?: number;
}

interface DiffCell {
  row: number;
  rowId?: number | null;
  rowPath?: string | null;
  rowKind?: "detail" | "parent_or_summary" | "new_row";
  column: string;
  columnId?: number | null;
  productionValue: unknown;
  stagingValue: unknown;
  isConflict: boolean;
  resolution: "keep_production" | "use_staging";
  action: "update" | "append" | "skip";
  skipReason?: "protected_formula" | "protected_parent_summary" | "locked" | "excluded_by_rule";
}

interface RowPolicyDiagnostics {
  totalOutputRows: number;
  totalDiffCells: number;
  skippedCells: number;
  updatedCells: number;
  appendedCells: number;
  skipByReason: Record<string, number>;
  parentRowsDetected: number;
  formulaRowsDetected: number;
  hierarchyRowsDetected: number;
}

export async function GET(req: NextRequest, { params }: { params: { id: string } }) {
  const session = await getServerSession(authOptions);
  if (!session?.user) return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  const userId = (session.user as { id: string }).id;

  const runs = await prisma.stagingRun.findMany({
    where: { mappingId: params.id, userId },
    include: { mappingVersion: { select: { versionNumber: true, changeSummary: true } } },
    orderBy: { createdAt: "desc" },
  });

  return NextResponse.json(runs);
}

export async function POST(req: NextRequest, { params }: { params: { id: string } }) {
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

  const mapping = await prisma.mapping.findFirst({
    where: { id: params.id, userId },
    include: { versions: { orderBy: { versionNumber: "desc" }, take: 1 } },
  });
  if (!mapping) return NextResponse.json({ error: "Not found" }, { status: 404 });

  const { direction, excelData, applyResults, outputRows, runOptions } = await req.json();
  const currentVersion = mapping.versions[0];
  if (!currentVersion) return NextResponse.json({ error: "No mapping version" }, { status: 400 });

  let snapshotProduction: object | null = null;
  let normalizedDiffResult: unknown = applyResults || null;
  let enrichedStagingExcelData: unknown = (excelData && typeof excelData === "object" ? excelData : null);

  if (direction === "excel_to_ss" && mapping.smartsheetSheetId) {
    const user = await prisma.user.findUnique({ where: { id: userId } });
    if (!user?.smartsheetToken) {
      return NextResponse.json({ error: "Smartsheet not connected" }, { status: 400 });
    }
    try {
      const [sheetMeta, prodRows] = await Promise.all([
        getSheet(user.smartsheetToken, mapping.smartsheetSheetId),
        getSheetRows(user.smartsheetToken, mapping.smartsheetSheetId),
      ]);
      snapshotProduction = prodRows;

      const options: Required<RunOptions> = {
        hierarchyAware: runOptions?.hierarchyAware !== false,
        protectFormulaCells: runOptions?.protectFormulaCells !== false,
        protectParentSummaryRows: runOptions?.protectParentSummaryRows !== false,
        excludedRowPatterns: Array.isArray(runOptions?.excludedRowPatterns)
          ? runOptions.excludedRowPatterns.filter((v: string) => typeof v === "string" && v.trim() !== "")
          : [],
        columnExcludedRowPatterns:
          runOptions && typeof runOptions.columnExcludedRowPatterns === "object" && runOptions.columnExcludedRowPatterns !== null
            ? (runOptions.columnExcludedRowPatterns as Record<string, string[]>)
            : {},
        excludedRowNumbers: Array.isArray(runOptions?.excludedRowNumbers)
          ? runOptions.excludedRowNumbers
              .map((v: number) => Number(v))
              .filter((v: number) => Number.isFinite(v) && v > 0)
              .map((v: number) => Math.floor(v))
          : [],
        skipTopRows:
          typeof runOptions?.skipTopRows === "number" && runOptions.skipTopRows > 0
            ? Math.floor(runOptions.skipTopRows)
            : 0,
      };

      const rowById = new Map<number, SmartsheetRow>();
      prodRows.forEach((row) => rowById.set(row.id, row));

      const parentIds = new Set<number>();
      prodRows.forEach((row) => {
        if (typeof row.parentId === "number") parentIds.add(row.parentId);
      });

      const titleColId = Number(sheetMeta.columns[0]?.id ?? NaN);
      const getRowTitle = (row: SmartsheetRow): string => {
        const titleCell = row.cells.find((cell) => cell.columnId === titleColId);
        const directTitle = String(titleCell?.displayValue ?? titleCell?.value ?? "").trim();
        if (directTitle) return directTitle;
        const fallbackTitle = row.cells
          .map((cell) => String(cell.displayValue ?? cell.value ?? "").trim())
          .find((value) => value !== "");
        return fallbackTitle || `Row ${row.rowNumber ?? ""}`;
      };

      const getRowPath = (row: SmartsheetRow): string => {
        const seen = new Set<number>();
        const segments: string[] = [getRowTitle(row)];
        let currentParentId = row.parentId;
        while (typeof currentParentId === "number" && !seen.has(currentParentId)) {
          seen.add(currentParentId);
          const parent = rowById.get(currentParentId);
          if (!parent) break;
          segments.unshift(getRowTitle(parent));
          currentParentId = parent.parentId;
        }
        return segments.join(" > ");
      };

      const rowContainsFormula = (row: SmartsheetRow): boolean =>
        row.cells.some((cell) => typeof cell.formula === "string" && cell.formula.trim() !== "");

      const formulaRowsDetected = prodRows.filter((r) => rowContainsFormula(r)).length;
      const hierarchyRowsDetected = prodRows.filter((r) => typeof r.parentId === "number").length;

      const isParentOrSummaryRow = (row: SmartsheetRow): boolean => parentIds.has(row.id);

      const mappedOutputRows = Array.isArray(outputRows)
        ? (outputRows as Record<string, unknown>[])
        : [];

      const effectiveOutputRows = mappedOutputRows
        .map((row, sourceIndex) => ({ row, sourceIndex }));

      const targetLabelToColId = new Map<string, number>();
      const connMeta = (currentVersion.connections as {
        meta?: { dynamicTargetColumn?: DynamicTargetColumnConfig };
      } | null)?.meta;
      const dynamicTargetCfg = connMeta?.dynamicTargetColumn;
      const dynamicTargetColumnEnabled = dynamicTargetCfg?.enabled === true;
      const dynamicTemplate = String(dynamicTargetCfg?.nameTemplate ?? "Enrollments {{DATE}}");
      const dynamicBaseLabel = dynamicTargetColumnEnabled
        ? formatDynamicColumnName(dynamicTemplate, { ensureUnique: false })
        : "";
      const dynamicTargetColumnLabel = dynamicTargetColumnEnabled
        ? formatDynamicColumnName(dynamicTemplate, {
            existingTitles: sheetMeta.columns.map((col) => col.title),
            ensureUnique: true,
          })
        : "";
      const connectionNodes = ((currentVersion.connections as { nodes?: Array<{ type?: string; data?: Record<string, unknown> }> } | null)?.nodes ?? []);
      const duplicateBaseLabels = collectDuplicateBaseLabels(connectionNodes);
      const normalizeLabel = (value: unknown): string => String(value ?? "").trim().toLowerCase();
      const activeTargetLabels = new Set<string>();
      const activeTargetColumnIds = new Set<number>();
      connectionNodes
        .filter((n) => n.type === "ssCol")
        .forEach((n) => {
          const label = String(n.data?.label ?? "").trim();
          const normalizedLabel = label.toLowerCase();
          const colId = Number(n.data?.colId ?? NaN);
          if (label) activeTargetLabels.add(label);
          if (label && duplicateBaseLabels.normalized.has(normalizedLabel)) {
            // Duplicate-per-run targets are resolved at merge time to the newly created column.
          } else if (Number.isFinite(colId)) {
            activeTargetColumnIds.add(colId);
          }
          if (label && Number.isFinite(colId)) targetLabelToColId.set(label, colId);
        });
      sheetMeta.columns.forEach((c) => {
        if (!targetLabelToColId.has(c.title)) targetLabelToColId.set(c.title, c.id);
      });
      if (dynamicTargetColumnEnabled && dynamicTargetColumnLabel) {
        activeTargetLabels.add(dynamicTargetColumnLabel);
      }
      console.log("[staging:dsp]", { dynamicTargetColumnEnabled, dynamicTargetColumnLabel, dynamicBaseLabel, firstOutputRowKeys: mappedOutputRows[0] ? Object.keys(mappedOutputRows[0]) : [] });

      const shouldSkipTargetRowForAlignment = (row: SmartsheetRow): boolean => {
        const rowNumber = row.rowNumber ?? 0;
        if (options.skipTopRows > 0 && rowNumber > 0 && rowNumber <= options.skipTopRows) return true;
        const targetRowIsParentSummary = isParentOrSummaryRow(row);
        const targetRowTitle = getRowTitle(row);
        const targetRowPath = getRowPath(row);
        const isTopLevelFormulaRow = options.protectFormulaCells && typeof row.parentId !== "number" && rowContainsFormula(row);
        const shouldSkipForExcludedPattern = options.excludedRowPatterns.some((pattern) => {
          const p = pattern.trim().toLowerCase();
          if (!p) return false;
          return targetRowTitle.toLowerCase().includes(p) || targetRowPath.toLowerCase().includes(p);
        });
        const shouldSkipForExcludedRowNumber = options.excludedRowNumbers.includes(rowNumber);
        const shouldSkipForParentSummary = options.protectParentSummaryRows && targetRowIsParentSummary;
        return shouldSkipForParentSummary || shouldSkipForExcludedRowNumber || shouldSkipForExcludedPattern || isTopLevelFormulaRow;
      };

      const rowDebugSummary = prodRows.slice(0, 20).map((row) => {
        const rowNumber = row.rowNumber ?? 0;
        const title = getRowTitle(row);
        const isParentSummary = isParentOrSummaryRow(row);
        const hasFormula = rowContainsFormula(row);
        const excludedByPattern = options.excludedRowPatterns.some((pattern) => {
          const p = pattern.trim().toLowerCase();
          if (!p) return false;
          const path = getRowPath(row).toLowerCase();
          return title.toLowerCase().includes(p) || path.includes(p);
        });
        return {
          rowNumber,
          rowId: row.id,
          title,
          parentId: row.parentId ?? null,
          expanded: row.expanded ?? null,
          isParentSummary,
          hasFormula,
          excludedByNumber: options.excludedRowNumbers.includes(rowNumber),
          excludedByPattern,
          skippedForAlignment: shouldSkipTargetRowForAlignment(row),
        };
      });
      console.log("[staging] row classification", JSON.stringify({
        totalOutputRows: mappedOutputRows.length,
        totalProdRows: prodRows.length,
        options,
        firstRows: rowDebugSummary,
      }));

      const alignedTargetRows: Array<SmartsheetRow | null> = [];
      let targetCursor = 0;
      for (let sourceIndex = 0; sourceIndex < effectiveOutputRows.length; sourceIndex += 1) {
        let matched: SmartsheetRow | null = null;
        while (targetCursor < prodRows.length) {
          const candidate = prodRows[targetCursor];
          targetCursor += 1;
          if (shouldSkipTargetRowForAlignment(candidate)) continue;
          matched = candidate;
          break;
        }
        alignedTargetRows.push(matched);
      }

      console.log("[staging] alignment preview", JSON.stringify(
        effectiveOutputRows.slice(0, 20).map(({ row, sourceIndex }, outputIndex) => {
          const targetRow = alignedTargetRows[outputIndex] ?? null;
          const nonEmptyKeys = Object.entries(row)
            .filter(([, value]) => value !== null && value !== undefined && String(value).trim() !== "")
            .map(([key]) => key)
            .slice(0, 5);
          return {
            sourceIndex,
            outputIndex,
            sourceHasValues: nonEmptyKeys.length > 0,
            nonEmptyKeys,
            targetRowNumber: targetRow?.rowNumber ?? null,
            targetRowId: targetRow?.id ?? null,
            targetRowTitle: targetRow ? getRowTitle(targetRow) : null,
            targetIsParentSummary: targetRow ? isParentOrSummaryRow(targetRow) : false,
            targetHasFormula: targetRow ? rowContainsFormula(targetRow) : false,
          };
        })
      ));

      const diffCells: DiffCell[] = [];

      effectiveOutputRows.forEach(({ row: outputRow, sourceIndex }, outputIndex) => {
        const targetRow = alignedTargetRows[outputIndex] ?? null;
        const targetRowPath = targetRow ? getRowPath(targetRow) : null;
        const targetRowIsParentSummary = targetRow ? isParentOrSummaryRow(targetRow) : false;
        const targetRowTitle = targetRow ? getRowTitle(targetRow) : "";
        const rowNumber = targetRow?.rowNumber ?? sourceIndex + 1;
        const normalizedOutputEntries = Object.entries(outputRow).map(([column, value]) => {
          const normalizedColumn =
            dynamicTargetColumnEnabled && (column === dynamicTargetColumnLabel || column === dynamicBaseLabel)
              ? dynamicTargetColumnLabel
              : column;
          return [normalizedColumn, value] as const;
        });
        const normalizedOutputColumns = new Set(normalizedOutputEntries.map(([column]) => column));
        const matchesExcludedPattern = targetRow
          ? options.excludedRowPatterns.some((pattern) => {
              const p = pattern.trim().toLowerCase();
              if (!p) return false;
              return targetRowTitle.toLowerCase().includes(p) || (targetRowPath || "").toLowerCase().includes(p);
            })
          : false;

        const rowEntries = targetRow
          ? [
              ...normalizedOutputEntries,
              ...Array.from(activeTargetLabels)
                .filter((column) => !normalizedOutputColumns.has(column))
                .filter((column) => !(dynamicTargetColumnEnabled && column === dynamicTargetColumnLabel))
                .map((column) => [column, null] as const),
            ]
          : normalizedOutputEntries;

        rowEntries.forEach(([normalizedColumn, value]) => {
          if (!activeTargetLabels.has(normalizedColumn) && !(dynamicTargetColumnEnabled && normalizedColumn === dynamicTargetColumnLabel)) {
            return;
          }
          const isDuplicateBaseColumn = duplicateBaseLabels.normalized.has(normalizeLabel(normalizedColumn));
          const columnId = targetLabelToColId.get(normalizedColumn) ?? null;

          if (!targetRow) {
            diffCells.push({
              row: sourceIndex + 1,
              rowId: null,
              rowPath: null,
              rowKind: "new_row",
              column: normalizedColumn,
              columnId,
              productionValue: null,
              stagingValue: value ?? null,
              isConflict: false,
              resolution: "use_staging",
              action: "append",
            });
            return;
          }

          const targetCell = !isDuplicateBaseColumn && typeof columnId === "number"
            ? targetRow.cells.find((cell) => cell.columnId === columnId)
            : undefined;
          const productionValue = isDuplicateBaseColumn ? null : (targetCell?.value ?? null);
          const stagingValue = value ?? null;
          const left = productionValue === null || productionValue === undefined ? "" : String(productionValue).trim();
          const right = stagingValue === null || stagingValue === undefined ? "" : String(stagingValue).trim();
          const changed = isDuplicateBaseColumn ? true : left !== right;
          if (!changed) return;

          const cellHasFormula = isDuplicateBaseColumn
            ? false
            : (typeof targetCell?.formula === "string" && targetCell.formula.trim() !== "");
          const isLocked = isDuplicateBaseColumn ? false : Boolean(targetRow.locked || targetCell?.locked);
          const shouldSkipForTopRow = options.skipTopRows > 0 && rowNumber > 0 && rowNumber <= options.skipTopRows;
          const shouldSkipForParentSummary = options.protectParentSummaryRows && targetRowIsParentSummary;
          const shouldSkipForFormula = options.protectFormulaCells && cellHasFormula;
          const shouldSkipForExcludedPattern = matchesExcludedPattern;
          const shouldSkipForExcludedRowNumber = options.excludedRowNumbers.includes(rowNumber);
          const columnSpecificPatterns = options.columnExcludedRowPatterns[normalizedColumn] ?? [];
          const shouldSkipForColumnPattern = columnSpecificPatterns.some((pattern) => {
            const p = String(pattern || "").trim().toLowerCase();
            if (!p) return false;
            return targetRowTitle.toLowerCase().includes(p) || (targetRowPath || "").toLowerCase().includes(p);
          });

          let skipReason: DiffCell["skipReason"] | undefined;
          if (shouldSkipForTopRow) skipReason = "excluded_by_rule";
          else if (shouldSkipForFormula) skipReason = "protected_formula";
          else if (shouldSkipForParentSummary) skipReason = "protected_parent_summary";
          else if (shouldSkipForExcludedRowNumber) skipReason = "excluded_by_rule";
          else if (shouldSkipForExcludedPattern) skipReason = "excluded_by_rule";
          else if (shouldSkipForColumnPattern) skipReason = "excluded_by_rule";
          else if (isLocked) skipReason = "locked";

          diffCells.push({
            row: sourceIndex + 1,
            rowId: targetRow.id,
            rowPath: targetRowPath,
            rowKind: targetRowIsParentSummary ? "parent_or_summary" : "detail",
            column: normalizedColumn,
            columnId,
            productionValue,
            stagingValue,
            isConflict: false,
            resolution: skipReason ? "keep_production" : "use_staging",
            action: skipReason ? "skip" : "update",
            skipReason,
          });
        });
      });

      const diagnostics: RowPolicyDiagnostics = {
        totalOutputRows: effectiveOutputRows.length,
        totalDiffCells: diffCells.length,
        skippedCells: diffCells.filter((c) => c.action === "skip").length,
        updatedCells: diffCells.filter((c) => c.action === "update").length,
        appendedCells: diffCells.filter((c) => c.action === "append").length,
        skipByReason: diffCells
          .filter((c) => c.action === "skip")
          .reduce((acc, c) => {
            const key = c.skipReason || "unknown";
            acc[key] = (acc[key] || 0) + 1;
            return acc;
          }, {} as Record<string, number>),
        parentRowsDetected: parentIds.size,
        formulaRowsDetected,
        hierarchyRowsDetected,
      };

      const existingStaging = (excelData && typeof excelData === "object" ? excelData : {}) as Record<string, unknown>;
      const enrichedStagingData = {
        ...existingStaging,
        _dynamicTargetColumnResolvedName: dynamicTargetColumnEnabled ? dynamicTargetColumnLabel : null,
        _activeTargetLabels: Array.from(activeTargetLabels),
        _activeTargetColumnIds: Array.from(activeTargetColumnIds),
        _duplicateBaseTargetLabels: Array.from(duplicateBaseLabels.raw),
        _rowPolicyDiagnostics: diagnostics,
        _rowPolicyApplied: {
          protectFormulaCells: options.protectFormulaCells,
          protectParentSummaryRows: options.protectParentSummaryRows,
          excludedRowPatterns: options.excludedRowPatterns,
          excludedRowNumbers: options.excludedRowNumbers,
        },
      };

      normalizedDiffResult = effectiveOutputRows.length > 0 ? diffCells : applyResults || null;
      snapshotProduction = {
        rows: prodRows,
        diagnostics,
      };
      enrichedStagingExcelData = enrichedStagingData as object;
    } catch (e) {
      return NextResponse.json({ error: `Smartsheet error: ${(e as Error).message}` }, { status: 500 });
    }
  }

  const stagingRun = await prisma.stagingRun.create({
    data: {
      mappingId: params.id,
      mappingVersionId: currentVersion.id,
      userId,
      direction,
      status: "open",
      stagingSheetId: null,
      stagingExcelData: enrichedStagingExcelData as object,
      snapshotProduction: snapshotProduction as object,
      diffResult: normalizedDiffResult as object,
    },
  });

  return NextResponse.json(stagingRun, { status: 201 });
}
