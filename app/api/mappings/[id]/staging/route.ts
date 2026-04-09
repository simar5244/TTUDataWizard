import { NextRequest, NextResponse } from "next/server";
import { getServerSession } from "next-auth";
import { authOptions } from "@/lib/auth";
import { prisma } from "@/lib/prisma";
import { getSheet, getSheetRows, type SmartsheetRow } from "@/lib/smartsheet";
import { assertEditAllowed, SecurityPolicyError } from "@/lib/security";

interface RunOptions {
  hierarchyAware?: boolean;
  protectFormulaCells?: boolean;
  protectParentSummaryRows?: boolean;
  excludedRowPatterns?: string[];
  columnExcludedRowPatterns?: Record<string, string[]>;
  excludedRowNumbers?: number[];
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
      };

      const rowById = new Map<number, SmartsheetRow>();
      prodRows.forEach((row) => rowById.set(row.id, row));

      const parentIds = new Set<number>();
      prodRows.forEach((row) => {
        if (typeof row.parentId === "number") parentIds.add(row.parentId);
      });

      const titleColId = Number(sheetMeta.columns[0]?.id ?? NaN);
      const columnFormulaById = new Map<number, string>();
      sheetMeta.columns.forEach((col) => {
        if (typeof col.formula === "string" && col.formula.trim() !== "") {
          columnFormulaById.set(col.id, col.formula);
        }
      });
      const getRowTitle = (row: SmartsheetRow): string => {
        const titleCell = row.cells.find((cell) => cell.columnId === titleColId);
        return String(titleCell?.displayValue ?? titleCell?.value ?? `Row ${row.rowNumber ?? ""}`)
          .trim() || `Row ${row.rowNumber ?? ""}`;
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

      // Parent/group-header detection should only be hierarchy-based.
      // Formula-based protection is handled per-cell below (cellHasFormula).
      const isParentOrSummaryRow = (row: SmartsheetRow): boolean => parentIds.has(row.id);

      const mappedOutputRows = Array.isArray(outputRows)
        ? (outputRows as Record<string, unknown>[])
        : [];

      const targetLabelToColId = new Map<string, number>();
      const connectionNodes = ((currentVersion.connections as { nodes?: Array<{ type?: string; data?: Record<string, unknown> }> } | null)?.nodes ?? []);
      connectionNodes
        .filter((n) => n.type === "ssCol")
        .forEach((n) => {
          const label = String(n.data?.label ?? "").trim();
          const colId = Number(n.data?.colId ?? NaN);
          if (label && Number.isFinite(colId)) targetLabelToColId.set(label, colId);
        });
      sheetMeta.columns.forEach((c) => {
        if (!targetLabelToColId.has(c.title)) targetLabelToColId.set(c.title, c.id);
      });

      const diffCells: DiffCell[] = [];

      mappedOutputRows.forEach((outputRow, sourceIndex) => {
        const targetRow = prodRows[sourceIndex];
        const targetRowPath = targetRow ? getRowPath(targetRow) : null;
        const targetRowIsParentSummary = targetRow ? isParentOrSummaryRow(targetRow) : false;
        const targetRowTitle = targetRow ? getRowTitle(targetRow) : "";
        const rowNumber = targetRow?.rowNumber ?? sourceIndex + 1;
        const matchesExcludedPattern = targetRow
          ? options.excludedRowPatterns.some((pattern) => {
              const p = pattern.trim().toLowerCase();
              if (!p) return false;
              return targetRowTitle.toLowerCase().includes(p) || (targetRowPath || "").toLowerCase().includes(p);
            })
          : false;

        Object.entries(outputRow).forEach(([column, value]) => {
          const columnId = targetLabelToColId.get(column) ?? null;
          if (!columnId) return;

          if (!targetRow) {
            diffCells.push({
              row: sourceIndex + 1,
              rowId: null,
              rowPath: null,
              rowKind: "new_row",
              column,
              columnId,
              productionValue: null,
              stagingValue: value ?? null,
              isConflict: false,
              resolution: "use_staging",
              action: "append",
            });
            return;
          }

          const targetCell = targetRow.cells.find((cell) => cell.columnId === columnId);
          const productionValue = targetCell?.value ?? null;
          const stagingValue = value ?? null;
          const left = productionValue === null || productionValue === undefined ? "" : String(productionValue).trim();
          const right = stagingValue === null || stagingValue === undefined ? "" : String(stagingValue).trim();
          const changed = left !== right;
          if (!changed) return;

          const cellHasFormula =
            (typeof targetCell?.formula === "string" && targetCell.formula.trim() !== "") ||
            columnFormulaById.has(columnId);
          const isLocked = Boolean(targetRow.locked || targetCell?.locked);
          const shouldSkipForParentSummary = options.protectParentSummaryRows && targetRowIsParentSummary;
          const shouldSkipForFormula = options.protectFormulaCells && cellHasFormula;
          const shouldSkipForExcludedPattern = matchesExcludedPattern;
          const shouldSkipForExcludedRowNumber = options.excludedRowNumbers.includes(rowNumber);
          const columnSpecificPatterns = options.columnExcludedRowPatterns[column] ?? [];
          const shouldSkipForColumnPattern = columnSpecificPatterns.some((pattern) => {
            const p = String(pattern || "").trim().toLowerCase();
            if (!p) return false;
            return targetRowTitle.toLowerCase().includes(p) || (targetRowPath || "").toLowerCase().includes(p);
          });

          let skipReason: DiffCell["skipReason"] | undefined;
          if (shouldSkipForFormula) skipReason = "protected_formula";
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
            column,
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
        totalOutputRows: mappedOutputRows.length,
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
        _rowPolicyDiagnostics: diagnostics,
        _rowPolicyApplied: {
          protectFormulaCells: options.protectFormulaCells,
          protectParentSummaryRows: options.protectParentSummaryRows,
          excludedRowPatterns: options.excludedRowPatterns,
          excludedRowNumbers: options.excludedRowNumbers,
        },
      };

      normalizedDiffResult = mappedOutputRows.length > 0 ? diffCells : applyResults || null;
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
