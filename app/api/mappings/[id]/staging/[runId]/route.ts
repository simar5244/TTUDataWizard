import { NextRequest, NextResponse } from "next/server";
import { getServerSession } from "next-auth";
import { authOptions } from "@/lib/auth";
import { prisma } from "@/lib/prisma";
import { assertEditAllowed, SecurityPolicyError } from "@/lib/security";
import { addSheetRows, getSheet, getSheetRows, updateSheetRows } from "@/lib/smartsheet";

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

  const { status, diffResult, mergeResolution, rowsChanged, conflictCount, mergedAt } = await req.json();
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

      const columnId =
        (typeof cell.columnId === "number" && Number.isFinite(cell.columnId) ? cell.columnId : undefined) ??
        colIdByTitle.get(String(cell.column));
      if (!columnId) continue;

      const explicitRowId = typeof cell.rowId === "number" && Number.isFinite(cell.rowId) ? cell.rowId : undefined;
      const resolvedRowId = explicitRowId ?? rowIdByNumber.get(Number(cell.row));

      if (resolvedRowId) {
        const targetRow = prodRows.find((r) => r.id === resolvedRowId);
        if (!targetRow) continue;
        const existing = updatesByRow.get(targetRow.id) ?? { id: targetRow.id, cells: [] };
        existing.cells.push({
          columnId,
          value: (cell.stagingValue as string | number | boolean | null) ?? null,
        });
        updatesByRow.set(targetRow.id, existing);
      } else {
        const append = appendsByRowIndex.get(rowIdx) ?? { toBottom: true, cells: [] };
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
  }

  const updated = await prisma.stagingRun.update({
    where: { id: params.runId },
    data: {
      status: status ?? run.status,
      diffResult: diffResult ?? run.diffResult,
      mergeResolution: mergeResolution ?? run.mergeResolution,
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
