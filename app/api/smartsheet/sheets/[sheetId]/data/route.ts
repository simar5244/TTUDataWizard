import { NextRequest, NextResponse } from "next/server";
import { getServerSession } from "next-auth";
import { authOptions } from "@/lib/auth";
import { prisma } from "@/lib/prisma";
import { getSheet, getSheetRows } from "@/lib/smartsheet";
import type { ExcelSheet, ExcelColumn } from "@/lib/excel";

export async function GET(req: NextRequest, { params }: { params: { sheetId: string } }) {
  const session = await getServerSession(authOptions);
  if (!session?.user) return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  const userId = (session.user as { id: string }).id;

  const user = await prisma.user.findUnique({ where: { id: userId } });
  if (!user?.smartsheetToken) {
    return NextResponse.json({ error: "Smartsheet not connected" }, { status: 400 });
  }

  try {
    const [sheetMeta, rows] = await Promise.all([
      getSheet(user.smartsheetToken, params.sheetId),
      getSheetRows(user.smartsheetToken, params.sheetId),
    ]);

    const rowById = new Map<number, (typeof rows)[number]>();
    rows.forEach((r) => rowById.set(r.id, r));

    const parentIds = new Set<number>();
    rows.forEach((r) => {
      if (typeof r.parentId === "number") parentIds.add(r.parentId);
    });

    const titleColId = Number(sheetMeta.columns[0]?.id ?? NaN);
    const getRowTitle = (row: (typeof rows)[number]) => {
      const titleCell = row.cells.find((c) => c.columnId === titleColId);
      return String(titleCell?.displayValue ?? titleCell?.value ?? `Row ${row.rowNumber ?? ""}`).trim();
    };
    const getAncestors = (row: (typeof rows)[number]) => {
      const chain: (typeof rows)[number][] = [];
      const seen = new Set<number>();
      let currentParentId = row.parentId;
      while (typeof currentParentId === "number" && !seen.has(currentParentId)) {
        seen.add(currentParentId);
        const parent = rowById.get(currentParentId);
        if (!parent) break;
        chain.unshift(parent);
        currentParentId = parent.parentId;
      }
      return chain;
    };

    const columns: ExcelColumn[] = sheetMeta.columns.map((col, index) => {
      const sampleValues = rows.slice(0, 10).map((r) => {
        const cell = r.cells.find((c) => c.columnId === col.id);
        return cell?.value ?? null;
      });

      // Detect data type
      const nonNull = sampleValues.filter((v) => v !== null && v !== undefined && v !== "");
      let dataType: ExcelColumn["dataType"] = "string";
      if (nonNull.length === 0) {
        dataType = "empty";
      } else {
        const numericCount = nonNull.filter((v) => typeof v === "number" || !isNaN(Number(v))).length;
        if (numericCount / nonNull.length > 0.8) {
          dataType = "number";
        } else {
          const dateCount = nonNull.filter((v) => {
            const s = String(v);
            return /^\d{1,4}[-/]\d{1,2}[-/]\d{1,4}/.test(s) || !isNaN(Date.parse(s));
          }).length;
          if (dateCount / nonNull.length > 0.7) dataType = "date";
        }
      }

      return {
        key: `col_${index}`,
        header: col.title,
        index,
        dataType,
        sampleValues,
      };
    });

    // Convert rows to data format
    const data = rows.map((row) => {
      const entry: Record<string, string | number | boolean | null> = {};
      sheetMeta.columns.forEach((col) => {
        const cell = row.cells.find((c) => c.columnId === col.id);
        entry[col.title] = cell?.value ?? null;
      });
      return entry;
    });

    const hierarchyRows = rows.map((row) => {
      const values: Record<string, string | number | boolean | null> = {};
      sheetMeta.columns.forEach((col) => {
        const cell = row.cells.find((c) => c.columnId === col.id);
        values[col.title] = cell?.value ?? null;
      });
      const ancestors = getAncestors(row);
      const depth = ancestors.length;
      const path = [...ancestors.map(getRowTitle), getRowTitle(row)].filter(Boolean).join(" > ");
      const sectionPath = ancestors.map(getRowTitle).filter(Boolean).join(" > ") || "Top level";
      return {
        rowId: row.id,
        rowNumber: row.rowNumber ?? null,
        parentId: row.parentId ?? null,
        depth,
        isParent: parentIds.has(row.id),
        path,
        sectionPath,
        values,
      };
    });

    const hasHierarchy = hierarchyRows.some((r) => r.depth > 0 || r.isParent);

    const excelSheet: ExcelSheet = {
      name: sheetMeta.name,
      columns,
      rowCount: data.length,
      data,
    };

    return NextResponse.json({
      sheet: excelSheet,
      sourceId: params.sheetId,
      type: "smartsheet",
      hasHierarchy,
      hierarchyRows,
      lastRefreshed: new Date().toISOString(),
    });
  } catch (e) {
    return NextResponse.json({ error: (e as Error).message }, { status: 500 });
  }
}
