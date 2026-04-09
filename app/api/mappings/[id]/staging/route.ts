import { NextRequest, NextResponse } from "next/server";
import { getServerSession } from "next-auth";
import { authOptions } from "@/lib/auth";
import { prisma } from "@/lib/prisma";
import { duplicateSheet, getSheetRows } from "@/lib/smartsheet";

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

  const mapping = await prisma.mapping.findFirst({
    where: { id: params.id, userId },
    include: { versions: { orderBy: { versionNumber: "desc" }, take: 1 } },
  });
  if (!mapping) return NextResponse.json({ error: "Not found" }, { status: 404 });

  const { direction, excelData, applyResults } = await req.json();
  const currentVersion = mapping.versions[0];
  if (!currentVersion) return NextResponse.json({ error: "No mapping version" }, { status: 400 });

  let stagingSheetId: string | null = null;
  let snapshotProduction: object | null = null;

  if (direction === "excel_to_ss" && mapping.smartsheetSheetId) {
    const user = await prisma.user.findUnique({ where: { id: userId } });
    if (!user?.smartsheetToken) {
      return NextResponse.json({ error: "Smartsheet not connected" }, { status: 400 });
    }
    const timestamp = new Date().toISOString().replace(/[:.]/g, "-").slice(0, 19);
    const stagingName = `${mapping.smartsheetSheetName || "Sheet"} — Staging ${timestamp}`;

    try {
      const dupResult = await duplicateSheet(user.smartsheetToken, mapping.smartsheetSheetId, stagingName);
      stagingSheetId = String(dupResult.id);

      const prodRows = await getSheetRows(user.smartsheetToken, mapping.smartsheetSheetId);
      snapshotProduction = prodRows;
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
      stagingSheetId,
      stagingExcelData: excelData || null,
      snapshotProduction: snapshotProduction as object,
      diffResult: applyResults || null,
    },
  });

  return NextResponse.json(stagingRun, { status: 201 });
}
