import { NextRequest, NextResponse } from "next/server";
import { getServerSession } from "next-auth";
import { authOptions } from "@/lib/auth";
import { prisma } from "@/lib/prisma";
import { assertEditAllowed, SecurityPolicyError } from "@/lib/security";

export async function GET(req: NextRequest, { params }: { params: { id: string } }) {
  const session = await getServerSession(authOptions);
  if (!session?.user) return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  const userId = (session.user as { id: string }).id;

  const runs = await prisma.$queryRaw<Array<Record<string, unknown>>>
  `SELECT *
   FROM "MappingRun"
   WHERE "mappingId" = ${params.id} AND "userId" = ${userId}
   ORDER BY "createdAt" DESC
   LIMIT 50`;

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

  try {
    const { mappingVersionId, direction, inputFileName, inputData, outputData, rowCount, changeSet, changedColumns, changedCellCount, changedRowCount } = await req.json();
    const id = `mr_${Date.now()}_${Math.random().toString(36).slice(2, 10)}`;

    const inserted = await prisma.$queryRaw<Array<Record<string, unknown>>>
    `INSERT INTO "MappingRun" (
      "id", "mappingId", "mappingVersionId", "userId", "direction", "inputFileName", "inputData", "outputData", "changeSet", "changedColumns", "changedCellCount", "changedRowCount", "rowCount", "createdAt"
    ) VALUES (
      ${id}, ${params.id}, ${mappingVersionId}, ${userId}, ${direction ?? "excel_to_excel"}, ${inputFileName ?? null}, ${JSON.stringify(inputData ?? null)}::jsonb, ${JSON.stringify(outputData)}::jsonb, ${JSON.stringify(changeSet ?? [])}::jsonb, ${JSON.stringify(changedColumns ?? [])}::jsonb, ${Number(changedCellCount) || 0}, ${Number(changedRowCount) || 0}, ${Number(rowCount) || 0}, NOW()
    ) RETURNING *`;

    const run = inserted[0];
    if (!run) throw new Error("Failed to persist mapping run");

    return NextResponse.json(run, { status: 201 });
  } catch (e) {
    console.error("Mapping run save error:", e);
    return NextResponse.json({
      error: (e as Error).message || "Failed to save run",
      hint: "If this persists, run `npx prisma db push && npx prisma generate` and restart dev server.",
    }, { status: 500 });
  }
}
