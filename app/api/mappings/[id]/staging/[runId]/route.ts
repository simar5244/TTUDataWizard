import { NextRequest, NextResponse } from "next/server";
import { getServerSession } from "next-auth";
import { authOptions } from "@/lib/auth";
import { prisma } from "@/lib/prisma";

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

  const run = await prisma.stagingRun.findFirst({
    where: { id: params.runId, mappingId: params.id, userId },
  });
  if (!run) return NextResponse.json({ error: "Not found" }, { status: 404 });

  const { status, diffResult, mergeResolution, rowsChanged, conflictCount, mergedAt } = await req.json();

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

  return NextResponse.json(updated);
}
