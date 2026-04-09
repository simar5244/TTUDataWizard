import { NextRequest, NextResponse } from "next/server";
import { getServerSession } from "next-auth";
import { authOptions } from "@/lib/auth";
import { prisma } from "@/lib/prisma";

export async function GET(req: NextRequest, { params }: { params: { id: string } }) {
  const session = await getServerSession(authOptions);
  if (!session?.user) return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  const userId = (session.user as { id: string }).id;

  const mapping = await prisma.mapping.findFirst({
    where: { id: params.id, userId },
    include: {
      versions: { orderBy: { versionNumber: "asc" } },
      stagingRuns: { orderBy: { createdAt: "desc" } },
    },
  });

  if (!mapping) return NextResponse.json({ error: "Not found" }, { status: 404 });
  return NextResponse.json(mapping);
}

export async function PUT(req: NextRequest, { params }: { params: { id: string } }) {
  const session = await getServerSession(authOptions);
  if (!session?.user) return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  const userId = (session.user as { id: string }).id;

  const mapping = await prisma.mapping.findFirst({ where: { id: params.id, userId } });
  if (!mapping) return NextResponse.json({ error: "Not found" }, { status: 404 });

  const { name, smartsheetSheetId, smartsheetSheetName, connections, formulas, schemaFingerprint, changeSummary } =
    await req.json();

  const latestVersion = await prisma.mappingVersion.findFirst({
    where: { mappingId: params.id },
    orderBy: { versionNumber: "desc" },
  });

  const newVersionNumber = (latestVersion?.versionNumber ?? 0) + 1;

  const newVersion = await prisma.mappingVersion.create({
    data: {
      mappingId: params.id,
      userId,
      versionNumber: newVersionNumber,
      connections: connections || {},
      formulas: formulas || {},
      schemaFingerprint: schemaFingerprint || {},
      changeSummary: changeSummary || `Version ${newVersionNumber}`,
    },
  });

  const updated = await prisma.mapping.update({
    where: { id: params.id },
    data: {
      name: name ?? mapping.name,
      smartsheetSheetId: smartsheetSheetId ?? mapping.smartsheetSheetId,
      smartsheetSheetName: smartsheetSheetName ?? mapping.smartsheetSheetName,
      currentVersionId: newVersion.id,
    },
    include: { versions: { orderBy: { versionNumber: "asc" } } },
  });

  return NextResponse.json(updated);
}

export async function DELETE(req: NextRequest, { params }: { params: { id: string } }) {
  const session = await getServerSession(authOptions);
  if (!session?.user) return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  const userId = (session.user as { id: string }).id;

  const mapping = await prisma.mapping.findFirst({ where: { id: params.id, userId } });
  if (!mapping) return NextResponse.json({ error: "Not found" }, { status: 404 });

  await prisma.mapping.delete({ where: { id: params.id } });
  return NextResponse.json({ success: true });
}
