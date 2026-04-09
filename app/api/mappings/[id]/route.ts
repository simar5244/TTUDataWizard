import { NextRequest, NextResponse } from "next/server";
import { getServerSession } from "next-auth";
import { authOptions } from "@/lib/auth";
import { prisma } from "@/lib/prisma";
import { assertDeleteAllowed, assertEditAllowed, SecurityPolicyError } from "@/lib/security";

export async function GET(req: NextRequest, { params }: { params: Promise<{ id: string }> }) {
  const session = await getServerSession(authOptions);
  if (!session?.user) return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  const userId = (session.user as { id: string }).id;

  try {
    const { id } = await params;

    const mapping = await prisma.mapping.findFirst({
      where: { id, userId },
      include: {
        versions: { orderBy: { versionNumber: "asc" } },
        stagingRuns: { orderBy: { createdAt: "desc" } },
      },
    });

    if (!mapping) return NextResponse.json({ error: "Not found" }, { status: 404 });
    const autoPushRows = await prisma.$queryRaw<Array<{ autoPush: boolean | null }>>`
      SELECT "autoPush"
      FROM "Mapping"
      WHERE "id" = ${id}
      LIMIT 1
    `;
    return NextResponse.json({ ...mapping, autoPush: autoPushRows[0]?.autoPush ?? false });
  } catch (e) {
    return NextResponse.json({ error: (e as Error).message }, { status: 500 });
  }
}

export async function PUT(req: NextRequest, { params }: { params: Promise<{ id: string }> }) {
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
    const { id } = await params;

    const mapping = await prisma.mapping.findFirst({ where: { id, userId } });
    if (!mapping) return NextResponse.json({ error: "Not found" }, { status: 404 });

    const autoPushRows = await prisma.$queryRaw<Array<{ autoPush: boolean | null }>>`
      SELECT "autoPush"
      FROM "Mapping"
      WHERE "id" = ${id}
      LIMIT 1
    `;

    const { name, smartsheetSheetId, smartsheetSheetName, autoPush, connections, formulas, schemaFingerprint, changeSummary } =
      await req.json();

    const latestVersion = await prisma.mappingVersion.findFirst({
      where: { mappingId: id },
      orderBy: { versionNumber: "desc" },
    });

    const newVersionNumber = (latestVersion?.versionNumber ?? 0) + 1;

    const newVersion = await prisma.mappingVersion.create({
      data: {
        mappingId: id,
        userId,
        versionNumber: newVersionNumber,
        connections: connections || {},
        formulas: formulas || {},
        schemaFingerprint: schemaFingerprint || {},
        changeSummary: changeSummary || `Version ${newVersionNumber}`,
      },
    });

    const updated = await prisma.mapping.update({
      where: { id },
      data: {
        name: name ?? mapping.name,
        smartsheetSheetId: smartsheetSheetId ?? mapping.smartsheetSheetId,
        smartsheetSheetName: smartsheetSheetName ?? mapping.smartsheetSheetName,
        currentVersionId: newVersion.id,
      },
      include: { versions: { orderBy: { versionNumber: "asc" } } },
    });

    await prisma.$queryRaw`
      UPDATE "Mapping"
      SET "autoPush" = ${typeof autoPush === "boolean" ? autoPush : (autoPushRows[0]?.autoPush ?? false)}
      WHERE "id" = ${id}
    `;

    const refreshedAutoPush = await prisma.$queryRaw<Array<{ autoPush: boolean | null }>>`
      SELECT "autoPush"
      FROM "Mapping"
      WHERE "id" = ${id}
      LIMIT 1
    `;

    const payload = { ...updated, autoPush: refreshedAutoPush[0]?.autoPush ?? false };

    return NextResponse.json(payload);
  } catch (e) {
    console.error("Mapping update error:", e);
    return NextResponse.json({ error: (e as Error).message || "Failed to update mapping" }, { status: 500 });
  }
}

export async function DELETE(req: NextRequest, { params }: { params: Promise<{ id: string }> }) {
  const session = await getServerSession(authOptions);
  if (!session?.user) return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  const userId = (session.user as { id: string }).id;

  try {
    await assertDeleteAllowed(userId);
  } catch (e) {
    if (e instanceof SecurityPolicyError) {
      return NextResponse.json({ error: e.message }, { status: e.status });
    }
    throw e;
  }

  try {
    const { id } = await params;

    const mapping = await prisma.mapping.findFirst({ where: { id, userId } });
    if (!mapping) return NextResponse.json({ error: "Not found" }, { status: 404 });

    await prisma.mapping.delete({ where: { id } });
    return NextResponse.json({ success: true });
  } catch (e) {
    return NextResponse.json({ error: (e as Error).message }, { status: 500 });
  }
}
