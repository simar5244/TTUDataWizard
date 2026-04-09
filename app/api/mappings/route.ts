import { NextRequest, NextResponse } from "next/server";
import { getServerSession } from "next-auth";
import { authOptions } from "@/lib/auth";
import { prisma } from "@/lib/prisma";
import { generateSlug } from "@/lib/utils";
import { assertEditAllowed, SecurityPolicyError } from "@/lib/security";

export async function GET() {
  const session = await getServerSession(authOptions);
  if (!session?.user) return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  const userId = (session.user as { id: string }).id;

  const mappings = await prisma.mapping.findMany({
    where: { userId },
    include: {
      versions: { orderBy: { versionNumber: "desc" }, take: 1 },
      stagingRuns: { orderBy: { createdAt: "desc" }, take: 1 },
    },
    orderBy: { updatedAt: "desc" },
  });

  const autoPushRows = await prisma.$queryRaw<Array<{ id: string; autoPush: boolean | null }>>`
    SELECT "id", "autoPush"
    FROM "Mapping"
    WHERE "userId" = ${userId}
  `;
  const autoPushById = new Map(autoPushRows.map((r) => [r.id, r.autoPush ?? false]));

  const enriched = await Promise.all(
    mappings.map(async (m) => {
      try {
        const rows = await prisma.$queryRaw<Array<{ count: bigint }>>`
          SELECT COUNT(*)::bigint AS "count"
          FROM "MappingRun"
          WHERE "mappingId" = ${m.id} AND "userId" = ${userId}
        `;
        return {
          ...m,
          autoPush: autoPushById.get(m.id) ?? false,
          mappingRunCount: Number(rows[0]?.count ?? 0),
        };
      } catch {
        return { ...m, autoPush: autoPushById.get(m.id) ?? false, mappingRunCount: 0 };
      }
    })
  );

  return NextResponse.json(enriched);
}

export async function POST(req: NextRequest) {
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
    const { name, smartsheetSheetId, smartsheetSheetName, autoPush, connections, formulas, schemaFingerprint, changeSummary } =
      await req.json();

    const slug = generateSlug(10);

    const mapping = await prisma.mapping.create({
      data: {
        userId,
        name,
        slug,
        smartsheetSheetId,
        smartsheetSheetName,
      },
    });

    await prisma.$queryRaw`
      UPDATE "Mapping"
      SET "autoPush" = ${!!autoPush}
      WHERE "id" = ${mapping.id}
    `;

    const version = await prisma.mappingVersion.create({
      data: {
        mappingId: mapping.id,
        userId,
        versionNumber: 1,
        connections: connections || {},
        formulas: formulas || {},
        schemaFingerprint: schemaFingerprint || {},
        changeSummary: changeSummary || "Initial version",
      },
    });

    const updated = await prisma.mapping.update({
      where: { id: mapping.id },
      data: { currentVersionId: version.id },
    });

    return NextResponse.json({ ...updated, currentVersion: version }, { status: 201 });
  } catch (e) {
    console.error("Mapping create error:", e);
    return NextResponse.json({ error: (e as Error).message || "Failed to save mapping" }, { status: 500 });
  }
}
