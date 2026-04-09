import { NextRequest, NextResponse } from "next/server";
import { getServerSession } from "next-auth";
import { authOptions } from "@/lib/auth";
import { prisma } from "@/lib/prisma";
import { generateSlug } from "@/lib/utils";

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

  return NextResponse.json(mappings);
}

export async function POST(req: NextRequest) {
  const session = await getServerSession(authOptions);
  if (!session?.user) return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  const userId = (session.user as { id: string }).id;

  const { name, smartsheetSheetId, smartsheetSheetName, connections, formulas, schemaFingerprint, changeSummary } =
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
}
