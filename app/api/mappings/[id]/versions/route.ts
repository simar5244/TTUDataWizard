import { NextRequest, NextResponse } from "next/server";
import { getServerSession } from "next-auth";
import { authOptions } from "@/lib/auth";
import { prisma } from "@/lib/prisma";

export async function GET(req: NextRequest, { params }: { params: { id: string } }) {
  const session = await getServerSession(authOptions);
  if (!session?.user) return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  const userId = (session.user as { id: string }).id;

  const mapping = await prisma.mapping.findFirst({ where: { id: params.id, userId } });
  if (!mapping) return NextResponse.json({ error: "Not found" }, { status: 404 });

  const versions = await prisma.mappingVersion.findMany({
    where: { mappingId: params.id },
    orderBy: { versionNumber: "desc" },
    include: { user: { select: { name: true, email: true } } },
  });

  return NextResponse.json(versions);
}

export async function POST(req: NextRequest, { params }: { params: { id: string } }) {
  const session = await getServerSession(authOptions);
  if (!session?.user) return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  const userId = (session.user as { id: string }).id;

  const mapping = await prisma.mapping.findFirst({ where: { id: params.id, userId } });
  if (!mapping) return NextResponse.json({ error: "Not found" }, { status: 404 });

  const { versionId } = await req.json();

  const version = await prisma.mappingVersion.findFirst({
    where: { id: versionId, mappingId: params.id },
  });
  if (!version) return NextResponse.json({ error: "Version not found" }, { status: 404 });

  await prisma.mapping.update({
    where: { id: params.id },
    data: { currentVersionId: version.id },
  });

  return NextResponse.json({ success: true, currentVersionId: version.id });
}
