import { NextRequest, NextResponse } from "next/server";
import { getServerSession } from "next-auth";
import { authOptions } from "@/lib/auth";
import { prisma } from "@/lib/prisma";

export async function GET(req: NextRequest, { params }: { params: { id: string } }) {
  const session = await getServerSession(authOptions);
  if (!session?.user) return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  const userId = (session.user as { id: string }).id;

  const rows = await prisma.$queryRaw<Array<Record<string, unknown>>>
  `SELECT *
   FROM "MappingRun"
   WHERE "mappingId" = ${params.id} AND "userId" = ${userId}
   ORDER BY "createdAt" DESC
   LIMIT 1`;

  const run = rows[0];

  if (!run) return NextResponse.json({ error: "No runs found" }, { status: 404 });

  return NextResponse.json(run);
}
