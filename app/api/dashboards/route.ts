import { NextRequest, NextResponse } from "next/server";
import { getServerSession } from "next-auth";
import { authOptions } from "@/lib/auth";
import { prisma } from "@/lib/prisma";
import { generateSlug } from "@/lib/utils";

export async function GET() {
  const session = await getServerSession(authOptions);
  if (!session?.user) return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  const userId = (session.user as { id: string }).id;

  const dashboards = await prisma.dashboard.findMany({
    where: { userId },
    orderBy: { updatedAt: "desc" },
  });

  return NextResponse.json(dashboards);
}

export async function POST(req: NextRequest) {
  const session = await getServerSession(authOptions);
  if (!session?.user) return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  const userId = (session.user as { id: string }).id;

  const { name, excelData, charts, layout } = await req.json();
  const slug = generateSlug(10);

  const dashboard = await prisma.dashboard.create({
    data: {
      userId,
      name,
      slug,
      excelData: excelData || null,
      charts: charts || [],
      layout: layout || [],
    },
  });

  return NextResponse.json(dashboard, { status: 201 });
}
