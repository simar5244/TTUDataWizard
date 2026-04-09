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
  if (!userId) return NextResponse.json({ error: "User ID not found in session" }, { status: 401 });

  try {
    await assertEditAllowed(userId);
  } catch (e) {
    if (e instanceof SecurityPolicyError) {
      return NextResponse.json({ error: e.message }, { status: e.status });
    }
    throw e;
  }

  try {
    const body = await req.json();
    const { name, excelData, dataSources, charts, layout, linkedMappings } = body;
    const slug = generateSlug(10);

    const dashboard = await prisma.dashboard.create({
      data: {
        userId,
        name,
        slug,
        excelData: excelData || null,
        dataSources: dataSources || [],
        charts: charts || [],
        layout: layout || [],
        linkedMappings: linkedMappings || [],
      },
    });

    return NextResponse.json(dashboard, { status: 201 });
  } catch (e) {
    console.error("Dashboard create error:", e);
    return NextResponse.json({ error: (e as Error).message || "Failed to save dashboard" }, { status: 500 });
  }
}
