import { NextRequest, NextResponse } from "next/server";
import { getServerSession } from "next-auth";
import { authOptions } from "@/lib/auth";
import { prisma } from "@/lib/prisma";
import { assertDeleteAllowed, assertEditAllowed, SecurityPolicyError } from "@/lib/security";

export async function GET(req: NextRequest, { params }: { params: { id: string } }) {
  const session = await getServerSession(authOptions);
  if (!session?.user) return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  const userId = (session.user as { id: string }).id;

  try {
    const dashboard = await prisma.dashboard.findFirst({
      where: { id: params.id, userId },
    });
    if (!dashboard) return NextResponse.json({ error: "Not found" }, { status: 404 });
    return NextResponse.json(dashboard);
  } catch (e) {
    return NextResponse.json({ error: (e as Error).message }, { status: 500 });
  }
}

export async function PUT(req: NextRequest, { params }: { params: { id: string } }) {
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
    const dashboard = await prisma.dashboard.findFirst({ where: { id: params.id, userId } });
    if (!dashboard) return NextResponse.json({ error: "Not found" }, { status: 404 });

    const { name, excelData, dataSources, charts, layout, linkedMappings } = await req.json();

    const updated = await prisma.dashboard.update({
      where: { id: params.id },
      data: {
        name: name ?? dashboard.name,
        excelData: excelData ?? dashboard.excelData,
        dataSources: dataSources ?? dashboard.dataSources,
        charts: charts ?? dashboard.charts,
        layout: layout ?? dashboard.layout,
        linkedMappings: linkedMappings ?? dashboard.linkedMappings,
      },
    });

    return NextResponse.json(updated);
  } catch (e) {
    console.error("Dashboard update error:", e);
    return NextResponse.json({ error: (e as Error).message || "Failed to update dashboard" }, { status: 500 });
  }
}

export async function DELETE(req: NextRequest, { params }: { params: { id: string } }) {
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
    const dashboard = await prisma.dashboard.findFirst({ where: { id: params.id, userId } });
    if (!dashboard) return NextResponse.json({ error: "Not found" }, { status: 404 });
    await prisma.dashboard.delete({ where: { id: params.id } });
    return NextResponse.json({ success: true });
  } catch (e) {
    return NextResponse.json({ error: (e as Error).message }, { status: 500 });
  }
}
