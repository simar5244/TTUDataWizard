import { NextRequest, NextResponse } from "next/server";
import { getServerSession } from "next-auth";
import { authOptions } from "@/lib/auth";
import { prisma } from "@/lib/prisma";

export async function GET(req: NextRequest, { params }: { params: { id: string } }) {
  const session = await getServerSession(authOptions);
  if (!session?.user) return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  const userId = (session.user as { id: string }).id;

  const dashboard = await prisma.dashboard.findFirst({
    where: { id: params.id, userId },
  });

  if (!dashboard) return NextResponse.json({ error: "Not found" }, { status: 404 });
  return NextResponse.json(dashboard);
}

export async function PUT(req: NextRequest, { params }: { params: { id: string } }) {
  const session = await getServerSession(authOptions);
  if (!session?.user) return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  const userId = (session.user as { id: string }).id;

  const dashboard = await prisma.dashboard.findFirst({ where: { id: params.id, userId } });
  if (!dashboard) return NextResponse.json({ error: "Not found" }, { status: 404 });

  const { name, excelData, charts, layout } = await req.json();

  const updated = await prisma.dashboard.update({
    where: { id: params.id },
    data: {
      name: name ?? dashboard.name,
      excelData: excelData ?? dashboard.excelData,
      charts: charts ?? dashboard.charts,
      layout: layout ?? dashboard.layout,
    },
  });

  return NextResponse.json(updated);
}

export async function DELETE(req: NextRequest, { params }: { params: { id: string } }) {
  const session = await getServerSession(authOptions);
  if (!session?.user) return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  const userId = (session.user as { id: string }).id;

  const dashboard = await prisma.dashboard.findFirst({ where: { id: params.id, userId } });
  if (!dashboard) return NextResponse.json({ error: "Not found" }, { status: 404 });

  await prisma.dashboard.delete({ where: { id: params.id } });
  return NextResponse.json({ success: true });
}
