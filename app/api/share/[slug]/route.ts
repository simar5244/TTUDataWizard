import { NextRequest, NextResponse } from "next/server";
import { getServerSession } from "next-auth";
import { authOptions } from "@/lib/auth";
import { prisma } from "@/lib/prisma";

export async function GET(req: NextRequest, { params }: { params: { slug: string } }) {
  const { searchParams } = new URL(req.url);
  const type = searchParams.get("type");

  if (type === "dashboard") {
    const dashboard = await prisma.dashboard.findFirst({
      where: { slug: params.slug },
      select: { id: true, name: true, slug: true, userId: true, excelData: true, charts: true, layout: true },
    });
    if (!dashboard) return NextResponse.json({ error: "Dashboard not found" }, { status: 404 });
    return NextResponse.json({ type: "dashboard", data: dashboard });
  }

  if (type === "mapping") {
    const session = await getServerSession(authOptions);
    if (!session?.user) {
      return NextResponse.json({ error: "Sign in to view shared mappings" }, { status: 401 });
    }
    const mapping = await prisma.mapping.findFirst({
      where: { slug: params.slug },
      include: {
        versions: { orderBy: { versionNumber: "asc" } },
        user: { select: { name: true } },
      },
    });
    if (!mapping) return NextResponse.json({ error: "Mapping not found" }, { status: 404 });
    return NextResponse.json({ type: "mapping", data: mapping });
  }

  return NextResponse.json({ error: "Invalid type" }, { status: 400 });
}
