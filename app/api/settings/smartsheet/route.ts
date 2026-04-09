import { NextRequest, NextResponse } from "next/server";
import { getServerSession } from "next-auth";
import { authOptions } from "@/lib/auth";
import { prisma } from "@/lib/prisma";
import { validateToken } from "@/lib/smartsheet";

export async function GET() {
  const session = await getServerSession(authOptions);
  if (!session?.user) return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  const userId = (session.user as { id: string }).id;

  const user = await prisma.user.findUnique({
    where: { id: userId },
    select: { smartsheetToken: true },
  });

  return NextResponse.json({ connected: !!user?.smartsheetToken });
}

export async function POST(req: NextRequest) {
  const session = await getServerSession(authOptions);
  if (!session?.user) return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  const userId = (session.user as { id: string }).id;

  const { token } = await req.json();
  if (!token) return NextResponse.json({ error: "Token required" }, { status: 400 });

  const validation = await validateToken(token);
  if (!validation.valid) {
    return NextResponse.json({ error: "Invalid Smartsheet token" }, { status: 400 });
  }

  await prisma.user.update({
    where: { id: userId },
    data: { smartsheetToken: token },
  });

  return NextResponse.json({ success: true, name: validation.name, email: validation.email });
}

export async function DELETE() {
  const session = await getServerSession(authOptions);
  if (!session?.user) return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  const userId = (session.user as { id: string }).id;

  await prisma.user.update({
    where: { id: userId },
    data: { smartsheetToken: null },
  });

  return NextResponse.json({ success: true });
}
