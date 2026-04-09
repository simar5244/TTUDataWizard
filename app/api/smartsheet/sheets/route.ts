import { NextResponse } from "next/server";
import { getServerSession } from "next-auth";
import { authOptions } from "@/lib/auth";
import { prisma } from "@/lib/prisma";
import { listSheets } from "@/lib/smartsheet";

export async function GET() {
  const session = await getServerSession(authOptions);
  if (!session?.user) return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  const userId = (session.user as { id: string }).id;

  const user = await prisma.user.findUnique({ where: { id: userId } });
  if (!user?.smartsheetToken) {
    return NextResponse.json({ error: "Smartsheet not connected" }, { status: 400 });
  }

  try {
    const sheets = await listSheets(user.smartsheetToken);
    return NextResponse.json(sheets);
  } catch (e) {
    return NextResponse.json({ error: (e as Error).message }, { status: 500 });
  }
}
