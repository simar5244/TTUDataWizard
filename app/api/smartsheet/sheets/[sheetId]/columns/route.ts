import { NextRequest, NextResponse } from "next/server";
import { getServerSession } from "next-auth";
import { authOptions } from "@/lib/auth";
import { prisma } from "@/lib/prisma";
import { addSheetColumn, deleteSheetColumn, updateSheetColumn } from "@/lib/smartsheet";

export async function POST(req: NextRequest, { params }: { params: { sheetId: string } }) {
  const session = await getServerSession(authOptions);
  if (!session?.user) return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  const userId = (session.user as { id: string }).id;

  const user = await prisma.user.findUnique({ where: { id: userId } });
  if (!user?.smartsheetToken) {
    return NextResponse.json({ error: "Smartsheet not connected" }, { status: 400 });
  }
  const payload = await req.json().catch(() => ({}));
  const title = String(payload?.title ?? "").trim();
  const type = String(payload?.type ?? "TEXT_NUMBER").trim() || "TEXT_NUMBER";
  const index =
    typeof payload?.index === "number" && Number.isFinite(payload.index) && payload.index >= 0
      ? Math.floor(payload.index)
      : undefined;

  if (!title) {
    return NextResponse.json({ error: "Column title is required" }, { status: 400 });
  }

  try {
    const created = await addSheetColumn(user.smartsheetToken, params.sheetId, { title, type, index });
    return NextResponse.json(created, { status: 201 });
  } catch (e) {
    return NextResponse.json({ error: (e as Error).message }, { status: 500 });
  }
}

export async function DELETE(req: NextRequest, { params }: { params: { sheetId: string } }) {
  const session = await getServerSession(authOptions);
  if (!session?.user) return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  const userId = (session.user as { id: string }).id;

  const user = await prisma.user.findUnique({ where: { id: userId } });
  if (!user?.smartsheetToken) {
    return NextResponse.json({ error: "Smartsheet not connected" }, { status: 400 });
  }

  const columnIdRaw = req.nextUrl.searchParams.get("columnId");
  const columnId = Number(columnIdRaw);
  if (!Number.isFinite(columnId) || columnId <= 0) {
    return NextResponse.json({ error: "Valid columnId is required" }, { status: 400 });
  }

  try {
    await deleteSheetColumn(user.smartsheetToken, params.sheetId, Math.floor(columnId));
    return NextResponse.json({ ok: true });
  } catch (e) {
    return NextResponse.json({ error: (e as Error).message }, { status: 500 });
  }
}

export async function PUT(req: NextRequest, { params }: { params: { sheetId: string } }) {
  const session = await getServerSession(authOptions);
  if (!session?.user) return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  const userId = (session.user as { id: string }).id;

  const user = await prisma.user.findUnique({ where: { id: userId } });
  if (!user?.smartsheetToken) {
    return NextResponse.json({ error: "Smartsheet not connected" }, { status: 400 });
  }

  const payload = await req.json().catch(() => ({}));
  const columnId = Number(payload?.columnId);
  if (!Number.isFinite(columnId) || columnId <= 0) {
    return NextResponse.json({ error: "Valid columnId is required" }, { status: 400 });
  }

  const title = typeof payload?.title === "string" ? payload.title.trim() : undefined;
  const type = typeof payload?.type === "string" ? payload.type.trim() : undefined;
  const index =
    typeof payload?.index === "number" && Number.isFinite(payload.index) && payload.index >= 0
      ? Math.floor(payload.index)
      : undefined;

  if (title !== undefined && title === "") {
    return NextResponse.json({ error: "Column title cannot be empty" }, { status: 400 });
  }

  if (title === undefined && type === undefined && index === undefined) {
    return NextResponse.json({ error: "No column updates provided" }, { status: 400 });
  }

  try {
    const updated = await updateSheetColumn(user.smartsheetToken, params.sheetId, Math.floor(columnId), {
      title,
      type,
      index,
    });
    return NextResponse.json(updated);
  } catch (e) {
    return NextResponse.json({ error: (e as Error).message }, { status: 500 });
  }
}
