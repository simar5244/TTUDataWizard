import { NextRequest, NextResponse } from "next/server";
import { getServerSession } from "next-auth";
import { authOptions } from "@/lib/auth";
import { prisma } from "@/lib/prisma";
import { validateToken } from "@/lib/smartsheet";
import { assertDeleteAllowed, assertEditAllowed, SecurityPolicyError } from "@/lib/security";

export async function GET() {
  const session = await getServerSession(authOptions);
  if (!session?.user) return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  const userId = (session.user as { id: string }).id;

  const user = await prisma.user.findUnique({
    where: { id: userId },
    select: { smartsheetToken: true },
  });

  const securityRows = await prisma.$queryRaw<Array<{ allowEdits: boolean | null; allowDeletes: boolean | null }>>`
    SELECT "allowEdits", "allowDeletes"
    FROM "User"
    WHERE "id" = ${userId}
    LIMIT 1
  `;
  const security = securityRows[0];

  return NextResponse.json({
    connected: !!user?.smartsheetToken,
    allowEdits: security?.allowEdits ?? true,
    allowDeletes: security?.allowDeletes ?? true,
  });
}

export async function POST(req: NextRequest) {
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

  try {
    await assertDeleteAllowed(userId);
  } catch (e) {
    if (e instanceof SecurityPolicyError) {
      return NextResponse.json({ error: e.message }, { status: e.status });
    }
    throw e;
  }

  await prisma.user.update({
    where: { id: userId },
    data: { smartsheetToken: null },
  });

  return NextResponse.json({ success: true });
}

export async function PATCH(req: NextRequest) {
  const session = await getServerSession(authOptions);
  if (!session?.user) return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  const userId = (session.user as { id: string }).id;

  const body = await req.json();
  const allowEdits = typeof body.allowEdits === "boolean" ? body.allowEdits : undefined;
  const allowDeletes = typeof body.allowDeletes === "boolean" ? body.allowDeletes : undefined;

  if (allowEdits === undefined && allowDeletes === undefined) {
    return NextResponse.json({ error: "No security settings provided" }, { status: 400 });
  }

  const prevRows = await prisma.$queryRaw<Array<{ allowEdits: boolean | null; allowDeletes: boolean | null }>>`
    SELECT "allowEdits", "allowDeletes"
    FROM "User"
    WHERE "id" = ${userId}
    LIMIT 1
  `;
  const prev = prevRows[0] || { allowEdits: true, allowDeletes: true };

  const updatedRows = await prisma.$queryRaw<Array<{ allowEdits: boolean; allowDeletes: boolean }>>`
    UPDATE "User"
    SET
      "allowEdits" = ${allowEdits ?? prev.allowEdits ?? true},
      "allowDeletes" = ${allowDeletes ?? prev.allowDeletes ?? true}
    WHERE "id" = ${userId}
    RETURNING "allowEdits", "allowDeletes"
  `;

  const updated = updatedRows[0];

  return NextResponse.json(updated);
}
