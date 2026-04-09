import { prisma } from "@/lib/prisma";

export class SecurityPolicyError extends Error {
  status: number;

  constructor(message: string, status = 403) {
    super(message);
    this.name = "SecurityPolicyError";
    this.status = status;
  }
}

export async function getUserSecurityPolicy(userId: string): Promise<{ allowEdits: boolean; allowDeletes: boolean }> {
  const rows = await prisma.$queryRaw<Array<{ allowEdits: boolean | null; allowDeletes: boolean | null }>>`
    SELECT "allowEdits", "allowDeletes"
    FROM "User"
    WHERE "id" = ${userId}
    LIMIT 1
  `;

  const row = rows[0];
  return {
    allowEdits: row?.allowEdits ?? true,
    allowDeletes: row?.allowDeletes ?? true,
  };
}

export async function assertEditAllowed(userId: string): Promise<void> {
  const policy = await getUserSecurityPolicy(userId);
  if (!policy.allowEdits) {
    throw new SecurityPolicyError("Edits are disabled in Settings. Turn on Allow edits to continue.");
  }
}

export async function assertDeleteAllowed(userId: string): Promise<void> {
  const policy = await getUserSecurityPolicy(userId);
  if (!policy.allowDeletes) {
    throw new SecurityPolicyError("Deletes are disabled in Settings. Turn on Allow deletes to continue.");
  }
}
