export interface SmartsheetColumn {
  id: number;
  title: string;
  type: string;
  index: number;
  formula?: string;
}

export interface SmartsheetSheet {
  id: number;
  name: string;
  columns: SmartsheetColumn[];
  rowCount: number;
  permalink: string;
}

export interface SmartsheetRow {
  id: number;
  rowNumber?: number;
  parentId?: number;
  siblingId?: number;
  expanded?: boolean;
  locked?: boolean;
  cells: {
    columnId: number;
    value: string | number | boolean | null;
    displayValue?: string;
    formula?: string;
    locked?: boolean;
  }[];
}

const SS_BASE = "https://api.smartsheet.com/2.0";

async function ssRequest(token: string, path: string, options: RequestInit = {}) {
  const res = await fetch(`${SS_BASE}${path}`, {
    ...options,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      ...(options.headers || {}),
    },
  });
  if (!res.ok) {
    const err = await res.json().catch(() => ({ message: res.statusText }));
    throw new Error(err.message || `Smartsheet API error ${res.status}`);
  }
  return res.json();
}

export async function listSheets(token: string): Promise<SmartsheetSheet[]> {
  const data = await ssRequest(token, "/sheets?includeAll=true");
  return (data.data || []).map((s: Record<string, unknown>) => ({
    id: s.id,
    name: s.name,
    columns: [],
    rowCount: s.totalRowCount || 0,
    permalink: s.permalink || "",
  }));
}

export async function getSheet(token: string, sheetId: string): Promise<SmartsheetSheet> {
  const data = await ssRequest(token, `/sheets/${sheetId}`);
  return {
    id: Number(data.id),
    name: data.name,
    columns: (data.columns || []).map((c: Record<string, unknown>, i: number) => ({
      id: Number(c.id),
      title: c.title,
      type: c.type,
      index: i,
      formula: typeof c.formula === "string" ? c.formula : undefined,
    })),
    rowCount: (data.rows || []).length,
    permalink: data.permalink || "",
  };
}

export async function getSheetRows(token: string, sheetId: string): Promise<SmartsheetRow[]> {
  const data = await ssRequest(token, `/sheets/${sheetId}?include=formulas`);
  return (data.rows || []).map((r: Record<string, unknown>) => ({
    id: Number(r.id),
    rowNumber: Number.isFinite(Number(r.rowNumber)) ? Number(r.rowNumber) : undefined,
    parentId: Number.isFinite(Number(r.parentId)) ? Number(r.parentId) : undefined,
    siblingId: Number.isFinite(Number(r.siblingId)) ? Number(r.siblingId) : undefined,
    expanded: typeof r.expanded === "boolean" ? r.expanded : undefined,
    locked: typeof r.locked === "boolean" ? r.locked : undefined,
    cells: ((r.cells as Record<string, unknown>[] | undefined) || []).map((c: Record<string, unknown>) => ({
      columnId: Number(c.columnId),
      value: c.value ?? null,
      displayValue: c.displayValue,
      formula: typeof c.formula === "string" ? c.formula : undefined,
      locked: typeof c.locked === "boolean" ? c.locked : undefined,
    })),
  }));
}

export async function duplicateSheet(
  token: string,
  sheetId: string,
  newName: string
): Promise<{ id: number; name: string; permalink: string }> {
  const data = await ssRequest(token, `/sheets/${sheetId}/copy`, {
    method: "POST",
    body: JSON.stringify({
      newName,
      destinationType: "home",
    }),
  });
  return {
    id: data.result?.id,
    name: data.result?.name,
    permalink: data.result?.permalink || "",
  };
}

export async function updateSheetRows(
  token: string,
  sheetId: string,
  rows: { id: number; cells: { columnId: number; value: string | number | boolean | null }[] }[]
): Promise<void> {
  await ssRequest(token, `/sheets/${sheetId}/rows`, {
    method: "PUT",
    body: JSON.stringify(rows),
  });
}

export async function addSheetRows(
  token: string,
  sheetId: string,
  rows: { toBottom: boolean; cells: { columnId: number; value: string | number | boolean | null }[] }[]
): Promise<void> {
  await ssRequest(token, `/sheets/${sheetId}/rows`, {
    method: "POST",
    body: JSON.stringify(rows),
  });
}

export async function addSheetColumn(
  token: string,
  sheetId: string,
  column: { title: string; type?: string; index?: number }
): Promise<SmartsheetColumn> {
  const title = String(column.title ?? "").trim();
  if (!title) {
    throw new Error("Column title is required");
  }

  const payload: Record<string, unknown> = {
    title,
    type: column.type || "TEXT_NUMBER",
  };
  if (typeof column.index === "number" && Number.isFinite(column.index) && column.index >= 0) {
    payload.index = Math.floor(column.index);
  }

  const data = await ssRequest(token, `/sheets/${sheetId}/columns`, {
    method: "POST",
    body: JSON.stringify([payload]),
  });

  const created =
    (Array.isArray(data?.result) ? data.result[0] : null) ||
    (Array.isArray(data?.data) ? data.data[0] : null) ||
    data?.result ||
    data;

  return {
    id: Number(created?.id),
    title: String(created?.title ?? column.title),
    type: String(created?.type ?? column.type ?? "TEXT_NUMBER"),
    index: Number.isFinite(Number(created?.index)) ? Number(created.index) : (column.index ?? 0),
    formula: typeof created?.formula === "string" ? created.formula : undefined,
  };
}

export async function deleteSheetColumn(token: string, sheetId: string, columnId: number): Promise<void> {
  await ssRequest(token, `/sheets/${sheetId}/columns/${columnId}`, {
    method: "DELETE",
  });
}

export async function updateSheetColumn(
  token: string,
  sheetId: string,
  columnId: number,
  updates: { title?: string; type?: string; index?: number }
): Promise<SmartsheetColumn> {
  const payload: Record<string, unknown> = {};

  if (typeof updates.title === "string") {
    const title = updates.title.trim();
    if (!title) {
      throw new Error("Column title is required");
    }
    payload.title = title;
  }

  if (typeof updates.type === "string" && updates.type.trim() !== "") {
    payload.type = updates.type.trim();
  }

  if (typeof updates.index === "number" && Number.isFinite(updates.index) && updates.index >= 0) {
    payload.index = Math.floor(updates.index);
  }

  if (Object.keys(payload).length === 0) {
    throw new Error("No column updates provided");
  }

  const data = await ssRequest(token, `/sheets/${sheetId}/columns/${columnId}`, {
    method: "PUT",
    body: JSON.stringify(payload),
  });

  const updated = data?.result ?? data;
  return {
    id: Number(updated?.id ?? columnId),
    title: String(updated?.title ?? updates.title ?? ""),
    type: String(updated?.type ?? updates.type ?? "TEXT_NUMBER"),
    index: Number.isFinite(Number(updated?.index)) ? Number(updated.index) : 0,
    formula: typeof updated?.formula === "string" ? updated.formula : undefined,
  };
}

export async function validateToken(token: string): Promise<{ valid: boolean; name?: string; email?: string }> {
  try {
    const data = await ssRequest(token, "/users/me");
    return { valid: true, name: data.name, email: data.email };
  } catch {
    return { valid: false };
  }
}
