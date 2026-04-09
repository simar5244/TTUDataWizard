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
    id: data.id,
    name: data.name,
    columns: (data.columns || []).map((c: Record<string, unknown>, i: number) => ({
      id: c.id,
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
  const data = await ssRequest(token, `/sheets/${sheetId}`);
  return (data.rows || []).map((r: Record<string, unknown>) => ({
    id: r.id,
    rowNumber: typeof r.rowNumber === "number" ? r.rowNumber : undefined,
    parentId: typeof r.parentId === "number" ? r.parentId : undefined,
    siblingId: typeof r.siblingId === "number" ? r.siblingId : undefined,
    expanded: typeof r.expanded === "boolean" ? r.expanded : undefined,
    locked: typeof r.locked === "boolean" ? r.locked : undefined,
    cells: ((r.cells as Record<string, unknown>[] | undefined) || []).map((c: Record<string, unknown>) => ({
      columnId: c.columnId,
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

export async function validateToken(token: string): Promise<{ valid: boolean; name?: string; email?: string }> {
  try {
    const data = await ssRequest(token, "/users/me");
    return { valid: true, name: data.name, email: data.email };
  } catch {
    return { valid: false };
  }
}
