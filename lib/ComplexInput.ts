import * as XLSX from "xlsx";

export interface ExcelColumn {
  key: string;
  header: string;
  index: number;
  dataType: "string" | "number" | "date" | "boolean" | "empty";
  sampleValues: (string | number | boolean | null)[];
}

export interface ExcelSheet {
  name: string;
  columns: ExcelColumn[];
  rowCount: number;
  data: Record<string, string | number | boolean | null>[];
}

export interface ParsedExcel {
  sheets: ExcelSheet[];
  fileName: string;
  fileSize: number;
}

function detectDataType(values: (string | number | boolean | null)[]): ExcelColumn["dataType"] {
  const nonNull = values.filter((v) => v !== null && v !== undefined && v !== "");
  if (nonNull.length === 0) return "empty";
  const numericCount = nonNull.filter((v) => typeof v === "number" || !Number.isNaN(Number(v))).length;
  if (numericCount / nonNull.length > 0.8) return "number";
  const dateCount = nonNull.filter((v) => {
    const s = String(v);
    return /^\d{1,4}[-/]\d{1,2}[-/]\d{1,4}/.test(s) || !Number.isNaN(Date.parse(s));
  }).length;
  if (dateCount / nonNull.length > 0.7) return "date";
  return "string";
}

function normalizeCell(value: unknown): string | number | boolean | null {
  if (value === undefined || value === null) return null;
  if (typeof value === "number" || typeof value === "boolean") return value;
  const asString = String(value).trim();
  if (asString === "") return null;
  return asString;
}

function hasLetters(value: unknown): boolean {
  return /[A-Za-z]/.test(String(value ?? ""));
}

function looksDateLike(value: unknown): boolean {
  const text = String(value ?? "").trim();
  if (!text) return false;
  if (/^\d{4}-\d{1,2}-\d{1,2}/.test(text)) return true;
  if (/^\d{1,2}[/-]\d{1,2}[/-]\d{2,4}/.test(text)) return true;
  const parsed = Date.parse(text);
  return !Number.isNaN(parsed);
}

function scoreHeaderRow(rows: unknown[][], rowIndex: number): { score: number; nonEmptyIndexes: number[]; textCells: number } {
  const row = rows[rowIndex] ?? [];
  const nonEmptyIndexes: number[] = [];
  let textCells = 0;
  let numericLikeCells = 0;

  for (let i = 0; i < row.length; i += 1) {
    const value = normalizeCell(row[i]);
    if (value === null) continue;
    nonEmptyIndexes.push(i);
    const s = String(value);
    const numericLike = !Number.isNaN(Number(s));
    if (hasLetters(s) && !numericLike && !looksDateLike(s)) textCells += 1;
    if (numericLike || looksDateLike(s)) numericLikeCells += 1;
  }

  if (nonEmptyIndexes.length === 0) {
    return { score: Number.NEGATIVE_INFINITY, nonEmptyIndexes: [], textCells: 0 };
  }

  let continuation = 0;
  const lookahead = Math.min(rows.length, rowIndex + 8);
  for (let r = rowIndex + 1; r < lookahead; r += 1) {
    const next = rows[r] ?? [];
    for (const colIdx of nonEmptyIndexes) {
      if (normalizeCell(next[colIdx]) !== null) continuation += 1;
    }
  }

  const uniqueHeaders = new Set(
    nonEmptyIndexes
      .map((idx) => String(normalizeCell(row[idx]) ?? "").toLowerCase())
      .filter((v) => v.length > 0)
  ).size;

  const score =
    textCells * 3 +
    uniqueHeaders * 1.2 +
    continuation * 0.35 +
    nonEmptyIndexes.length * 0.5 -
    numericLikeCells * 1.75;

  return { score, nonEmptyIndexes, textCells };
}

function indexToColRef(index: number): string {
  let result = "";
  let n = index;
  while (n >= 0) {
    result = String.fromCharCode(65 + (n % 26)) + result;
    n = Math.floor(n / 26) - 1;
  }
  return result;
}

function makeUniqueHeaders(headers: string[]): string[] {
  const counts = new Map<string, number>();
  return headers.map((header) => {
    const base = header.trim() || "Column";
    const seen = counts.get(base) ?? 0;
    counts.set(base, seen + 1);
    return seen === 0 ? base : `${base}_${seen + 1}`;
  });
}

function parseComplexSheet(worksheet: XLSX.WorkSheet, sheetName: string): ExcelSheet | null {
  const grid = XLSX.utils.sheet_to_json<unknown[]>(worksheet, {
    header: 1,
    raw: false,
    defval: null,
    blankrows: false,
  });

  if (!grid.length) return null;

  const probeLimit = Math.min(grid.length, 60);
  let bestRowIndex = 0;
  let best = { score: Number.NEGATIVE_INFINITY, nonEmptyIndexes: [] as number[], textCells: 0 };

  for (let i = 0; i < probeLimit; i += 1) {
    const candidate = scoreHeaderRow(grid, i);
    if (candidate.score > best.score && candidate.nonEmptyIndexes.length >= 2) {
      best = candidate;
      bestRowIndex = i;
    }
  }

  const headerRow = grid[bestRowIndex] ?? [];
  const explicitHeader = best.nonEmptyIndexes.length > 0 && best.textCells / best.nonEmptyIndexes.length >= 0.35;

  const activeIndexes = best.nonEmptyIndexes.length > 0
    ? best.nonEmptyIndexes
    : Array.from(
        new Set(
          grid
            .slice(0, Math.min(grid.length, 40))
            .flatMap((r) => r.map((v, idx) => (normalizeCell(v) === null ? -1 : idx)).filter((idx) => idx >= 0))
        )
      );

  if (!activeIndexes.length) return null;

  const rawHeaders = activeIndexes.map((colIdx) => {
    const headerCell = normalizeCell(headerRow[colIdx]);
    if (explicitHeader && headerCell !== null) return String(headerCell);
    return `Column ${indexToColRef(colIdx)}`;
  });
  const headers = makeUniqueHeaders(rawHeaders);

  const dataStart = explicitHeader ? bestRowIndex + 1 : bestRowIndex;
  const data: Record<string, string | number | boolean | null>[] = [];

  for (let r = dataStart; r < grid.length; r += 1) {
    const row = grid[r] ?? [];
    const out: Record<string, string | number | boolean | null> = {};
    for (let i = 0; i < activeIndexes.length; i += 1) {
      const colIdx = activeIndexes[i];
      const key = headers[i];
      out[key] = normalizeCell(row[colIdx]);
    }
    data.push(out);
  }

  const columns: ExcelColumn[] = headers.map((header, idx) => {
    const sampleValues = data.slice(0, 10).map((row) => row[header] ?? null);
    return {
      key: `col_${activeIndexes[idx]}`,
      header,
      index: activeIndexes[idx],
      dataType: detectDataType(sampleValues),
      sampleValues,
    };
  });

  return {
    name: sheetName,
    columns,
    rowCount: data.length,
    data,
  };
}

export function parseExcelFileComplex(buffer: ArrayBuffer, fileName: string, fileSize: number): ParsedExcel {
  const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
  const sheets: ExcelSheet[] = [];

  for (const sheetName of workbook.SheetNames) {
    const worksheet = workbook.Sheets[sheetName];
    const parsed = parseComplexSheet(worksheet, sheetName);
    if (parsed && parsed.data.length > 0) {
      sheets.push(parsed);
    }
  }

  return { sheets, fileName, fileSize };
}
