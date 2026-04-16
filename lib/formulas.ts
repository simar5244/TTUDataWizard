import { evaluate } from "mathjs";

export interface FormulaContext {
  [columnKey: string]: number | string | null;
}

function isBlank(value: unknown): boolean {
  return value === null || value === undefined || String(value).trim() === "";
}

function toNumber(value: unknown): number {
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  if (typeof value === "boolean") return value ? 1 : 0;
  const text = String(value ?? "").trim();
  const normalized = text.replace(/[$,%\s]/g, "").replace(/,/g, "");
  const n = Number(normalized);
  return Number.isFinite(n) ? n : 0;
}

function excelSerialToDate(serial: number): Date | null {
  if (!Number.isFinite(serial)) return null;
  // Excel serial date: days since 1899-12-30
  const epoch = Date.UTC(1899, 11, 30);
  const millis = epoch + Math.floor(serial) * 24 * 60 * 60 * 1000;
  const date = new Date(millis);
  return Number.isNaN(date.getTime()) ? null : date;
}

function parseDateValue(value: unknown): Date | null {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
  if (isBlank(value)) return null;
  if (typeof value === "number") {
    if (!Number.isFinite(value)) return null;
    if (value > 59 && value < 60000) {
      const excelDate = excelSerialToDate(value);
      if (excelDate) return excelDate;
    }
    const numericDate = new Date(value);
    return Number.isNaN(numericDate.getTime()) ? null : numericDate;
  }
  const text = String(value).trim();
  if (/^\d{5}(\.\d+)?$/.test(text)) {
    const excelDate = excelSerialToDate(Number(text));
    if (excelDate) return excelDate;
  }
  if (/^\d{8}$/.test(text)) {
    const year = Number(text.slice(0, 4));
    const month = Number(text.slice(4, 6));
    const day = Number(text.slice(6, 8));
    const compact = new Date(year, month - 1, day);
    if (!Number.isNaN(compact.getTime())) return compact;
  }
  if (!/[\-/:T\s]/.test(text)) return null;
  if (/^\d{4}-\d{1,2}-\d{1,2}/.test(text) || /^\d{1,2}[/-]\d{1,2}[/-]\d{2,4}/.test(text)) {
    const parsed = new Date(text);
    if (!Number.isNaN(parsed.getTime())) return parsed;
  }
  const fallback = new Date(text);
  return Number.isNaN(fallback.getTime()) ? null : fallback;
}

function isNumericLike(value: unknown): boolean {
  if (isBlank(value)) return false;
  if (typeof value === "number") return Number.isFinite(value);
  if (typeof value === "boolean") return false;
  const text = String(value).trim();
  if (text === "") return false;
  const normalized = text.replace(/[$,%\s]/g, "").replace(/,/g, "");
  return normalized !== "" && Number.isFinite(Number(normalized));
}

function isDateLike(value: unknown): boolean {
  return parseDateValue(value) !== null;
}

function isStringLike(value: unknown): boolean {
  if (isBlank(value)) return false;
  if (typeof value !== "string") return false;
  const trimmed = value.trim();
  if (trimmed === "") return false;
  return !isNumericLike(trimmed) && !isDateLike(trimmed);
}

// String manipulation functions (Excel-like)
function trim(value: string): string {
  return String(value).trim();
}

function upper(value: string): string {
  return String(value).toUpperCase();
}

function lower(value: string): string {
  return String(value).toLowerCase();
}

function left(value: string, count: number): string {
  return String(value).slice(0, count);
}

function right(value: string, count: number): string {
  const str = String(value);
  return str.slice(Math.max(0, str.length - count));
}

function mid(value: string, start: number, length: number): string {
  return String(value).slice(start - 1, start - 1 + length);
}

function len(value: string): number {
  return String(value).length;
}

function replace(value: string, oldText: string, newText: string): string {
  return String(value).split(oldText).join(newText);
}

function remove(value: string, textToRemove: string): string {
  return String(value).split(textToRemove).join('');
}

function clean(value: string): string {
  // Remove non-printable characters (0-31 and 127)
  return String(value).replace(/[\x00-\x1F\x7F]/g, '');
}

function concat(...values: (string | number)[]): string {
  return values.map(v => String(v)).join('');
}

function andFn(...values: unknown[]): boolean {
  return values.every(Boolean);
}

function orFn(...values: unknown[]): boolean {
  return values.some(Boolean);
}

function notFn(value: unknown): boolean {
  return !value;
}

function ifFn(condition: unknown, trueValue: unknown, falseValue: unknown): unknown {
  return condition ? trueValue : falseValue;
}

function ifnumeric(value: unknown, trueValue: unknown, falseValue: unknown): unknown {
  return isNumericLike(value) ? trueValue : falseValue;
}

function ifdate(value: unknown, trueValue: unknown, falseValue: unknown): unknown {
  return isDateLike(value) ? trueValue : falseValue;
}

function ifstring(value: unknown, trueValue: unknown, falseValue: unknown): unknown {
  return isStringLike(value) ? trueValue : falseValue;
}

function sum(...values: unknown[]): number {
  return values.reduce<number>((acc, value) => acc + toNumber(value), 0);
}

export function normalizeFormulaExpression(formula: string): string {
  return formula
    .trim()
    .replace(/^=/, "")
    .replace(/""([^\"]*)"/g, '"$1"')
    .replace(/\s*&\s*/g, " + ")
    .replace(/\bIF\s*\(/gi, "if(")
    .replace(/\bIFNUMERIC\s*\(/gi, "ifnumeric(")
    .replace(/\bIFDATE\s*\(/gi, "ifdate(")
    .replace(/\bIFSTRING\s*\(/gi, "ifstring(")
    .replace(/\bAND\s*\(/gi, "and(")
    .replace(/\bOR\s*\(/gi, "or(")
    .replace(/\bNOT\s*\(/gi, "not(")
    .replace(/\bSUM\s*\(/gi, "sum(");
}

export function createFormulaScope(context: FormulaContext): Record<string, unknown> {
  return {
    ...context,
    trim,
    upper,
    lower,
    left,
    right,
    mid,
    len,
    replace,
    remove,
    clean,
    concat,
    abs,
    min,
    max,
    round,
    if: ifFn,
    IF: ifFn,
    and: andFn,
    AND: andFn,
    or: orFn,
    OR: orFn,
    not: notFn,
    NOT: notFn,
    ifnumeric,
    IFNUMERIC: ifnumeric,
    ifdate,
    IFDATE: ifdate,
    ifstring,
    IFSTRING: ifstring,
    sum,
    SUM: sum,
    isnumber: isNumericLike,
    ISNUMBER: isNumericLike,
    isdate: isDateLike,
    ISDATE: isDateLike,
    istext: isStringLike,
    ISTEXT: isStringLike,
  };
}

// Math functions
function abs(value: number): number {
  return Math.abs(value);
}

function min(...values: number[]): number {
  return Math.min(...values);
}

function max(...values: number[]): number {
  return Math.max(...values);
}

function round(value: number, decimals: number = 0): number {
  const factor = Math.pow(10, decimals);
  return Math.round(value * factor) / factor;
}

export function evaluateFormula(formula: string, context: FormulaContext): number | string | null {
  try {
    const scope = createFormulaScope(context);
    const expression = normalizeFormulaExpression(formula);
    const result = evaluate(expression, scope);

    if (typeof result === "number") {
      return isNaN(result) || !isFinite(result) ? null : Math.round(result * 10000) / 10000;
    }
    if (result === null || result === undefined) return null;
    if (typeof result === "boolean") return result ? "true" : "false";
    return String(result);
  } catch (e) {
    console.error("Formula evaluation error:", e);
    return null;
  }
}

export function applyFormulaToRows(
  formula: string,
  columnKeys: string[],
  data: Record<string, string | number | boolean | null>[]
): (number | string | null)[] {
  return data.map((row) => {
    const context: FormulaContext = {};
    for (const key of columnKeys) {
      const val = row[key];
      context[key] = val !== undefined && val !== null ? (typeof val === "boolean" ? (val ? 1 : 0) : val as string | number) : null;
    }
    return evaluateFormula(formula, context);
  });
}

export function validateFormula(formula: string, sampleContext: FormulaContext): { valid: boolean; error?: string } {
  try {
    const scope = createFormulaScope(sampleContext);
    const expression = normalizeFormulaExpression(formula);
    
    for (const [key, value] of Object.entries(sampleContext)) {
      if (value === null || value === undefined) {
        scope[key] = 0;
      }
    }
    
    evaluate(expression, scope);
    return { valid: true };
  } catch (e) {
    return { valid: false, error: String(e) };
  }
}

// Detect column references in formula (supports A, B, C... AA, AB style)
export function extractColumnRefs(formula: string): string[] {
  // Match single to four letters (Excel-style and LA/LAA prefixed refs)
  const matches = formula.match(/\b[A-Z]{1,4}\b/g) || [];
  const reserved = new Set([
    "IF", "SUM", "MIN", "MAX", "ABS", "AND", "OR", "NOT",
    "ROUND", "TRIM", "LEFT", "RIGHT", "MID", "LEN", "REPLACE",
    "REMOVE", "CLEAN", "CONCAT", "IFNUMERIC", "IFDATE", "IFSTRING",
    "ISNUMBER", "ISDATE", "ISTEXT", "TRUE", "FALSE",
  ]);
  return Array.from(new Set(matches.filter((m) => !reserved.has(m))));
}

// Convert column index to Excel-style reference (0 -> A, 1 -> B, 25 -> Z, 26 -> AA)
export function indexToColRef(index: number): string {
  let result = '';
  let n = index;
  while (n >= 0) {
    result = String.fromCharCode(65 + (n % 26)) + result;
    n = Math.floor(n / 26) - 1;
  }
  return result;
}

// Convert Excel-style reference to column index (A -> 0, B -> 1, Z -> 25, AA -> 26)
export function colRefToIndex(ref: string): number {
  let result = 0;
  for (let i = 0; i < ref.length; i++) {
    result = result * 26 + (ref.charCodeAt(i) - 64);
  }
  return result - 1;
}
