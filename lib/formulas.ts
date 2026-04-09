import { evaluate } from "mathjs";

export interface FormulaContext {
  [columnKey: string]: number | string | null;
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
    // Build scope with all available functions
    const scope: Record<string, unknown> = {
      ...context,
      trim, upper, lower, left, right, mid, len, replace, remove, clean, concat,
      abs, min, max, round,
    };
    
    // Replace Excel-style & with + for string concat in mathjs
    let expression = formula.replace(/\s*&\s*/g, ' + ');
    
    const result = evaluate(expression, scope);
    
    if (typeof result === "number") {
      return isNaN(result) || !isFinite(result) ? null : Math.round(result * 10000) / 10000;
    }
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
    const scope: Record<string, unknown> = {
      ...sampleContext,
      trim, upper, lower, left, right, mid, len, replace, remove, clean, concat,
      abs, min, max, round,
    };
    
    let expression = formula.replace(/\s*&\s*/g, ' + ');
    
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
  // Match single letters or double letters (Excel-style)
  const matches = formula.match(/\b[A-Z]{1,2}\b/g) || [];
  return Array.from(new Set(matches));
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
