import { evaluate } from "mathjs";

export interface FormulaContext {
  [columnKey: string]: number | string | null;
}

export function evaluateFormula(formula: string, context: FormulaContext): number | string | null {
  try {
    let expression = formula;
    for (const [key, value] of Object.entries(context)) {
      const numVal = typeof value === "string" ? parseFloat(value) : value;
      if (numVal !== null && !isNaN(numVal as number)) {
        expression = expression.replace(new RegExp(`\\b${key}\\b`, "g"), String(numVal));
      }
    }
    const result = evaluate(expression);
    if (typeof result === "number") {
      return isNaN(result) || !isFinite(result) ? null : Math.round(result * 10000) / 10000;
    }
    return String(result);
  } catch {
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
    let expression = formula;
    for (const [key, value] of Object.entries(sampleContext)) {
      const numVal = typeof value === "string" ? parseFloat(value) : value;
      expression = expression.replace(new RegExp(`\\b${key}\\b`, "g"), String(numVal ?? 0));
    }
    evaluate(expression);
    return { valid: true };
  } catch (e) {
    return { valid: false, error: String(e) };
  }
}

export function extractColumnRefs(formula: string): string[] {
  const matches = formula.match(/\bcol_\d+\b/g) || [];
  return Array.from(new Set(matches));
}
