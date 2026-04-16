interface DynamicColumnNameOptions {
  existingTitles?: string[];
  ensureUnique?: boolean;
  incrementStep?: number;
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

export function formatDynamicColumnName(template: string, options: DynamicColumnNameOptions = {}): string {
  const now = new Date();
  const yyyy = String(now.getFullYear());
  const mm = String(now.getMonth() + 1).padStart(2, "0");
  const dd = String(now.getDate()).padStart(2, "0");
  const baseTemplate = (template || "Enrollments {{DATE}}").trim();
  const cleanTemplate = baseTemplate
    .replace(/\{\{\s*YYYY-MM-DD\s*\}\}/gi, `${yyyy}-${mm}-${dd}`)
    .replace(/\{\{\s*YYYYMMDD\s*\}\}/gi, `${yyyy}${mm}${dd}`)
    .replace(/\{\{\s*MM\/DD\/YYYY\s*\}\}/gi, `${mm}/${dd}/${yyyy}`)
    .replace(/\{\{\s*DATE\s*\}\}/gi, `${mm}/${dd}/${yyyy}`)
    .trim() || `Enrollments ${mm}/${dd}/${yyyy}`;
  const existingTitles = (options.existingTitles ?? []).map((t) => String(t));
  const fallbackIncrementStep = Number.isFinite(Number(options.incrementStep)) && Number(options.incrementStep) > 0
    ? Math.floor(Number(options.incrementStep))
    : 1;

  const incMatch = cleanTemplate.match(/\{\{\s*incrementer(?:\s*:\s*(\d+))?\s*\}\}/i);
  const decMatch = cleanTemplate.match(/\{\{\s*decrementer(?:\s*:\s*(\d+))?\s*\}\}/i);
  const hasIncrementer = Boolean(incMatch);
  const hasDecrementer = Boolean(decMatch);
  const tokenStep = hasIncrementer
    ? Number(incMatch?.[1] || fallbackIncrementStep)
    : hasDecrementer
      ? Number(decMatch?.[1] || fallbackIncrementStep)
      : fallbackIncrementStep;
  const incrementStep = Number.isFinite(tokenStep) && tokenStep > 0 ? Math.floor(tokenStep) : 1;

  let candidate = cleanTemplate
    .replace(/\{\{\s*incrementer(?:\s*:\s*\d+)?\s*\}\}/gi, "")
    .replace(/\{\{\s*decrementer(?:\s*:\s*\d+)?\s*\}\}/gi, "")
    .replace(/\s{2,}/g, " ")
    .trim();

  if ((hasIncrementer || hasDecrementer) && candidate) {
    const numberMatch = Array.from(candidate.matchAll(/-?\d+/g)).at(-1);
    if (numberMatch && typeof numberMatch.index === "number") {
      const start = numberMatch.index;
      const end = start + numberMatch[0].length;
      const baseNumber = Number(numberMatch[0]);
      const prefix = candidate.slice(0, start);
      const suffix = candidate.slice(end);
      const regex = new RegExp(`^${escapeRegExp(prefix)}(-?\\d+)${escapeRegExp(suffix)}$`, "i");
      const seenNumbers = existingTitles
        .map((title) => title.match(regex))
        .filter((m): m is RegExpMatchArray => Boolean(m))
        .map((m) => Number(m[1]))
        .filter((n) => Number.isFinite(n));

      let resolved = baseNumber;
      if (hasIncrementer && seenNumbers.length > 0) {
        resolved = Math.max(...seenNumbers) + incrementStep;
      } else if (hasDecrementer && seenNumbers.length > 0) {
        resolved = Math.min(...seenNumbers) - incrementStep;
      }
      candidate = `${prefix}${resolved}${suffix}`.trim();
    }
  }

  if (options.ensureUnique && !hasIncrementer && !hasDecrementer) {
    if (!existingTitles.includes(candidate)) return candidate;
    let suffix = 2;
    let unique = `${candidate} (${suffix})`;
    while (existingTitles.includes(unique)) {
      suffix += 1;
      unique = `${candidate} (${suffix})`;
    }
    return unique;
  }

  return candidate;
}
