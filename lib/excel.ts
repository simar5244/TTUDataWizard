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

export interface SchemaFingerprint {
  sheetName: string;
  columnNames: string[];
  columnOrder: string[];
  dataTypes: Record<string, string>;
  rowCountApprox: number;
}

function detectDataType(values: (string | number | boolean | null)[]): ExcelColumn["dataType"] {
  const nonNull = values.filter((v) => v !== null && v !== undefined && v !== "");
  if (nonNull.length === 0) return "empty";
  const numericCount = nonNull.filter((v) => typeof v === "number" || !isNaN(Number(v))).length;
  if (numericCount / nonNull.length > 0.8) return "number";
  const dateCount = nonNull.filter((v) => {
    const s = String(v);
    return /^\d{1,4}[-/]\d{1,2}[-/]\d{1,4}/.test(s) || !isNaN(Date.parse(s));
  }).length;
  if (dateCount / nonNull.length > 0.7) return "date";
  return "string";
}

export function parseExcelFile(buffer: ArrayBuffer, fileName: string, fileSize: number): ParsedExcel {
  const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
  const sheets: ExcelSheet[] = [];

  for (const sheetName of workbook.SheetNames) {
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json<Record<string, string | number | boolean | null>>(worksheet, {
      raw: false,
      defval: null,
    });

    if (jsonData.length === 0) continue;

    const headers = Object.keys(jsonData[0] || {});
    const columns: ExcelColumn[] = headers.map((header, index) => {
      const sampleValues = jsonData.slice(0, 10).map((row) => row[header] ?? null);
      return {
        key: `col_${index}`,
        header,
        index,
        dataType: detectDataType(sampleValues),
        sampleValues,
      };
    });

    sheets.push({
      name: sheetName,
      columns,
      rowCount: jsonData.length,
      data: jsonData,
    });
  }

  return { sheets, fileName, fileSize };
}

export function generateSchemaFingerprint(sheet: ExcelSheet): SchemaFingerprint {
  const dataTypes: Record<string, string> = {};
  for (const col of sheet.columns) {
    dataTypes[col.header] = col.dataType;
  }
  return {
    sheetName: sheet.name,
    columnNames: sheet.columns.map((c) => c.header),
    columnOrder: sheet.columns.map((c) => c.header),
    dataTypes,
    rowCountApprox: sheet.rowCount,
  };
}

export interface ValidationResult {
  status: "exact" | "remapped" | "blocked";
  missingColumns: string[];
  remappedColumns: { expected: string; found: string; fromIndex: number; toIndex: number }[];
  extraColumns: string[];
}

export function validateExcelAgainstFingerprint(
  sheet: ExcelSheet,
  fingerprint: SchemaFingerprint
): ValidationResult {
  const uploadedHeaders = sheet.columns.map((c) => c.header);
  const expectedHeaders = Array.isArray(fingerprint.columnNames)
    ? fingerprint.columnNames
    : Array.isArray(fingerprint.columnOrder)
      ? fingerprint.columnOrder
      : [];

  const missingColumns: string[] = [];
  const remappedColumns: ValidationResult["remappedColumns"] = [];
  const extraColumns: string[] = [];

  for (const expected of expectedHeaders) {
    const exactIdx = uploadedHeaders.indexOf(expected);
    if (exactIdx !== -1) {
      const originalIdx = Array.isArray(fingerprint.columnOrder)
        ? fingerprint.columnOrder.indexOf(expected)
        : expectedHeaders.indexOf(expected);
      if (originalIdx !== exactIdx) {
        remappedColumns.push({
          expected,
          found: expected,
          fromIndex: originalIdx,
          toIndex: exactIdx,
        });
      }
    } else {
      const caseInsensitiveIdx = uploadedHeaders.findIndex(
        (h) => h.toLowerCase() === expected.toLowerCase()
      );
      if (caseInsensitiveIdx !== -1) {
        remappedColumns.push({
          expected,
          found: uploadedHeaders[caseInsensitiveIdx],
            fromIndex: Array.isArray(fingerprint.columnOrder)
              ? fingerprint.columnOrder.indexOf(expected)
              : expectedHeaders.indexOf(expected),
          toIndex: caseInsensitiveIdx,
        });
      } else {
        missingColumns.push(expected);
      }
    }
  }

  for (const header of uploadedHeaders) {
    if (!expectedHeaders.some((e) => e.toLowerCase() === header.toLowerCase())) {
      extraColumns.push(header);
    }
  }

  if (missingColumns.length > 0) return { status: "blocked", missingColumns, remappedColumns, extraColumns };
  if (remappedColumns.length > 0) return { status: "remapped", missingColumns, remappedColumns, extraColumns };
  return { status: "exact", missingColumns, remappedColumns, extraColumns };
}
