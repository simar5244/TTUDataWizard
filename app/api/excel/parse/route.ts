import { NextRequest, NextResponse } from "next/server";
import { getServerSession } from "next-auth";
import { authOptions } from "@/lib/auth";
import * as XLSX from "xlsx";

function detectDataType(values: (string | number | boolean | null)[]): "string" | "number" | "date" | "boolean" | "empty" {
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

export async function POST(req: NextRequest) {
  const session = await getServerSession(authOptions);
  if (!session?.user) return NextResponse.json({ error: "Unauthorized" }, { status: 401 });

  const schemaOnly = req.nextUrl.searchParams.get("schemaOnly") === "true";

  try {
    const formData = await req.formData();
    const file = formData.get("file") as File | null;
    if (!file) return NextResponse.json({ error: "No file provided" }, { status: 400 });

    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(Buffer.from(buffer), { type: "buffer", cellDates: true });

    const result: {
      fileName: string;
      fileSize: number;
      sheets: {
        name: string;
        rowCount: number;
        columns: { key: string; header: string; index: number; dataType: string; sampleValues: unknown[] }[];
        data?: Record<string, unknown>[];
      }[];
    } = {
      fileName: file.name,
      fileSize: file.size,
      sheets: [],
    };

    for (const sheetName of workbook.SheetNames) {
      const worksheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json<Record<string, string | number | boolean | null>>(worksheet, {
        raw: false,
        defval: null,
      });
      if (jsonData.length === 0) continue;

      const headers = Object.keys(jsonData[0] || {});
      const columns = headers.map((header, index) => {
        const sampleValues = jsonData.slice(0, 10).map((row) => row[header] ?? null);
        return {
          key: `col_${index}`,
          header,
          index,
          dataType: detectDataType(sampleValues),
          sampleValues,
        };
      });

      result.sheets.push({
        name: sheetName,
        rowCount: jsonData.length,
        columns,
        ...(schemaOnly ? {} : { data: jsonData }),
      });
    }

    return NextResponse.json(result);
  } catch (err) {
    return NextResponse.json({ error: "Failed to parse file", detail: String(err) }, { status: 500 });
  }
}
