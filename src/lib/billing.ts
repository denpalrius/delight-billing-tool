import * as XLSX from "xlsx";

export interface UploadedFile {
  name: string;
  content: Uint8Array;
}

export interface ProcessingStats {
  fileCount: number;
  rowCount: number;
  dateCount: number;
  providerCount: number;
  individualCount: number;
}

interface ParsedRow {
  date: string;
  serviceProvider: string;
  individual: string;
  durationHours: number;
}

type RawCell = string | number | boolean | Date | null | undefined;
type RawSheet = RawCell[][];
type AoACell = string | number | XLSX.CellObject | null;

export class ProcessingError extends Error {}

export function processFiles(files: UploadedFile[]): Uint8Array {
  return processFilesWithStats(files).workbookBytes;
}

export function processFilesWithStats(
  files: UploadedFile[],
): { workbookBytes: Uint8Array; stats: ProcessingStats } {
  const rows = files.flatMap((file) => parseFile(file.content, file.name));

  if (rows.length === 0) {
    throw new ProcessingError("No valid 'Date' blocks were found in the uploaded files.");
  }

  const stats: ProcessingStats = {
    fileCount: files.length,
    rowCount: rows.length,
    dateCount: uniqueSorted(rows.map((row) => row.date)).length,
    providerCount: uniqueSorted(rows.map((row) => row.serviceProvider)).length,
    individualCount: uniqueSorted(rows.map((row) => row.individual)).length,
  };

  return {
    workbookBytes: buildSummaryWorkbook(rows),
    stats,
  };
}

export function parseFile(content: Uint8Array, filename: string): ParsedRow[] {
  const workbook = XLSX.read(content, {
    type: "array",
    cellFormula: true,
    cellDates: false,
    dense: false,
  });

  const firstSheetName = workbook.SheetNames[0];
  if (!firstSheetName) {
    return [];
  }

  const worksheet = workbook.Sheets[firstSheetName];
  const raw = XLSX.utils.sheet_to_json<RawCell[]>(worksheet, {
    header: 1,
    raw: true,
    defval: null,
    blankrows: false,
  });

  if (isDailyMatrixSheet(raw)) {
    return parseDailyMatrixSheet(raw);
  }

  const individual = getAcronym(getIndividualName(worksheet, filename));
  const headerRows = findAllSections(raw);
  const parsedRows: ParsedRow[] = [];

  for (const headerRow of headerRows) {
    const endRow = findSectionEnd(raw, headerRow);
    const section = raw.slice(headerRow + 1, endRow);

    for (const row of section) {
      const normalizedDate = normalizeDate(row[0]);
      if (!normalizedDate) {
        continue;
      }

      const durationMinutes = toNumber(row[4]);
      const serviceProvider = String(row[6] ?? "").trim();

      if (!Number.isFinite(durationMinutes) || !serviceProvider) {
        continue;
      }

      parsedRows.push({
        date: normalizedDate,
        serviceProvider,
        individual,
        durationHours: durationMinutes / 60,
      });
    }
  }

  return parsedRows;
}

export function buildSummaryWorkbook(rows: ParsedRow[]): Uint8Array {
  const workbook = XLSX.utils.book_new();
  const individuals = uniqueSorted(rows.map((row) => row.individual));
  const dates = uniqueSorted(rows.map((row) => row.date));
  const sheetRows: AoACell[][] = [];
  const formulaCells: Array<{ cell: string; formula: string; value: number }> = [];

  for (const date of dates) {
    const dayRows = rows.filter((row) => row.date === date);
    const providerNames = uniqueInOrder(dayRows.map((row) => row.serviceProvider));

    sheetRows.push([formatDateForDisplay(date)]);
    sheetRows.push(["Service Provider", ...individuals, "Provider Total"]);

    const providerStart = sheetRows.length + 1;

    for (const providerName of providerNames) {
      const excelRowNumber = sheetRows.length + 1;
      const providerRow: AoACell[] = [providerName];
      let providerTotal = 0;

      for (const individual of individuals) {
        const totalHours = roundToTwo(
          dayRows
            .filter(
              (row) =>
                row.serviceProvider === providerName && row.individual === individual,
            )
            .reduce((sum, row) => sum + row.durationHours, 0),
        );

        providerTotal += totalHours;
        providerRow.push(makeNumberCell(totalHours));
      }

      const providerTotalColumn = columnLetter(individuals.length + 2);
      providerRow.push(null);
      formulaCells.push({
        cell: `${providerTotalColumn}${excelRowNumber}`,
        formula: `SUM(B${excelRowNumber}:${columnLetter(individuals.length + 1)}${excelRowNumber})`,
        value: roundToTwo(providerTotal),
      });
      sheetRows.push(providerRow);
    }

    const providerEnd = sheetRows.length;
    const totalsRow: AoACell[] = ["Total hours for individual"];
    for (let index = 0; index < individuals.length; index += 1) {
      const column = columnLetter(index + 2);
      const individual = individuals[index];
      const totalForIndividual = roundToTwo(
        dayRows
          .filter((row) => row.individual === individual)
          .reduce((sum, row) => sum + row.durationHours, 0),
      );
      totalsRow.push(null);
      formulaCells.push({
        cell: `${column}${sheetRows.length + 1}`,
        formula: `SUM(${column}${providerStart}:${column}${providerEnd})`,
        value: totalForIndividual,
      });
    }
    sheetRows.push(totalsRow);

    const pendingRowNumber = sheetRows.length + 1;
    const pendingRow: AoACell[] = ["Total hrs pending in a 24hr period"];
    for (let index = 0; index < individuals.length; index += 1) {
      const column = columnLetter(index + 2);
      const individual = individuals[index];
      const totalForIndividual = roundToTwo(
        dayRows
          .filter((row) => row.individual === individual)
          .reduce((sum, row) => sum + row.durationHours, 0),
      );
      pendingRow.push(null);
      formulaCells.push({
        cell: `${column}${pendingRowNumber}`,
        formula: `24 - ${column}${pendingRowNumber - 1}`,
        value: roundToTwo(24 - totalForIndividual),
      });
    }
    sheetRows.push(pendingRow);
    sheetRows.push([]);
  }

  const worksheet = XLSX.utils.aoa_to_sheet(sheetRows);
  for (const { cell, formula, value } of formulaCells) {
    worksheet[cell] = {
      t: "n",
      f: formula,
      v: value,
    };
  }
  worksheet["!cols"] = [
    { wch: 32 },
    ...individuals.map(() => ({ wch: 12 })),
    { wch: 16 },
  ];

  XLSX.utils.book_append_sheet(workbook, worksheet, "DailyMatrix");
  const output = XLSX.write(workbook, {
    bookType: "xlsx",
    type: "array",
  }) as ArrayBuffer | Uint8Array;

  return output instanceof Uint8Array ? output : new Uint8Array(output);
}

function getIndividualName(sheet: XLSX.WorkSheet, filename: string): string {
  const cellValue = sheet.D3?.v;

  if (cellValue == null || String(cellValue).trim() === "") {
    throw new ProcessingError(`"${filename}" is missing the expected D3 individual name cell.`);
  }

  return String(cellValue).split(",")[0].trim();
}

function isDailyMatrixSheet(raw: RawSheet): boolean {
  if (raw.length < 2) {
    return false;
  }

  return (
    normalizeDate(raw[0]?.[0] ?? null) !== null &&
    String(raw[1]?.[0] ?? "").trim() === "Service Provider"
  );
}

function parseDailyMatrixSheet(raw: RawSheet): ParsedRow[] {
  const parsedRows: ParsedRow[] = [];
  let rowIndex = 0;

  while (rowIndex < raw.length) {
    const date = normalizeDate(raw[rowIndex]?.[0] ?? null);
    const nextHeader = String(raw[rowIndex + 1]?.[0] ?? "").trim();

    if (!date || nextHeader !== "Service Provider") {
      rowIndex += 1;
      continue;
    }

    const headerRow = raw[rowIndex + 1] ?? [];
    const providerTotalIndex = headerRow.findIndex(
      (value) => String(value ?? "").trim() === "Provider Total",
    );
    const individualNames = headerRow
      .slice(1, providerTotalIndex >= 0 ? providerTotalIndex : undefined)
      .map((value) => String(value ?? "").trim())
      .filter(Boolean);

    rowIndex += 2;

    while (rowIndex < raw.length) {
      const row = raw[rowIndex] ?? [];
      const label = String(row[0] ?? "").trim();

      const nextRowStartsBlock =
        normalizeDate(row[0] ?? null) !== null &&
        String(raw[rowIndex + 1]?.[0] ?? "").trim() === "Service Provider";

      if (!label || nextRowStartsBlock || label === "Service Provider") {
        break;
      }

      if (
        label === "Total hours for individual" ||
        label === "Total hrs pending in a 24hr period"
      ) {
        rowIndex += 1;
        continue;
      }

      for (let individualIndex = 0; individualIndex < individualNames.length; individualIndex += 1) {
        const durationHours = toNumber(row[individualIndex + 1]);
        if (!Number.isFinite(durationHours) || durationHours <= 0) {
          continue;
        }

        parsedRows.push({
          date,
          serviceProvider: label,
          individual: individualNames[individualIndex],
          durationHours,
        });
      }

      rowIndex += 1;
    }
  }

  return parsedRows;
}

function getAcronym(name: string): string {
  return name
    .split(/\s+/)
    .filter(Boolean)
    .map((part) => part[0]?.toUpperCase() ?? "")
    .join("");
}

function findAllSections(raw: RawSheet): number[] {
  const timeZoneIndex = raw.findIndex((row) =>
    String(row[0] ?? "").toLowerCase().includes("time zone:"),
  );

  return raw
    .map((row, index) => ({ row, index }))
    .filter(
      ({ row, index }) =>
        index > timeZoneIndex && String(row[0] ?? "").trim().toLowerCase() === "date",
    )
    .map(({ index }) => index);
}

function findSectionEnd(raw: RawSheet, headerRow: number): number {
  for (let index = headerRow + 1; index < raw.length; index += 1) {
    if (String(raw[index]?.[2] ?? "").trim().toLowerCase() === "total") {
      return index;
    }
  }

  return raw.length;
}

function normalizeDate(value: RawCell): string | null {
  if (value == null || value === "") {
    return null;
  }

  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value.toISOString().slice(0, 10);
  }

  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (!parsed) {
      return null;
    }

    return `${String(parsed.y).padStart(4, "0")}-${String(parsed.m).padStart(2, "0")}-${String(parsed.d).padStart(2, "0")}`;
  }

  const stringValue = String(value).trim();
  const slashMatch = stringValue.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (slashMatch) {
    const month = Number(slashMatch[1]);
    const day = Number(slashMatch[2]);
    const year = slashMatch[3].length === 2 ? 2000 + Number(slashMatch[3]) : Number(slashMatch[3]);
    return `${String(year).padStart(4, "0")}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
  }

  const parsed = new Date(stringValue);
  if (Number.isNaN(parsed.getTime())) {
    return null;
  }

  return parsed.toISOString().slice(0, 10);
}

function toNumber(value: RawCell): number {
  if (typeof value === "number") {
    return value;
  }

  const parsed = Number(String(value ?? "").trim());
  return Number.isFinite(parsed) ? parsed : Number.NaN;
}

function roundToTwo(value: number): number {
  return Math.round(value * 100) / 100;
}

function makeNumberCell(value: number): AoACell {
  if (Number.isInteger(value)) {
    return value;
  }

  return {
    t: "n",
    v: value,
    z: "0.00",
  };
}

function uniqueInOrder(values: string[]): string[] {
  const seen = new Set<string>();
  const result: string[] = [];

  for (const value of values) {
    if (seen.has(value)) {
      continue;
    }

    seen.add(value);
    result.push(value);
  }

  return result;
}

function uniqueSorted(values: string[]): string[] {
  return Array.from(new Set(values)).sort((left, right) => left.localeCompare(right));
}

function formatDateForDisplay(date: string): string {
  const [year, month, day] = date.split("-");
  return `${month}/${day}/${year}`;
}

function columnLetter(columnNumber: number): string {
  let value = columnNumber;
  let result = "";

  while (value > 0) {
    const remainder = (value - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    value = Math.floor((value - 1) / 26);
  }

  return result;
}
