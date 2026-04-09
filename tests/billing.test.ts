import * as XLSX from "xlsx";
import { describe, expect, it } from "vitest";

import { parseFile, processFiles } from "../src/lib/billing";

describe("billing processor", () => {
  it("parses the legacy raw worksheet shape", () => {
    const content = buildLegacyInputWorkbook([
      ["07/02/2025", 480, "Provider Alpha"],
      ["07/02/2025", 720, "Provider Beta"],
      ["07/03/2025", 240, "Provider Alpha"],
    ]);

    const rows = parseFile(content, "sample.xlsx");

    expect(rows).toEqual([
      {
        date: "2025-07-02",
        durationHours: 8,
        individual: "DD",
        serviceProvider: "Provider Alpha",
      },
      {
        date: "2025-07-02",
        durationHours: 12,
        individual: "DD",
        serviceProvider: "Provider Beta",
      },
      {
        date: "2025-07-03",
        durationHours: 4,
        individual: "DD",
        serviceProvider: "Provider Alpha",
      },
    ]);
  });

  it("builds the DailyMatrix workbook from uploaded files", () => {
    const first = buildLegacyInputWorkbook([
      ["07/02/2025", 480, "Provider Alpha"],
      ["07/02/2025", 720, "Provider Beta"],
    ]);
    const second = buildLegacyInputWorkbook(
      [
        ["07/02/2025", 240, "Provider Alpha"],
        ["07/03/2025", 480, "Provider Beta"],
      ],
      "Dawn Marie, Client",
    );

    const output = processFiles([
      { name: "dd.xlsx", content: first },
      { name: "dm.xlsx", content: second },
    ]);

    const workbook = XLSX.read(output, { type: "array", cellFormula: true });
    const worksheet = workbook.Sheets.DailyMatrix;
    const rows = XLSX.utils.sheet_to_json<(string | number | null)[]>(
      worksheet,
      { header: 1, raw: false, defval: null },
    );

    expect(rows[0]?.[0]).toBe("07/02/2025");
    expect(rows[1]).toEqual(["Service Provider", "DD", "DM", "Provider Total"]);
    expect(rows[2]?.slice(0, 4)).toEqual([
      "Provider Alpha",
      "8",
      "4",
      "12",
    ]);
    expect(rows[3]?.slice(0, 4)).toEqual([
      "Provider Beta",
      "12",
      "0",
      "12",
    ]);
    expect(rows[4]?.slice(0, 4)).toEqual([
      "Total hours for individual",
      "20",
      "4",
      null,
    ]);
    expect(worksheet.D3?.f).toBe("SUM(B3:C3)");
    expect(rows[7]?.[0]).toBe("07/03/2025");
  });

  it("parses an existing DailyMatrix workbook and regenerates it", () => {
    const sourceWorkbook = buildDailyMatrixWorkbook([
      ["07/02/2025", [
        ["Provider Alpha", 8, 8, 0],
        ["Provider Beta", 12, 12, 0],
      ]],
      ["07/03/2025", [
        ["Provider Beta", 0, 12, 12],
      ]],
    ]);

    const output = processFiles([{ name: "matrix.xlsx", content: sourceWorkbook }]);
    const workbook = XLSX.read(output, { type: "array", cellFormula: true });
    const rows = XLSX.utils.sheet_to_json<(string | number | null)[]>(
      workbook.Sheets.DailyMatrix,
      { header: 1, raw: false, defval: null },
    );

    expect(rows[0]).toEqual(["07/02/2025", null, null, null, null]);
    expect(rows[1]).toEqual(["Service Provider", "DD", "DM", "OT", "Provider Total"]);
    expect(rows[2]).toEqual(["Provider Alpha", "8", "8", "0", "16"]);
    expect(rows[3]).toEqual(["Provider Beta", "12", "12", "0", "24"]);
    expect(rows[7]).toEqual(["07/03/2025", null, null, null, null]);
  });
});

function buildLegacyInputWorkbook(
  entries: Array<[string, number, string]>,
  individualName = "Dawn Doe, Client",
): Uint8Array {
  const rows: Array<Array<string | number | null>> = [
    [null, null, null, null, null, null, null],
    [null, null, null, null, null, null, null],
    [null, null, null, individualName, null, null, null],
    ["Time Zone:", null, null, null, null, null, null],
    ["Date", null, null, null, null, null, null],
    ...entries.map(([date, duration, provider]) => [
      date,
      null,
      null,
      null,
      duration,
      null,
      provider,
    ]),
    [null, null, "Total", null, null, null, null],
  ];

  const worksheet = XLSX.utils.aoa_to_sheet(rows);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");
  return XLSX.write(workbook, { type: "array", bookType: "xlsx" }) as Uint8Array;
}

function buildDailyMatrixWorkbook(
  blocks: Array<[string, Array<[string, number, number, number]>]>,
): Uint8Array {
  const rows: Array<Array<string | number | null>> = [];

  for (const [date, providerRows] of blocks) {
    rows.push([date, null, null, null, null]);
    rows.push(["Service Provider", "DD", "DM", "OT", "Provider Total"]);

    let dd = 0;
    let dm = 0;
    let ot = 0;

    for (const [provider, ddHours, dmHours, otHours] of providerRows) {
      dd += ddHours;
      dm += dmHours;
      ot += otHours;
      rows.push([provider, ddHours, dmHours, otHours, ddHours + dmHours + otHours]);
    }

    rows.push(["Total hours for individual", dd, dm, ot, null]);
    rows.push(["Total hrs pending in a 24hr period", 24 - dd, 24 - dm, 24 - ot, null]);
    rows.push([null, null, null, null, null]);
  }

  const worksheet = XLSX.utils.aoa_to_sheet(rows);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "DailyMatrix");
  return XLSX.write(workbook, { type: "array", bookType: "xlsx" }) as Uint8Array;
}
