import fs from "node:fs";
import path from "node:path";

import { chromium } from "playwright";
import * as XLSX from "xlsx";

const cwd = process.cwd();
const inputArg = process.argv[2] ?? "tests/fixtures/valid-input.xlsx";
const baseUrl = process.env.SMOKE_BASE_URL ?? "http://localhost:8788";
const inputPath = path.resolve(cwd, inputArg);
const outputDir = path.resolve(cwd, "output", "playwright");
const outputPath = path.join(
  outputDir,
  `${path.basename(inputPath, path.extname(inputPath))}.processed.xlsx`,
);
const beforeShot = path.join(outputDir, "before-submit.png");
const afterShot = path.join(outputDir, "after-submit.png");
const keepWorkbookOutput = process.env.KEEP_SMOKE_OUTPUT === "1";

if (!fs.existsSync(inputPath)) {
  console.error(`Input file not found: ${inputPath}`);
  process.exit(1);
}

fs.mkdirSync(outputDir, { recursive: true });

const browser = await chromium.launch({ channel: "chrome", headless: true });
const context = await browser.newContext({ acceptDownloads: true });
const page = await context.newPage();
const consoleMessages = [];
page.on("console", (msg) => consoleMessages.push(`${msg.type()}: ${msg.text()}`));

try {
  await page.goto(baseUrl, { waitUntil: "networkidle" });
  await page.screenshot({ path: beforeShot, fullPage: true });
  await page.locator("#files").setInputFiles(inputPath);

  const [download] = await Promise.all([
    page.waitForEvent("download"),
    page.locator("#submit-button").click(),
  ]);

  await download.saveAs(outputPath);
  await page.waitForTimeout(500);
  const statusText = await page.locator("#status").textContent();
  await page.screenshot({ path: afterShot, fullPage: true });

  const workbook = XLSX.read(fs.readFileSync(outputPath), {
    type: "buffer",
    cellFormula: true,
  });
  const sheet = workbook.Sheets.DailyMatrix;
  const rows = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    raw: false,
    defval: null,
  });

  const nonEmptyRows = rows.filter((row) => row.some((cell) => cell !== null)).length;
  const formulaCells = Object.keys(sheet).filter(
    (key) => /^[A-Z]+\d+$/.test(key) && typeof sheet[key]?.f === "string",
  ).length;

  console.log(
    JSON.stringify(
      {
        inputPath,
        statusText,
        downloadPath: keepWorkbookOutput ? outputPath : "(deleted after validation)",
        screenshots: [beforeShot, afterShot],
        workbookChecks: {
          sheetNames: workbook.SheetNames,
          nonEmptyRows,
          formulaCells,
          firstSheetRef: sheet["!ref"] ?? null,
        },
        consoleMessages,
      },
      null,
      2,
    ),
  );

  if (!keepWorkbookOutput) {
    fs.rmSync(outputPath, { force: true });
  }
} finally {
  await browser.close();
}
