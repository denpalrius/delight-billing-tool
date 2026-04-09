import { ProcessingError, processFilesWithStats } from "./lib/billing";

interface Env {
  ASSETS: Fetcher;
}

const MAX_FILES = 20;
const MAX_FILE_SIZE_BYTES = 10 * 1024 * 1024;
const MAX_TOTAL_SIZE_BYTES = 40 * 1024 * 1024;

export default {
  async fetch(request, env): Promise<Response> {
    const url = new URL(request.url);

    if (url.pathname === "/api/health") {
      return Response.json({ ok: true });
    }

    if (url.pathname === "/api/process") {
      return handleProcessRequest(request);
    }

    return env.ASSETS.fetch(request);
  },
} satisfies ExportedHandler<Env>;

async function handleProcessRequest(request: Request): Promise<Response> {
  const requestId = crypto.randomUUID().slice(0, 8);

  if (request.method !== "POST") {
    return jsonError("Method not allowed.", 405);
  }

  console.log(`[${requestId}] Received processing request.`);
  const formData = await request.formData();
  const files = Array.from(formData.values()).filter(
    (value): value is File => value instanceof File && value.size > 0,
  );

  if (files.length === 0) {
    return jsonError("Upload at least one Excel file.", 400);
  }

  if (files.length > MAX_FILES) {
    return jsonError(`Too many files. Maximum allowed is ${MAX_FILES}.`, 400);
  }

  const totalSize = files.reduce((sum, file) => sum + file.size, 0);
  console.log(
    `[${requestId}] Files received: ${files.length}, total bytes: ${totalSize}.`,
  );
  if (totalSize > MAX_TOTAL_SIZE_BYTES) {
    return jsonError("Uploads are too large for this deployment.", 400);
  }

  for (const file of files) {
    if (file.size > MAX_FILE_SIZE_BYTES) {
      return jsonError(`"${file.name}" exceeds the file size limit.`, 400);
    }
  }

  try {
    const preparedFiles = await Promise.all(
      files.map(async (file) => ({
        name: file.name,
        content: new Uint8Array(await file.arrayBuffer()),
      })),
    );

    for (const file of files) {
      console.log(
        `[${requestId}] Preparing file "${file.name}" (${file.size} bytes).`,
      );
    }

    const { workbookBytes, stats } = processFilesWithStats(preparedFiles);
    const responseBody = new Uint8Array(workbookBytes.byteLength);
    responseBody.set(workbookBytes);
    const today = new Date().toISOString().slice(0, 10);

    console.log(
      `[${requestId}] Workbook ready. rows=${stats.rowCount}, dates=${stats.dateCount}, providers=${stats.providerCount}, individuals=${stats.individualCount}.`,
    );

    return new Response(responseBody, {
      status: 200,
      headers: {
        "Content-Type":
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": `attachment; filename="daily_summary_${today}.xlsx"`,
        "Cache-Control": "no-store",
      },
    });
  } catch (error) {
    console.error(`[${requestId}] Processing failed.`, error);
    if (error instanceof ProcessingError) {
      return jsonError(error.message, 422);
    }

    return jsonError("The uploaded files could not be processed.", 500);
  }
}

function jsonError(message: string, status: number): Response {
  return Response.json({ error: message }, { status });
}
