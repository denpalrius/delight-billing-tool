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

    if (url.pathname.startsWith("/api/")) {
      return notFoundResponse(request);
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

function notFoundResponse(request: Request): Response {
  const acceptsHtml = request.headers.get("accept")?.includes("text/html");

  if (!acceptsHtml) {
    return jsonError("Not found.", 404);
  }

  const html = `<!doctype html>
<html lang="en">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>Page Not Found | Delight Billing Tool</title>
    <style>
      :root {
        color-scheme: light;
        --primary: #7ecfd8;
        --secondary: #ffd43b;
        --accent: #4b2580;
        --background: #fffdfa;
        --surface: rgba(255, 255, 255, 0.88);
        --surface-border: rgba(75, 37, 128, 0.1);
        --text: #2e2e2e;
        --muted: #736a63;
      }

      * { box-sizing: border-box; }

      body {
        margin: 0;
        min-height: 100vh;
        display: grid;
        place-items: center;
        padding: 1.5rem;
        font-family: "Inter", sans-serif;
        color: var(--text);
        background:
          radial-gradient(circle at 15% 10%, rgba(255, 212, 59, 0.22), transparent 22rem),
          radial-gradient(circle at 85% 18%, rgba(126, 207, 216, 0.22), transparent 25rem),
          linear-gradient(180deg, #fffdfa 0%, #fdfaf4 52%, #f7fcfc 100%);
      }

      main {
        width: min(42rem, 100%);
        padding: 2rem;
        border: 1px solid var(--surface-border);
        border-radius: 28px;
        background: var(--surface);
        box-shadow:
          0 16px 50px rgba(126, 207, 216, 0.14),
          0 22px 80px rgba(75, 37, 128, 0.14);
      }

      p {
        margin: 0;
        color: var(--muted);
        line-height: 1.6;
      }

      .eyebrow {
        margin-bottom: 0.75rem;
        letter-spacing: 0.16em;
        text-transform: uppercase;
        font-size: 0.82rem;
      }

      h1 {
        margin: 0 0 0.85rem;
        font-family: "Playfair Display", Georgia, serif;
        font-size: clamp(2rem, 4vw, 3rem);
        line-height: 1.08;
        letter-spacing: -0.03em;
        color: var(--accent);
      }

      a {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        margin-top: 1.4rem;
        min-width: 11rem;
        padding: 0.9rem 1.2rem;
        border-radius: 999px;
        font-weight: 700;
        color: white;
        text-decoration: none;
        background: linear-gradient(135deg, var(--primary) 0%, #63b7c3 100%);
        box-shadow: 0 14px 26px rgba(126, 207, 216, 0.24);
      }

      strong { color: var(--accent); }
    </style>
  </head>
  <body>
    <main>
      <p class="eyebrow">Delight Billing Tool</p>
      <h1>That page is not available.</h1>
      <p>
        The requested path could not be found on the billing service. If you are
        looking for the workbook uploader, return to the main billing page.
      </p>
      <a href="/">Back to billing tool</a>
    </main>
  </body>
</html>`;

  return new Response(html, {
    status: 404,
    headers: {
      "Content-Type": "text/html; charset=UTF-8",
      "Cache-Control": "no-store",
    },
  });
}
