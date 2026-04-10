# Delight Billing Tool

Cloudflare Workers rewrite of the Delight billing workflow.

## Architecture

- `public/`: static frontend assets
- `src/`: Worker API and Excel processing logic
- `legacy/`: preserved Streamlit implementation

The production flow is:

1. Browser uploads one or more Excel files to `POST /api/process`
2. Worker parses the files and generates the summary workbook server-side
3. Browser downloads the generated `.xlsx`

## Local Development

```bash
npm install
npm run dev
```

Then open the local Wrangler URL in a browser.

## Testing

```bash
npm test
npm run check
```

For a manual smoke test, start `npm run dev` and upload [`tests/fixtures/valid-input.xlsx`](/Users/mzitoh/Desktop/source/delight/delight-billing-tool/tests/fixtures/valid-input.xlsx).

For a browser-level smoke test with Playwright:

```bash
npm run dev
npm run test:smoke -- data/july_raw.xlsx
```

The Worker now supports both of these input shapes:

- legacy per-person sheets from the original Streamlit app
- existing `DailyMatrix` workbooks like [`data/july_raw.xlsx`](/Users/mzitoh/Desktop/source/delight/delight-billing-tool/data/july_raw.xlsx)

## Legacy App

The original Python/Streamlit implementation is preserved in [`legacy/README.md`](/Users/mzitoh/Desktop/source/delight/delight-billing-tool/legacy/README.md).
