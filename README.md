# excel.gumber.ai

Free, private spreadsheet tools that run entirely in the browser — files are never uploaded.

Tools: Excel ↔ CSV, Excel ↔ JSON, PDF → Excel, Image → Excel (OCR), Merge Excel, CSV Merger,
Excel Splitter, Compare, Remove Duplicates, VLOOKUP / XLOOKUP / HLOOKUP / multi-criteria / fuzzy lookup,
Combine Columns, Mail Merge, Excel Viewer, Compress Excel.

## Structure

- `site/index.html` — app shell (sidebar, command palette)
- `site/app.css` — styles (light/dark tokens)
- `site/app.js` — all tools, router and UI

No build step. Libraries (SheetJS, JSZip, pdf.js, Tesseract.js) load from CDNs.

## Run locally

```bash
python3 -m http.server 8765 -d site
```

## Deploy (Cloudflare Pages)

```bash
npx wrangler pages deploy site --project-name gumber-excel
```

Then add the custom domain `excel.gumber.ai` under Workers & Pages → gumber-excel → Custom domains.
