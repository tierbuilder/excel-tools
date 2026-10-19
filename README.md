# exceltools.gumber.ai

Free, private spreadsheet tools that run entirely in the browser — files are never uploaded.

Tools: Excel ↔ CSV, Excel ↔ JSON, PDF → Excel, Image → Excel (OCR), Merge Excel, CSV Merger,
Excel Splitter, Compare, Remove Duplicates, VLOOKUP / XLOOKUP / HLOOKUP / multi-criteria / fuzzy lookup,
Combine Columns, Mail Merge, Excel Viewer, Compress Excel.

## Structure

- `scripts/content.mjs` — **edit this** for page titles, meta descriptions, how-tos, FAQs, limitations and disclaimer
- `scripts/build.mjs` — generates every page plus SEO/AI files into `site/`
- `scripts/og.sh` — renders `og.png` and app icons with headless Chrome
- `scripts/serve.mjs` — local server with Cloudflare-style clean URLs
- `site/app.js`, `site/app.css` — the tools and UI (hand-written)
- `site/*.html`, `sitemap.xml`, `robots.txt`, `llms.txt`, `llms-full.txt`, `_headers`, `_redirects`, `site.webmanifest` — **generated**, don't edit by hand

Each tool has its own URL (e.g. `/merge-excel`) with a unique title and description, canonical and Open Graph tags,
JSON-LD (WebApplication, HowTo, FAQPage, BreadcrumbList), and static how-to/FAQ/limitations content that is
readable without JavaScript. `robots.txt` explicitly allows AI crawlers and `llms.txt` summarises the site for LLMs.

No build is needed at deploy time — generated files are committed.

## Develop

```bash
node scripts/build.mjs   # after editing content.mjs (Node 18+)
./scripts/og.sh          # only if the OG image or icon changes
node scripts/serve.mjs   # http://localhost:8765
```

## Deploy (Cloudflare Pages)

Connect the GitHub repo in Cloudflare Pages with no build command and output directory `site`,
then add the custom domain `exceltools.gumber.ai`. After launch, submit `https://exceltools.gumber.ai/sitemap.xml`
in Google Search Console and Bing Webmaster Tools.

## Disclaimer

Free hobby project, provided as is without warranty. See `/about` on the site.
