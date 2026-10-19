// Local preview that mimics Cloudflare Pages clean URLs: /merge-excel → site/merge-excel.html
import { createServer } from 'node:http';
import { readFile } from 'node:fs/promises';
import { join, extname, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..', 'site');
const PORT = +process.env.PORT || 8765;
const TYPES = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript', '.css': 'text/css', '.svg': 'image/svg+xml', '.png': 'image/png', '.xml': 'application/xml', '.txt': 'text/plain; charset=utf-8', '.webmanifest': 'application/manifest+json' };
createServer(async (req, res) => {
  const path = decodeURIComponent(new URL(req.url, 'http://x').pathname).replace(/\.\.+/g, '');
  for (const cand of [path, path + '.html', join(path, 'index.html')]) {
    try { const body = await readFile(join(ROOT, cand)); res.writeHead(200, { 'Content-Type': TYPES[extname(cand)] || 'application/octet-stream' }); return res.end(body); } catch {}
  }
  res.writeHead(404, { 'Content-Type': TYPES['.html'] }); res.end(await readFile(join(ROOT, '404.html')));
}).listen(PORT, () => console.log(`http://localhost:${PORT}`));
