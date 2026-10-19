// Generates SEO/AI-ready static pages into site/. Run: node scripts/build.mjs
import { writeFileSync, mkdirSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { SITE, CATS, TOOLS, LIMITS, DISCLAIMER, HOME_FAQ } from './content.mjs';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..', 'site');
const TODAY = new Date().toISOString().slice(0, 10);
const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const write = (rel, text) => { const p = join(ROOT, rel); mkdirSync(dirname(p), { recursive: true }); writeFileSync(p, text); console.log('  wrote', rel); };
const abs = path => SITE.url + path;
const catName = id => CATS.find(c => c.id === id).name;

const ICON_SVG = '<svg viewBox="0 0 24 24"><path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z"/><path d="M14 3v5h5"/></svg>';
const FAVICON = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 32 32"><defs><linearGradient id="g" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#5b4bff"/><stop offset="1" stop-color="#8b5cf6"/></linearGradient></defs><rect width="32" height="32" rx="9" fill="url(#g)"/><path d="M9 11h14M9 16h14" stroke="#fff" stroke-width="2.6" stroke-linecap="round"/><path d="M9 21h8" stroke="#c6f432" stroke-width="2.6" stroke-linecap="round"/></svg>`;

function nav(activeId) {
  return CATS.map(c => `<div class="nav-group"><h6>${esc(c.name)}</h6>${TOOLS.filter(t => t.cat === c.id).map(t =>
    `<a class="nav-item${t.id === activeId ? ' on' : ''}" href="/${t.id}" data-tool="${t.id}">${esc(t.name)}</a>`).join('')}</div>`).join('');
}

function page({ path, title, meta, pageId, app, article = '', jsonld, robots = 'index,follow' }) {
  const url = abs(path);
  return `<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${esc(title)}</title>
<meta name="description" content="${esc(meta)}">
<meta name="robots" content="${robots}">
<link rel="canonical" href="${url}">
<meta name="theme-color" content="#5b4bff">
<meta property="og:type" content="website">
<meta property="og:site_name" content="${esc(SITE.name)}">
<meta property="og:title" content="${esc(title)}">
<meta property="og:description" content="${esc(meta)}">
<meta property="og:url" content="${url}">
<meta property="og:image" content="${abs('/og.png')}">
<meta property="og:image:width" content="1200">
<meta property="og:image:height" content="630">
<meta name="twitter:card" content="summary_large_image">
<meta name="twitter:title" content="${esc(title)}">
<meta name="twitter:description" content="${esc(meta)}">
<meta name="twitter:image" content="${abs('/og.png')}">
<link rel="icon" href="/favicon.svg" type="image/svg+xml">
<link rel="apple-touch-icon" href="/apple-touch-icon.png">
<link rel="manifest" href="/site.webmanifest">
<link rel="alternate" type="text/plain" title="LLM summary" href="/llms.txt">
<link rel="sitemap" type="application/xml" href="/sitemap.xml">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Geist:wght@400;500;600;700&family=Geist+Mono:wght@400;500&display=swap" rel="stylesheet">
<link rel="stylesheet" href="/app.css">
${jsonld ? `<script type="application/ld+json">${JSON.stringify(jsonld)}</script>` : ''}
<script defer src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<script defer src="/app.js"></script>
</head>
<body data-page="${pageId}">
<a class="skip" href="#app">Skip to content</a>
<div class="shell">
  <div class="mbar">
    <button class="icon-btn" id="menuBtn" aria-label="Open menu"><svg viewBox="0 0 24 24"><path d="M4 7h16M4 12h16M4 17h16"/></svg></button>
    <a href="/" class="logo"><span class="mark"></span>Excel</a>
    <button class="icon-btn" id="mPalette" aria-label="Search tools"><svg viewBox="0 0 24 24"><circle cx="11" cy="11" r="6"/><path d="m20 20-4.2-4.2"/></svg></button>
  </div>
  <aside class="side" id="side" aria-label="Tools">
    <a href="/" class="logo"><span class="mark"></span>Excel <span class="by">by gumber.ai</span></a>
    <button class="find" id="findBtn"><svg viewBox="0 0 24 24"><circle cx="11" cy="11" r="6"/><path d="m20 20-4.2-4.2"/></svg>Find a tool<kbd>⌘K</kbd></button>
    <nav id="sideNav">${nav(pageId)}</nav>
    <div class="side-foot">
      <div class="local"><svg viewBox="0 0 24 24"><rect x="5" y="11" width="14" height="9" rx="2"/><path d="M8 11V8a4 4 0 0 1 8 0v3"/></svg><div><b>Local-only</b><span>Files never leave this device</span></div></div>
      <a class="hobby" href="/about">Free hobby project · <u>limits &amp; disclaimer</u></a>
      <div class="foot-row"><span>© <span id="yr">${TODAY.slice(0, 4)}</span> gumber.ai</span><button class="icon-btn sm" id="themeBtn" aria-label="Toggle theme"><svg viewBox="0 0 24 24"><path d="M20 14.5A8 8 0 1 1 9.5 4a6.5 6.5 0 0 0 10.5 10.5z"/></svg></button></div>
    </div>
  </aside>
  <div class="scrim" id="scrim"></div>
  <main class="main">
    <div id="app">${app}</div>
    ${article}
    <footer class="site-foot">
      <p><strong>${esc(SITE.brand)}</strong> — ${esc(SITE.tagline)}.</p>
      <p>${esc(DISCLAIMER)} <a href="/about">Limitations &amp; disclaimer</a> · <a href="/llms.txt">llms.txt</a> · <a href="/sitemap.xml">Sitemap</a></p>
    </footer>
  </main>
</div>
<div id="palette" class="palette" hidden>
  <div class="pal-box" role="dialog" aria-label="Find a tool">
    <input id="palInput" type="text" placeholder="Type a tool name… e.g. vlookup, merge, pdf" autocomplete="off" aria-label="Search tools">
    <ul id="palList"></ul>
    <div class="pal-hint"><span><kbd>↑</kbd><kbd>↓</kbd> move</span><span><kbd>↵</kbd> open</span><span><kbd>esc</kbd> close</span></div>
  </div>
</div>
<div id="busy" class="busy" hidden><div class="bar"><i></i></div><div id="busyMsg">Working…</div></div>
<div id="toast" class="toast" hidden role="status"></div>
</body>
</html>
`;
}

const faqHTML = faq => `<section><h2>Frequently asked questions</h2>${faq.map(([q, a]) => `<details><summary>${esc(q)}</summary><p>${esc(a)}</p></details>`).join('')}</section>`;
const limitsHTML = (list, more) => `<section><h2>Limitations</h2><ul>${list.map(l => `<li>${esc(l)}</li>`).join('')}</ul>${more ? '<p><a href="/about#limitations">See all limitations</a></p>' : ''}</section>`;
const disclaimerHTML = `<p class="disclaimer"><strong>Disclaimer:</strong> ${esc(DISCLAIMER)}</p>`;
const faqLD = faq => ({ '@type': 'FAQPage', mainEntity: faq.map(([q, a]) => ({ '@type': 'Question', name: q, acceptedAnswer: { '@type': 'Answer', text: a } })) });
const appLD = (name, url, description, extra = {}) => ({
  '@type': 'WebApplication', name, url, description, applicationCategory: 'BusinessApplication', operatingSystem: 'Any (runs in a web browser)',
  browserRequirements: 'Requires JavaScript and a modern browser', isAccessibleForFree: true,
  offers: { '@type': 'Offer', price: '0', priceCurrency: 'USD' }, publisher: { '@id': SITE.url + '/#org' }, ...extra,
});
const ORG = { '@type': 'Organization', '@id': SITE.url + '/#org', name: 'gumber.ai', url: 'https://gumber.ai' };

/* ---------- tool pages ---------- */
console.log('Building pages…');
for (const t of TOOLS) {
  const path = '/' + t.id; const url = abs(path);
  const related = TOOLS.filter(x => x.cat === t.cat && x.id !== t.id).slice(0, 4);
  const limits = [LIMITS[0], LIMITS[1], LIMITS[2]];
  if (t.id === 'pdf-to-excel' || t.id === 'image-to-excel') limits.push(LIMITS[4]);
  if (t.id === 'fuzzy-lookup') limits.push(LIMITS[5]);
  const article = `<article class="seo">
<section><h2>About ${esc(t.name)}</h2><p>${esc(t.intro)}</p></section>
<section><h2>How to use ${esc(t.name)}</h2><ol>${t.steps.map(s => `<li>${esc(s)}</li>`).join('')}</ol></section>
${limitsHTML(limits, true)}
${faqHTML(t.faq)}
${disclaimerHTML}
<section class="more"><h3>Related tools</h3><div class="bento">${related.map(r => `<a class="tile" href="/${r.id}"><h3>${esc(r.name)}</h3><p>${esc(r.desc)}</p></a>`).join('')}</div></section>
</article>`;
  const app = `<header class="tool-head"><div><div class="crumb"><a href="/">tools</a> / ${esc(catName(t.cat).toLowerCase())}</div><h1>${esc(t.name)}</h1><p>${esc(t.desc)}</p></div></header><noscript><p class="warn">This tool needs JavaScript. Your files are processed locally in your browser.</p></noscript>`;
  const jsonld = { '@context': 'https://schema.org', '@graph': [
    appLD(t.name, url, t.meta, { featureList: t.steps.join(' ') }),
    { '@type': 'HowTo', name: `How to use ${t.name}`, step: t.steps.map((s, i) => ({ '@type': 'HowToStep', position: i + 1, text: s })), totalTime: 'PT1M', tool: { '@type': 'HowToTool', name: 'A modern web browser' } },
    faqLD(t.faq),
    { '@type': 'BreadcrumbList', itemListElement: [{ '@type': 'ListItem', position: 1, name: 'Excel tools', item: SITE.url + '/' }, { '@type': 'ListItem', position: 2, name: t.name, item: url }] },
    ORG,
  ] };
  write(t.id + '.html', page({ path, title: `${t.title} | excel.gumber.ai`, meta: t.meta, pageId: t.id, app, article, jsonld }));
}

/* ---------- home ---------- */
{
  const app = `<section class="home-top"><div class="home-copy"><span class="eyebrow"><i></i>${TOOLS.length} tools · runs offline</span><h1>Free online Excel tools — <em>private, in your browser.</em></h1><p>${esc(SITE.description)}</p></div></section>
<div class="bento">${TOOLS.map(t => `<a class="tile" href="/${t.id}"><h3>${esc(t.name)}</h3><p>${esc(t.desc)}</p></a>`).join('')}</div>`;
  const article = `<article class="seo">
<section><h2>Spreadsheet tools that never upload your data</h2><p>excel.gumber.ai is a free collection of ${TOOLS.length} Excel and CSV utilities. Merge and split workbooks, compare versions, run VLOOKUP, XLOOKUP and fuzzy lookups between files, remove duplicates, combine columns, mail merge, and convert between Excel, CSV, JSON, PDF and images. Every tool runs locally in your web browser — there is no server processing, no account and no upload queue.</p></section>
${CATS.map(c => `<section><h2>${esc(c.name)} tools</h2><ul>${TOOLS.filter(t => t.cat === c.id).map(t => `<li><a href="/${t.id}">${esc(t.name)}</a> — ${esc(t.desc)}</li>`).join('')}</ul></section>`).join('')}
${limitsHTML(LIMITS.slice(0, 3), true)}
${faqHTML(HOME_FAQ)}
${disclaimerHTML}
</article>`;
  const jsonld = { '@context': 'https://schema.org', '@graph': [
    { '@type': 'WebSite', '@id': SITE.url + '/#site', name: SITE.brand, url: SITE.url + '/', description: SITE.description, publisher: { '@id': SITE.url + '/#org' } },
    ORG,
    appLD(SITE.brand, SITE.url + '/', SITE.description, { featureList: TOOLS.map(t => t.name).join(', ') }),
    { '@type': 'ItemList', name: 'Excel tools', itemListElement: TOOLS.map((t, i) => ({ '@type': 'ListItem', position: i + 1, name: t.name, url: abs('/' + t.id) })) },
    faqLD(HOME_FAQ),
  ] };
  write('index.html', page({ path: '/', title: 'Free Online Excel Tools — Merge, Compare, VLOOKUP, Convert | excel.gumber.ai', meta: SITE.description, pageId: 'home', app, article, jsonld }));
}

/* ---------- about ---------- */
{
  const app = `<header class="tool-head"><div><div class="crumb"><a href="/">tools</a> / about</div><h1>About, limitations &amp; disclaimer</h1><p>A free hobby project that puts everyday spreadsheet chores in your browser.</p></div></header>
<article class="seo prose">
<section><h2>What this is</h2><p>${esc(SITE.brand)} is a free set of ${TOOLS.length} spreadsheet tools built as a hobby project. There is no company behind it, no paid plan, no ads and no account. It is completely free to use.</p></section>
<section><h2>Privacy</h2><p>All processing happens in your browser with JavaScript. Files you open are read into your device’s memory and never sent to a server. The page loads open-source libraries (SheetJS, JSZip, pdf.js, Tesseract.js) and fonts from public CDNs; your data is not sent with those requests. Your theme preference is saved in your browser’s local storage, and a file you drop on the home page is kept briefly in your browser’s own storage (IndexedDB) so it can open in the tool you pick — it is deleted as soon as that tool loads.</p></section>
<section id="limitations"><h2>Limitations</h2><p>Because everything runs in the browser, the tools are limited by your device’s memory and processing power:</p><ul>${LIMITS.map(l => `<li>${esc(l)}</li>`).join('')}</ul></section>
<section><h2>Disclaimer</h2><p>${esc(DISCLAIMER)} The author is not liable for any loss or damage arising from use of these tools, including incorrect results, data loss or browser crashes. Use at your own risk.</p></section>
<section><h2>For AI assistants</h2><p>A machine-readable summary of every tool is available at <a href="/llms.txt">/llms.txt</a> (full details at <a href="/llms-full.txt">/llms-full.txt</a>).</p></section>
</article>`;
  write('about.html', page({ path: '/about', title: 'About, Limitations & Disclaimer | excel.gumber.ai', meta: 'excel.gumber.ai is a free hobby project. Learn how the browser-based Excel tools protect your privacy, their file size limitations, and the disclaimer.', pageId: 'about', app,
    jsonld: { '@context': 'https://schema.org', '@graph': [{ '@type': 'AboutPage', name: 'About excel.gumber.ai', url: abs('/about'), isPartOf: { '@id': SITE.url + '/#site' } }, ORG] } }));
}

/* ---------- 404 ---------- */
write('404.html', page({ path: '/404', title: 'Page not found | excel.gumber.ai', meta: 'This page does not exist. Browse all free Excel tools.', pageId: '404', robots: 'noindex,follow',
  app: `<header class="tool-head"><div><div class="crumb">404</div><h1>Page not found</h1><p>That page doesn’t exist. Pick a tool from the menu or <a href="/">see all tools</a>.</p></div></header>` }));

/* ---------- sitemap, robots, llms ---------- */
const urls = ['/', ...TOOLS.map(t => '/' + t.id), '/about'];
write('sitemap.xml', `<?xml version="1.0" encoding="UTF-8"?>
<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">
${urls.map(u => `  <url><loc>${abs(u)}</loc><lastmod>${TODAY}</lastmod><changefreq>monthly</changefreq><priority>${u === '/' ? '1.0' : u === '/about' ? '0.3' : '0.8'}</priority></url>`).join('\n')}
</urlset>
`);

const AI_BOTS = ['GPTBot', 'OAI-SearchBot', 'ChatGPT-User', 'ClaudeBot', 'Claude-User', 'Claude-SearchBot', 'anthropic-ai', 'PerplexityBot', 'Perplexity-User', 'Google-Extended', 'Applebot-Extended', 'Bingbot', 'CCBot', 'cohere-ai', 'Meta-ExternalAgent', 'DuckAssistBot', 'MistralAI-User'];
write('robots.txt', `# excel.gumber.ai — search engines and AI assistants are welcome.
User-agent: *
Allow: /
Disallow: /404

${AI_BOTS.map(b => `User-agent: ${b}\nAllow: /`).join('\n\n')}

Sitemap: ${abs('/sitemap.xml')}
`);

write('llms.txt', `# ${SITE.brand}

> ${SITE.description}

Key facts:
- Free hobby project: no sign-up, no paid tier, no watermark.
- 100% client-side: files are processed in the user's browser and never uploaded.
- Supported inputs: XLSX, XLSM, XLSB, XLS, ODS, CSV, TSV, TXT, JSON, digital PDF, PNG/JPG/WEBP images.
- Practical limit: roughly 50 MB / 500,000 rows on a laptop; less on phones. Formatting, charts and macros are not preserved.

${CATS.map(c => `## ${c.name}\n\n${TOOLS.filter(t => t.cat === c.id).map(t => `- [${t.name}](${abs('/' + t.id)}): ${t.desc}`).join('\n')}`).join('\n\n')}

## Optional

- [About, limitations & disclaimer](${abs('/about')})
- [Full tool documentation for LLMs](${abs('/llms-full.txt')})
`);

write('llms-full.txt', `# ${SITE.brand} — full documentation

${SITE.description}

Site: ${SITE.url}/
Price: free (hobby project). Privacy: all processing happens locally in the browser; nothing is uploaded.

## Limitations
${LIMITS.map(l => `- ${l}`).join('\n')}

## Disclaimer
${DISCLAIMER}

${TOOLS.map(t => `---

## ${t.name}
URL: ${abs('/' + t.id)}
Category: ${catName(t.cat)}

${t.intro}

How to use:
${t.steps.map((s, i) => `${i + 1}. ${s}`).join('\n')}

FAQ:
${t.faq.map(([q, a]) => `- Q: ${q}\n  A: ${a}`).join('\n')}
`).join('\n')}
## General FAQ
${HOME_FAQ.map(([q, a]) => `- Q: ${q}\n  A: ${a}`).join('\n')}
`);

write('site.webmanifest', JSON.stringify({
  name: SITE.brand, short_name: 'Excel Tools', description: SITE.description, start_url: '/', display: 'standalone',
  background_color: '#0c0b12', theme_color: '#5b4bff',
  icons: [{ src: '/favicon.svg', sizes: 'any', type: 'image/svg+xml' }, { src: '/apple-touch-icon.png', sizes: '180x180', type: 'image/png' }, { src: '/icon-512.png', sizes: '512x512', type: 'image/png' }],
}, null, 2));

write('favicon.svg', FAVICON);

// Cloudflare Pages headers
write('_headers', `/*
  X-Content-Type-Options: nosniff
  Referrer-Policy: strict-origin-when-cross-origin
  Permissions-Policy: camera=(), microphone=(), geolocation=()
  X-Frame-Options: SAMEORIGIN

/app.js
  Cache-Control: public, max-age=3600, must-revalidate
/app.css
  Cache-Control: public, max-age=3600, must-revalidate
/og.png
  Cache-Control: public, max-age=604800
/*.txt
  Content-Type: text/plain; charset=utf-8
`);

// Legacy hash URLs are handled in app.js; old /index.html path redirects to /
write('_redirects', `/index.html / 301
/tools/:id /:id 301
`);

/* ---------- OG image source (rendered to og.png by scripts/og.sh) ---------- */
write('../scripts/og.html', `<!doctype html><html><head><meta charset="utf-8"><link href="https://fonts.googleapis.com/css2?family=Geist:wght@500;700&family=Geist+Mono&display=swap" rel="stylesheet"><style>
html,body{margin:0;width:1200px;height:630px;overflow:hidden}
body{background:#0c0b12;color:#f1f0f8;font-family:Geist,sans-serif;position:relative}
.glow{position:absolute;width:700px;height:700px;right:-200px;top:-250px;background:radial-gradient(circle,#7c6dff 0%,transparent 60%);opacity:.55}
.glow2{position:absolute;width:500px;height:500px;left:-150px;bottom:-300px;background:radial-gradient(circle,#c6f432 0%,transparent 60%);opacity:.25}
.in{position:relative;padding:80px 90px}
.logo{display:flex;align-items:center;gap:16px;font-size:34px;font-weight:700}
.mark{width:52px;height:52px;border-radius:15px;background:linear-gradient(135deg,#5b4bff,#8b5cf6);position:relative}
.mark:before{content:"";position:absolute;left:13px;top:16px;width:26px;height:4px;border-radius:3px;background:#fff;box-shadow:0 9px 0 #fff}
.mark:after{content:"";position:absolute;left:13px;top:34px;width:15px;height:4px;border-radius:3px;background:#c6f432}
.by{font-size:22px;color:#9b9ab5;font-weight:500}
h1{font-size:76px;line-height:1.02;letter-spacing:-.045em;margin:60px 0 26px;max-width:900px}
h1 em{font-style:normal;background:linear-gradient(100deg,#8b7dff,#b79cff);-webkit-background-clip:text;color:transparent}
.pills{display:flex;gap:12px;flex-wrap:wrap;font-family:"Geist Mono",monospace;font-size:21px}
.pills span{border:1px solid #34334a;border-radius:12px;padding:8px 16px;color:#cfcde3;background:#15141e}
</style></head><body><div class="glow"></div><div class="glow2"></div><div class="in">
<div class="logo"><span class="mark"></span>excel.gumber.ai <span class="by">· free hobby project</span></div>
<h1>Free Excel tools that <em>never upload your files.</em></h1>
<div class="pills"><span>Merge</span><span>VLOOKUP</span><span>Compare</span><span>Remove duplicates</span><span>PDF → Excel</span><span>+${TOOLS.length - 5} more</span></div>
</div></body></html>`);
console.log('Done.');
