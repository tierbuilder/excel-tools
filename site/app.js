/* gumber.ai Excel Tools — everything runs client-side. */
'use strict';

const LIBS = {
  jszip: 'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js',
  pdfjs: 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js',
  pdfWorker: 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js',
  tesseract: 'https://cdn.jsdelivr.net/npm/tesseract.js@5.1.1/dist/tesseract.min.js',
};
const SHEET_ACCEPT = '.xlsx,.xlsm,.xlsb,.xls,.ods,.csv,.tsv,.txt';

/* ---------- DOM helpers ---------- */
const $ = (s, r = document) => r.querySelector(s);
function h(tag, attrs, ...kids) {
  const el = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs || {})) {
    if (v == null || v === false) continue;
    if (k === 'class') el.className = v;
    else if (k === 'html') el.innerHTML = v;
    else if (k.startsWith('on')) el.addEventListener(k.slice(2), v);
    else if (k === 'value') el.value = v;
    else if (k === 'checked') el.checked = !!v;
    else el.setAttribute(k, v === true ? '' : v);
  }
  for (const kid of kids.flat(Infinity)) {
    if (kid == null || kid === false) continue;
    el.append(kid.nodeType ? kid : document.createTextNode(String(kid)));
  }
  return el;
}
const tick = () => new Promise(r => setTimeout(r, 0));
let toastTimer;
function toast(msg) {
  const t = $('#toast'); t.textContent = msg; t.hidden = false;
  clearTimeout(toastTimer); toastTimer = setTimeout(() => (t.hidden = true), 4000);
}
function busy(msg) { $('#busyMsg').textContent = msg || 'Working…'; $('#busy').hidden = false; }
function done() { $('#busy').hidden = true; }
async function guard(msg, fn) {
  busy(msg); await tick();
  try { return await fn(); } catch (e) { console.error(e); toast('Error: ' + (e.message || e)); } finally { done(); }
}
const loaded = {};
function loadScript(src) {
  return loaded[src] ||= new Promise((res, rej) => {
    const s = document.createElement('script'); s.src = src; s.onload = res;
    s.onerror = () => { delete loaded[src]; rej(new Error('Could not load ' + src)); };
    document.head.append(s);
  });
}

/* ---------- data helpers ---------- */
function str(v) {
  if (v == null) return '';
  if (v instanceof Date) {
    if (isNaN(v)) return '';
    const iso = v.toISOString();
    return iso.endsWith('T00:00:00.000Z') ? iso.slice(0, 10) : iso.replace('T', ' ').slice(0, 19);
  }
  return String(v);
}
function colLetter(i) { let s = ''; i++; while (i > 0) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; }
function widthOf(rows) { let w = 0; for (const r of rows) if (r.length > w) w = r.length; return w; }
function headersOf(rows) {
  const w = widthOf(rows.slice(0, 50)); const first = rows[0] || []; const seen = {};
  return Array.from({ length: Math.max(w, first.length) }, (_, i) => {
    let s = str(first[i]).trim() || `Column ${colLetter(i)}`;
    if (seen[s]) { seen[s]++; s = `${s} (${seen[s]})`; } else seen[s] = 1;
    return s;
  });
}
function transpose(rows) {
  const w = widthOf(rows);
  return Array.from({ length: w }, (_, c) => rows.map(r => (r[c] ?? '')));
}
function autoType(v) {
  if (typeof v !== 'string') return v;
  const t = v.trim();
  if (/^-?(0|[1-9]\d{0,14})(\.\d+)?$/.test(t)) return Number(t);
  if (/^(true|false)$/i.test(t)) return t.toLowerCase() === 'true';
  return v;
}
function detectDelim(text) {
  const lines = text.slice(0, 20000).split(/\r?\n/).filter(l => l.length).slice(0, 15);
  let best = ',', bestScore = -1;
  for (const d of [',', ';', '\t', '|']) {
    const counts = lines.map(l => l.split(d).length - 1);
    if (!counts.some(c => c > 0)) continue;
    const mode = counts[0];
    const score = counts.filter(c => c === mode && c > 0).length * 100 + mode;
    if (score > bestScore) { bestScore = score; best = d; }
  }
  return best;
}
function parseCSV(text, delim) {
  if (text.charCodeAt(0) === 0xFEFF) text = text.slice(1);
  delim = delim || detectDelim(text);
  const rows = []; let row = [], f = '', q = false, started = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (q) {
      if (c === '"') { if (text[i + 1] === '"') { f += '"'; i++; } else q = false; }
      else f += c;
    } else if (c === '"' && !started) { q = true; started = true; }
    else if (c === delim) { row.push(f); f = ''; started = false; }
    else if (c === '\n' || c === '\r') {
      if (c === '\r' && text[i + 1] === '\n') i++;
      row.push(f); rows.push(row); row = []; f = ''; started = false;
    } else { f += c; started = true; }
  }
  if (f !== '' || row.length) { row.push(f); rows.push(row); }
  return { rows, delim };
}
function toCSV(rows, delim = ',') {
  const esc = v => { const s = str(v); return /[",\r\n]/.test(s) || s.includes(delim) ? '"' + s.replace(/"/g, '""') + '"' : s; };
  return rows.map(r => r.map(esc).join(delim)).join('\r\n');
}
function readAs(file, kind) {
  return new Promise((res, rej) => {
    const r = new FileReader(); r.onload = () => res(r.result); r.onerror = () => rej(r.error);
    kind === 'text' ? r.readAsText(file) : r.readAsArrayBuffer(file);
  });
}
const isTextSheet = name => /\.(csv|tsv|txt)$/i.test(name);
const baseName = name => name.replace(/\.[^.]+$/, '');
/** Read any spreadsheet into { name, sheets: [{ name, rows }] } */
async function readBook(file, { delim } = {}) {
  if (isTextSheet(file.name)) {
    const text = await readAs(file, 'text');
    const { rows } = parseCSV(text, delim || (/\.tsv$/i.test(file.name) ? '\t' : undefined));
    return { name: file.name, sheets: [{ name: baseName(file.name).slice(0, 31) || 'Sheet1', rows: rows.map(r => r.map(autoType)) }] };
  }
  const buf = await readAs(file, 'buffer');
  const wb = XLSX.read(buf, { type: 'array', cellDates: true });
  return {
    name: file.name, wb,
    sheets: wb.SheetNames.map(n => ({ name: n, rows: XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, defval: '', raw: true, blankrows: false }) })),
  };
}
function safeSheetName(name, used) {
  let s = String(name || 'Sheet').replace(/[\[\]:*?\/\\]/g, ' ').trim().slice(0, 31) || 'Sheet';
  let base = s, n = 2;
  while (used.has(s.toLowerCase())) { const suf = ` (${n++})`; s = base.slice(0, 31 - suf.length) + suf; }
  used.add(s.toLowerCase()); return s;
}
function makeBook(sheets) {
  const wb = XLSX.utils.book_new(); const used = new Set();
  for (const s of sheets) XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(s.rows.length ? s.rows : [[]], { cellDates: true }), safeSheetName(s.name, used));
  return wb;
}
function bookBytes(sheets, type = 'xlsx') { return XLSX.write(makeBook(sheets), { bookType: type, type: 'array', compression: true }); }
function saveBlob(blob, name) {
  const a = h('a', { href: URL.createObjectURL(blob), download: name }); document.body.append(a); a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 2000);
}
function saveXlsx(sheets, name) { saveBlob(new Blob([bookBytes(sheets)], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), name); }
function saveText(text, name, type = 'text/plain', bom = false) { saveBlob(new Blob([bom ? '﻿' : '', text], { type: type + ';charset=utf-8' }), name); }
async function saveZip(files, name) {
  await loadScript(LIBS.jszip);
  const z = new JSZip(); const used = new Set();
  for (const f of files) {
    let n = f.name.replace(/[\\/:*?"<>|]/g, '_'), i = 2;
    while (used.has(n.toLowerCase())) n = f.name.replace(/(\.[^.]+)?$/, m => ` (${i++})${m}`);
    used.add(n.toLowerCase()); z.file(n, f.data);
  }
  saveBlob(await z.generateAsync({ type: 'blob', compression: 'DEFLATE' }), name);
}
function fmtBytes(n) { return n < 1024 ? n + ' B' : n < 1048576 ? (n / 1024).toFixed(1) + ' KB' : (n / 1048576).toFixed(2) + ' MB'; }
function lev(a, b) {
  if (a === b) return 0; if (!a.length) return b.length; if (!b.length) return a.length;
  let prev = Array.from({ length: b.length + 1 }, (_, i) => i), cur = new Array(b.length + 1);
  for (let i = 1; i <= a.length; i++) {
    cur[0] = i;
    for (let j = 1; j <= b.length; j++) cur[j] = Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
    [prev, cur] = [cur, prev];
  }
  return prev[b.length];
}
const similarity = (a, b) => (!a.length && !b.length) ? 1 : 1 - lev(a, b) / Math.max(a.length, b.length);

/* ---------- UI components ---------- */
function step(n, title, ...content) {
  return h('section', { class: 'step' + (n === '' ? ' res' : '') }, h('span', { class: 'dot' }, n === '' ? icon('check') : n), h('h3', null, title), h('div', { class: 'card' }, ...content));
}
function field(label, input, cls) { return h('label', { class: 'field ' + (cls || '') }, label, input); }
function check(label, checked, onchange) {
  const i = h('input', { type: 'checkbox', checked, onchange });
  const el = h('label', { class: 'check' }, i, label);
  return Object.defineProperty(el, 'on', { get: () => i.checked });
}
function select(options, value, onchange) {
  const s = h('select', { onchange }, options.map(o => Array.isArray(o) ? h('option', { value: o[0] }, o[1]) : h('option', { value: o }, o)));
  if (value != null) s.value = value; return s;
}
function colSelect(headers, value, { none } = {}) {
  const opts = headers.map((hd, i) => [String(i), hd]);
  if (none) opts.unshift(['', none]);
  return select(opts, value == null ? (none ? '' : '0') : String(value));
}
function colChooser(headers, { selected = [], ordered = false, onChange } = {}) {
  let sel = selected.slice(); const el = h('div', { class: 'chips' });
  const draw = () => el.replaceChildren(...headers.map((hd, i) => {
    const pos = sel.indexOf(i);
    return h('button', { type: 'button', class: 'chip' + (pos >= 0 ? ' on' : ''), onclick: () => { pos >= 0 ? sel.splice(pos, 1) : sel.push(i); draw(); onChange && onChange(sel); } },
      ordered && pos >= 0 ? h('span', { class: 'ord' }, pos + 1) : null, hd);
  }));
  draw();
  const wrap = h('div', null, el, h('div', { class: 'actions' },
    h('button', { type: 'button', class: 'btn ghost sm', onclick: () => { sel = headers.map((_, i) => i); draw(); onChange && onChange(sel); } }, 'Select all'),
    h('button', { type: 'button', class: 'btn ghost sm', onclick: () => { sel = []; draw(); onChange && onChange(sel); } }, 'Clear')));
  return { el: wrap, get value() { return sel.slice(); } };
}
function dropZone({ multiple = false, accept = SHEET_ACCEPT, label, hint, onFiles }) {
  const input = h('input', { type: 'file', accept, multiple, hidden: true, onchange: () => { if (input.files.length) onFiles([...input.files]); input.value = ''; } });
  const zone = h('div', { class: 'drop', tabindex: 0, role: 'button',
    onclick: () => input.click(), onkeydown: e => (e.key === 'Enter' || e.key === ' ') && input.click(),
    ondragover: e => { e.preventDefault(); zone.classList.add('over'); },
    ondragleave: () => zone.classList.remove('over'),
    ondrop: e => { e.preventDefault(); zone.classList.remove('over'); const fs = [...e.dataTransfer.files]; if (fs.length) onFiles(multiple ? fs : fs.slice(0, 1)); },
  }, h('strong', null, label || (multiple ? 'Drop files here or click to browse' : 'Drop a file here or click to browse')),
     h('span', null, hint || 'XLSX, XLS, XLSM, ODS, CSV, TSV'), input);
  zone.prepend(h('span', { class: 'dic' }, icon('upload')));
  const pend = PENDING.take(accept);
  if (pend) setTimeout(() => onFiles([pend]), 0);
  return zone;
}
/** Single spreadsheet input with sheet picker. */
function tableInput({ label, onChange } = {}) {
  const st = {
    book: null, sheet: 0,
    get rows() { return this.book ? this.book.sheets[this.sheet].rows : null; },
    get headers() { return this.rows ? headersOf(this.rows) : []; },
    get sheetName() { return this.book ? this.book.sheets[this.sheet].name : ''; },
  };
  const info = h('span', { class: 'fchip' });
  const sel = h('select', { hidden: true, onchange: () => { st.sheet = +sel.value; update(); onChange && onChange(st); } });
  const update = () => { const r = st.rows; info.textContent = `${st.book.name} · ${Math.max(0, r.length - 1).toLocaleString()} rows × ${widthOf(r)} cols`; };
  const zone = dropZone({ label, onFiles: fs => guard('Reading file…', async () => {
    st.book = await readBook(fs[0]); st.sheet = 0;
    sel.replaceChildren(...st.book.sheets.map((s, i) => h('option', { value: i }, s.name)));
    sel.hidden = st.book.sheets.length < 2; update(); onChange && onChange(st);
  }) });
  st.el = h('div', null, zone, h('div', { class: 'file-row' }, info, sel));
  return st;
}
/** Multi-file input with list. */
function filesInput({ accept, hint, onChange }) {
  const st = { files: [] };
  const list = h('ul', { class: 'file-list' });
  const draw = () => {
    list.replaceChildren(...st.files.map((f, i) => h('li', null, h('span', null, `${i + 1}. ${f.name}`, h('span', { class: 'muted' }, ` · ${fmtBytes(f.size)}`)),
      h('button', { title: 'Remove', onclick: () => { st.files.splice(i, 1); draw(); onChange && onChange(st); } }, '✕'))));
  };
  st.el = h('div', null, dropZone({ multiple: true, accept, hint, onFiles: fs => { st.files.push(...fs); draw(); onChange && onChange(st); } }), list);
  return st;
}
function renderTable(rows, { limit = 300, header = true, rowClass, cellClass } = {}) {
  const shown = rows.slice(0, limit + (header ? 1 : 0));
  const w = widthOf(shown);
  const head = header ? headersOf(rows) : Array.from({ length: w }, (_, i) => colLetter(i));
  const body = header ? shown.slice(1) : shown;
  const off = header ? 1 : 0;
  const table = h('table', { class: 'tbl' },
    h('thead', null, h('tr', null, h('th', { class: 'rn' }, '#'), head.slice(0, w || head.length).map(x => h('th', { title: x }, x)))),
    h('tbody', null, body.map((r, ri) => h('tr', { class: rowClass ? rowClass(ri + off) : null }, h('td', { class: 'rn' }, ri + 1),
      Array.from({ length: w }, (_, ci) => { const s = str(r[ci]); return h('td', { class: cellClass ? cellClass(ri + off, ci) : null, title: s.length > 40 ? s : null }, s); })))));
  const total = rows.length - off;
  return h('div', null, h('div', { class: 'tbl-wrap' }, table),
    total > body.length ? h('p', { class: 'note' }, `Showing first ${body.length.toLocaleString()} of ${total.toLocaleString()} rows. Downloads include everything.`) : null);
}
function resultPanel(title = 'Result') {
  const body = h('div');
  const el = step('', title, body); el.hidden = true;
  return {
    el,
    show(...content) { body.replaceChildren(...content); el.hidden = false; el.scrollIntoView({ behavior: 'smooth', block: 'start' }); },
    hide() { el.hidden = true; },
  };
}
const stats = list => h('div', { class: 'stats' }, list.map(([label, val, cls]) => h('div', { class: 'stat ' + (cls || '') }, h('b', null, typeof val === 'number' ? val.toLocaleString() : val), label)));
const actions = (...btns) => h('div', { class: 'actions' }, btns);
const btn = (label, onclick, ghost) => h('button', { type: 'button', class: 'btn' + (ghost ? ' ghost' : ''), onclick }, label);

/* ---------- Tools ---------- */
const TOOLS = [];
const CATS = [
  { id: 'convert', name: 'Convert', desc: 'Move data between Excel, CSV, JSON, PDF and images.' },
  { id: 'combine', name: 'Merge & Split', desc: 'Combine many files into one, or break one big file apart.' },
  { id: 'data', name: 'Clean, Compare & Lookup', desc: 'Lookups, de-duplication, diffs and column operations without formulas.' },
  { id: 'other', name: 'View & Optimize', desc: 'Open spreadsheets anywhere and shrink file sizes.' },
];
function tool(def) { TOOLS.push(def); }

/* Excel Viewer */
tool({ id: 'excel-viewer', cat: 'other', ic: 'VW', name: 'Online Excel Viewer', desc: 'Open XLSX, XLS, ODS and CSV files in your browser. Search and browse every sheet.',
  render(root) {
    const res = resultPanel('Sheet');
    const draw = st => {
      const q = h('input', { type: 'text', placeholder: 'Search in this sheet…' });
      const holder = h('div');
      const paint = () => {
        const term = q.value.trim().toLowerCase(); const rows = st.rows;
        const filtered = term ? [rows[0] || [], ...rows.slice(1).filter(r => r.some(v => str(v).toLowerCase().includes(term)))] : rows;
        holder.replaceChildren(stats([['rows', Math.max(0, filtered.length - 1)], ['columns', widthOf(rows)], ['sheets', st.book.sheets.length]]), renderTable(filtered, { limit: 1000 }));
      };
      q.addEventListener('input', paint); paint();
      res.show(h('div', { class: 'opts', style: 'margin-bottom:12px' }, field('Filter rows', q, 'grow')), holder,
        actions(btn('Download as XLSX', () => saveXlsx(st.book.sheets, baseName(st.book.name) + '.xlsx'), true),
                btn('Download sheet as CSV', () => saveText(toCSV(st.rows), `${baseName(st.book.name)}-${st.sheetName}.csv`, 'text/csv', true), true)));
    };
    const inp = tableInput({ onChange: draw });
    root.append(h('div', { class: 'flow' }, step(1, 'Open a spreadsheet', inp.el), res.el));
  } });

/* Merge Excel / CSV */
function mergeTool(root, preset) {
  const csvOnly = preset.output === 'csv';
  const files = filesInput({ accept: csvOnly ? '.csv,.tsv,.txt' : SHEET_ACCEPT, hint: csvOnly ? 'CSV, TSV or TXT files' : 'XLSX, XLS, ODS, CSV — mix and match' });
  const mode = select([['stack', 'Stack all rows into one sheet (match columns by header)'], ['sheets', 'Keep each sheet as a separate tab in one workbook']], 'stack');
  const which = select([['first', 'First sheet of each file'], ['all', 'All sheets of each file']], 'first');
  const src = check('Add a "Source" column with the file name', true);
  const out = select([['xlsx', 'Excel (.xlsx)'], ['csv', 'CSV (.csv)']], preset.output || 'xlsx');
  const res = resultPanel();
  const run = () => guard('Merging…', async () => {
    if (files.files.length < 1) return toast('Add at least two files to merge.');
    const tables = [];
    for (const f of files.files) {
      const b = await readBook(f);
      const sheets = which.value === 'all' ? b.sheets : b.sheets.slice(0, 1);
      for (const s of sheets) tables.push({ label: b.sheets.length > 1 && which.value === 'all' ? `${f.name} / ${s.name}` : f.name, sheetName: b.sheets.length > 1 ? `${baseName(f.name)}-${s.name}` : baseName(f.name), rows: s.rows });
    }
    if (mode.value === 'sheets' && !csvOnly) {
      const sheets = tables.map(t => ({ name: t.sheetName, rows: t.rows }));
      res.show(stats([['sheets', sheets.length], ['total rows', sheets.reduce((a, s) => a + Math.max(0, s.rows.length - 1), 0)]]),
        renderTable(sheets[0].rows, { limit: 100 }), h('p', { class: 'note' }, `Preview of “${sheets[0].name}”.`),
        actions(btn('Download merged workbook', () => saveXlsx(sheets, 'merged.xlsx'))));
      return;
    }
    const rows = stackTables(tables, src.on);
    const dl = out.value === 'csv' ? () => saveText(toCSV(rows), 'merged.csv', 'text/csv', true) : () => saveXlsx([{ name: 'Merged', rows }], 'merged.xlsx');
    res.show(stats([['files', files.files.length], ['rows', rows.length - 1], ['columns', rows[0].length]]), renderTable(rows), actions(btn('Download merged file', dl)));
  });
  root.append(h('div', { class: 'flow' }, step(1, 'Add files', files.el),
    step(2, 'Options', h('div', { class: 'opts' }, csvOnly ? null : field('Merge mode', mode, 'grow'), csvOnly ? null : field('Sheets', which), csvOnly ? null : field('Output', out), src), actions(btn('Merge files', run))),
    res.el));
}
function stackTables(tables, addSource) {
  const hdr = [], idx = new Map();
  for (const t of tables) for (const hd of headersOf(t.rows)) { const k = hd.trim().toLowerCase(); if (!idx.has(k)) { idx.set(k, hdr.length); hdr.push(hd.trim()); } }
  const out = [addSource ? ['Source', ...hdr] : hdr.slice()];
  for (const t of tables) {
    const map = headersOf(t.rows).map(hd => idx.get(hd.trim().toLowerCase()));
    for (const r of t.rows.slice(1)) {
      const o = new Array(hdr.length).fill('');
      r.forEach((v, i) => { if (map[i] != null) o[map[i]] = v; });
      out.push(addSource ? [t.label, ...o] : o);
    }
  }
  return out;
}
tool({ id: 'merge-excel', cat: 'combine', ic: 'MG', name: 'Merge Excel Files', desc: 'Combine multiple workbooks into one sheet or one workbook. Columns are matched by header name.', render: r => mergeTool(r, {}) });
tool({ id: 'csv-merger', cat: 'combine', ic: 'CM', name: 'CSV Merger', desc: 'Join many CSV files into a single CSV, lining up columns by header.', render: r => mergeTool(r, { output: 'csv' }) });

/* Splitter */
tool({ id: 'excel-splitter', cat: 'combine', ic: 'SP', name: 'Excel Splitter', desc: 'Split a workbook by sheet, by row count, into equal parts, or by the values in a column.',
  render(root) {
    const res = resultPanel(); const opts = h('div');
    const inp = tableInput({ onChange: st => drawOpts(st) });
    function drawOpts(st) {
      const mode = select([['rows', 'Every N rows'], ['parts', 'Into N equal parts'], ['column', 'By unique values in a column'], ['sheets', 'Each sheet into its own file']], 'rows', () => sync());
      const n = h('input', { type: 'number', min: 1, value: 1000, style: 'width:120px' });
      const col = colSelect(st.headers, 0);
      const fmt = select([['zip-xlsx', 'ZIP of .xlsx files'], ['zip-csv', 'ZIP of .csv files'], ['tabs', 'One workbook, one tab per part']], 'zip-xlsx');
      const keepHeader = check('Repeat header row in every part', true);
      const nField = field('N', n), colField = field('Column', col);
      const sync = () => { nField.hidden = !['rows', 'parts'].includes(mode.value); colField.hidden = mode.value !== 'column'; keepHeader.hidden = mode.value === 'sheets'; };
      sync();
      opts.replaceChildren(h('div', { class: 'opts' }, field('Split mode', mode), nField, colField, field('Output', fmt), keepHeader),
        actions(btn('Split file', () => guard('Splitting…', async () => {
          const parts = splitParts(st, mode.value, Math.max(1, +n.value || 1), +col.value, keepHeader.on);
          if (!parts.length) return toast('Nothing to split.');
          const base = baseName(st.book.name);
          const download = async () => {
            if (fmt.value === 'tabs') return saveXlsx(parts, `${base}-split.xlsx`);
            const csv = fmt.value === 'zip-csv';
            await saveZip(parts.map(p => ({ name: `${base}-${p.name}.${csv ? 'csv' : 'xlsx'}`, data: csv ? '﻿' + toCSV(p.rows) : bookBytes([p]) })), `${base}-split.zip`);
          };
          res.show(stats([['parts', parts.length], ['largest part (rows)', Math.max(...parts.map(p => p.rows.length))]]),
            h('div', { class: 'tbl-wrap' }, h('table', { class: 'tbl' }, h('thead', null, h('tr', null, h('th', null, 'Part'), h('th', null, 'Rows'))),
              h('tbody', null, parts.slice(0, 500).map(p => h('tr', null, h('td', null, p.name), h('td', null, p.rows.length.toLocaleString())))))),
            actions(btn('Download parts', () => guard('Packaging…', download))));
        }))));
    }
    root.append(h('div', { class: 'flow' }, step(1, 'Choose a spreadsheet', inp.el), step(2, 'How to split', opts), res.el));
    opts.append(h('p', { class: 'muted' }, 'Load a file to see options.'));
  } });
function splitParts(st, mode, n, col, keepHeader) {
  if (mode === 'sheets') return st.book.sheets.map(s => ({ name: s.name, rows: s.rows }));
  const rows = st.rows; const head = rows[0] || []; const body = rows.slice(1);
  const wrap = chunk => keepHeader ? [head, ...chunk] : chunk;
  if (mode === 'column') {
    const groups = new Map();
    for (const r of body) { const k = str(r[col]).trim() || '(blank)'; if (!groups.has(k)) groups.set(k, []); groups.get(k).push(r); }
    return [...groups].map(([k, g]) => ({ name: k.slice(0, 60), rows: wrap(g) }));
  }
  const size = mode === 'parts' ? Math.ceil(body.length / n) : n;
  const parts = [];
  for (let i = 0, p = 1; i < body.length; i += size, p++) parts.push({ name: `part-${p}`, rows: wrap(body.slice(i, i + size)) });
  return parts;
}

/* Excel -> CSV */
tool({ id: 'excel-to-csv', cat: 'convert', ic: 'CSV', name: 'Excel to CSV', desc: 'Convert XLSX/XLS sheets to CSV with your choice of delimiter. Export one sheet or all.',
  render(root) {
    const res = resultPanel(); const opts = h('div', null, h('p', { class: 'muted' }, 'Load a file to see options.'));
    const inp = tableInput({ onChange: st => {
      const delim = select([[',', 'Comma ,'], [';', 'Semicolon ;'], ['\t', 'Tab'], ['|', 'Pipe |']], ',');
      const scope = select([['one', 'Selected sheet only'], ['all', 'All sheets (ZIP)']], 'one');
      const bom = check('Add UTF-8 BOM (helps Excel open accents correctly)', true);
      opts.replaceChildren(h('div', { class: 'opts' }, field('Delimiter', delim), st.book.sheets.length > 1 ? field('Sheets', scope) : null, bom),
        actions(btn('Convert to CSV', () => {
          const text = toCSV(st.rows, delim.value);
          const base = baseName(st.book.name);
          const dl = scope.value === 'all' && st.book.sheets.length > 1
            ? () => saveZip(st.book.sheets.map(s => ({ name: `${base}-${s.name}.csv`, data: (bom.on ? '﻿' : '') + toCSV(s.rows, delim.value) })), `${base}-csv.zip`)
            : () => saveText(text, `${base}${st.book.sheets.length > 1 ? '-' + st.sheetName : ''}.csv`, 'text/csv', bom.on);
          res.show(h('pre', { class: 'out' }, text.slice(0, 20000) + (text.length > 20000 ? '\n…' : '')), actions(btn('Download CSV', dl)));
        })));
    } });
    root.append(h('div', { class: 'flow' }, step(1, 'Choose an Excel file', inp.el), step(2, 'Options', opts), res.el));
  } });

/* CSV -> Excel */
tool({ id: 'csv-to-excel', cat: 'convert', ic: 'XLS', name: 'CSV to Excel', desc: 'Turn CSV, TSV or delimited text into a proper .xlsx workbook. Auto-detects delimiters and numbers.',
  render(root) {
    const files = filesInput({ accept: '.csv,.tsv,.txt', hint: 'CSV, TSV or TXT — add several to get one tab each' });
    const delim = select([['', 'Auto-detect'], [',', 'Comma ,'], [';', 'Semicolon ;'], ['\t', 'Tab'], ['|', 'Pipe |']], '');
    const types = check('Convert numbers (keeps leading zeros as text)', true);
    const res = resultPanel();
    const run = () => guard('Converting…', async () => {
      if (!files.files.length) return toast('Add a CSV file first.');
      const sheets = [];
      for (const f of files.files) {
        const { rows, delim: d } = parseCSV(await readAs(f, 'text'), delim.value || undefined);
        sheets.push({ name: baseName(f.name), rows: types.on ? rows.map(r => r.map(autoType)) : rows, d });
      }
      const label = { ',': 'comma', ';': 'semicolon', '\t': 'tab', '|': 'pipe' };
      res.show(stats([['sheets', sheets.length], ['rows (first file)', sheets[0].rows.length - 1], ['delimiter', label[sheets[0].d] || sheets[0].d]]),
        renderTable(sheets[0].rows, { limit: 100 }), actions(btn('Download .xlsx', () => saveXlsx(sheets, (sheets.length === 1 ? sheets[0].name : 'converted') + '.xlsx'))));
    });
    root.append(h('div', { class: 'flow' }, step(1, 'Add CSV files', files.el), step(2, 'Options', h('div', { class: 'opts' }, field('Delimiter', delim), types), actions(btn('Convert to Excel', run))), res.el));
  } });

/* Excel -> JSON */
function nestKeys(obj) {
  const out = {};
  for (const [k, v] of Object.entries(obj)) {
    const parts = k.split('.'); let cur = out;
    parts.forEach((p, i) => { if (i === parts.length - 1) cur[p] = v; else cur = (cur[p] = (typeof cur[p] === 'object' && cur[p]) ? cur[p] : {}); });
  }
  return out;
}
tool({ id: 'excel-to-json', cat: 'convert', ic: '{ }', name: 'Excel to JSON', desc: 'Convert spreadsheets to JSON arrays or objects. Supports nested keys like "address.city".',
  render(root) {
    const res = resultPanel(); const opts = h('div', null, h('p', { class: 'muted' }, 'Load a file to see options.'));
    const inp = tableInput({ onChange: st => {
      const shape = select([['objects', 'Array of objects (first row = keys)'], ['arrays', 'Array of arrays']], 'objects');
      const scope = select([['one', 'Selected sheet'], ['all', 'All sheets (object keyed by sheet name)']], 'one');
      const nested = check('Nest dotted headers (a.b → {a:{b}})', true);
      const pretty = check('Pretty print', true);
      const skipEmpty = check('Omit empty cells', false);
      const conv = rows => {
        if (shape.value === 'arrays') return rows.map(r => r.map(v => v instanceof Date ? str(v) : v));
        const hd = headersOf(rows);
        return rows.slice(1).map(r => {
          let o = {};
          hd.forEach((k, i) => { const v = r[i] ?? ''; if (skipEmpty.on && v === '') return; o[k] = v instanceof Date ? str(v) : v; });
          return nested.on ? nestKeys(o) : o;
        });
      };
      opts.replaceChildren(h('div', { class: 'opts' }, field('Shape', shape), st.book.sheets.length > 1 ? field('Sheets', scope) : null, nested, skipEmpty, pretty),
        actions(btn('Convert to JSON', () => {
          const data = scope.value === 'all' ? Object.fromEntries(st.book.sheets.map(s => [s.name, conv(s.rows)])) : conv(st.rows);
          const text = JSON.stringify(data, null, pretty.on ? 2 : 0);
          res.show(h('pre', { class: 'out' }, text.slice(0, 30000) + (text.length > 30000 ? '\n…' : '')),
            actions(btn('Download .json', () => saveText(text, baseName(st.book.name) + '.json', 'application/json')),
                    btn('Copy to clipboard', () => navigator.clipboard.writeText(text).then(() => toast('Copied!')), true)));
        })));
    } });
    root.append(h('div', { class: 'flow' }, step(1, 'Choose a spreadsheet', inp.el), step(2, 'Options', opts), res.el));
  } });

/* JSON -> Excel */
function flatten(obj, prefix = '', out = {}) {
  for (const [k, v] of Object.entries(obj ?? {})) {
    const key = prefix ? `${prefix}.${k}` : k;
    if (v && typeof v === 'object' && !Array.isArray(v)) flatten(v, key, out);
    else if (Array.isArray(v)) out[key] = v.every(x => x == null || typeof x !== 'object') ? v.join(', ') : JSON.stringify(v);
    else out[key] = v;
  }
  return out;
}
function jsonToRows(arr) {
  const objs = arr.map(x => (x && typeof x === 'object' && !Array.isArray(x)) ? flatten(x) : { value: Array.isArray(x) ? JSON.stringify(x) : x });
  const keys = []; const seen = new Set();
  for (const o of objs) for (const k of Object.keys(o)) if (!seen.has(k)) { seen.add(k); keys.push(k); }
  return [keys, ...objs.map(o => keys.map(k => o[k] ?? ''))];
}
tool({ id: 'json-to-excel', cat: 'convert', ic: 'J→X', name: 'JSON to Excel', desc: 'Paste or upload JSON and get a clean spreadsheet. Nested objects are flattened into columns.',
  render(root) {
    const ta = h('textarea', { placeholder: '[{"name":"Asha","city":"Delhi"},{"name":"Ravi","city":"Pune"}]' });
    const zone = dropZone({ accept: '.json,.txt', hint: 'or drop a .json file', onFiles: async fs => { ta.value = await readAs(fs[0], 'text'); } });
    const res = resultPanel();
    const run = () => guard('Converting…', async () => {
      let data; try { data = JSON.parse(ta.value); } catch (e) { return toast('Invalid JSON: ' + e.message); }
      let sheets;
      if (Array.isArray(data)) {
        sheets = [{ name: 'Data', rows: data.every(Array.isArray) ? data : jsonToRows(data) }];
      } else if (data && typeof data === 'object') {
        const arrays = Object.entries(data).filter(([, v]) => Array.isArray(v));
        sheets = arrays.length ? arrays.map(([k, v]) => ({ name: k, rows: v.every(Array.isArray) ? v : jsonToRows(v) })) : [{ name: 'Data', rows: jsonToRows([data]) }];
      } else return toast('JSON must be an array or object.');
      res.show(stats([['sheets', sheets.length], ['rows', sheets[0].rows.length - 1], ['columns', widthOf(sheets[0].rows)]]),
        renderTable(sheets[0].rows, { limit: 100 }), actions(btn('Download .xlsx', () => saveXlsx(sheets, 'from-json.xlsx'))));
    });
    root.append(h('div', { class: 'flow' }, step(1, 'Paste JSON or upload a file', zone, h('div', { style: 'margin-top:12px' }, ta), actions(btn('Convert to Excel', run))), res.el));
  } });

/* PDF -> Excel */
tool({ id: 'pdf-to-excel', cat: 'convert', ic: 'PDF', name: 'PDF to Excel', desc: 'Extract tables and text from digital PDFs into rows and columns.',
  render(root) {
    let file = null; const info = h('span', { class: 'fchip' });
    const zone = dropZone({ accept: '.pdf', hint: 'Digital PDFs (scanned PDFs need the Image to Excel tool)', onFiles: fs => { file = fs[0]; info.textContent = `${file.name} · ${fmtBytes(file.size)}`; } });
    const layout = select([['pages', 'One sheet per page'], ['one', 'All pages in one sheet']], 'pages');
    const gap = h('input', { type: 'number', min: 0.3, max: 6, step: 0.1, value: 1.2, style: 'width:90px' });
    const align = check('Align cells into columns across rows', true);
    const res = resultPanel();
    const run = () => guard('Reading PDF…', async () => {
      if (!file) return toast('Choose a PDF first.');
      await loadScript(LIBS.pdfjs); pdfjsLib.GlobalWorkerOptions.workerSrc = LIBS.pdfWorker;
      const pdf = await pdfjsLib.getDocument({ data: await readAs(file, 'buffer') }).promise;
      const pages = [];
      for (let p = 1; p <= pdf.numPages; p++) {
        $('#busyMsg').textContent = `Reading page ${p} of ${pdf.numPages}…`;
        const tc = await (await pdf.getPage(p)).getTextContent();
        pages.push({ name: `Page ${p}`, rows: pdfRows(tc.items, +gap.value || 1.2, align.on) });
      }
      const sheets = layout.value === 'one' ? [{ name: 'PDF', rows: pages.flatMap(pg => pg.rows) }] : pages;
      if (!sheets.some(s => s.rows.length)) return res.show(h('div', { class: 'warn' }, 'No text found. This PDF is probably scanned — try the Image to Excel (OCR) tool.'));
      res.show(stats([['pages', pdf.numPages], ['rows', sheets.reduce((a, s) => a + s.rows.length, 0)]]),
        renderTable(sheets[0].rows, { header: false, limit: 200 }), h('p', { class: 'note' }, `Preview of ${sheets[0].name}.`),
        actions(btn('Download .xlsx', () => saveXlsx(sheets.map(s => ({ ...s, rows: s.rows.map(r => r.map(autoType)) })), baseName(file.name) + '.xlsx'))));
    });
    root.append(h('div', { class: 'flow' }, step(1, 'Choose a PDF', zone, h('div', { class: 'file-row' }, info)),
      step(2, 'Options', h('div', { class: 'opts' }, field('Layout', layout), field('Column gap (× font size)', gap), align), actions(btn('Extract to Excel', run))), res.el));
  } });
function pdfRows(items, gapFactor, align) {
  const its = items.filter(i => i.str && i.str.trim()).map(i => ({ s: i.str, x: i.transform[4], y: i.transform[5], w: i.width, fh: Math.abs(i.transform[3]) || i.height || 10 }));
  its.sort((a, b) => b.y - a.y || a.x - b.x);
  const lines = [];
  for (const it of its) {
    const line = lines[lines.length - 1];
    if (line && Math.abs(line.y - it.y) <= Math.max(2, it.fh * 0.45)) line.items.push(it); else lines.push({ y: it.y, items: [it] });
  }
  const cellLines = lines.map(l => {
    l.items.sort((a, b) => a.x - b.x);
    const cells = [];
    for (const it of l.items) {
      const c = cells[cells.length - 1];
      const g = c ? it.x - c.end : Infinity;
      if (c && g <= gapFactor * it.fh) { c.t += (g > it.fh * 0.12 && !c.t.endsWith(' ') && !it.s.startsWith(' ') ? ' ' : '') + it.s; c.end = Math.max(c.end, it.x + it.w); }
      else cells.push({ x: it.x, end: it.x + it.w, t: it.s });
    }
    return cells;
  });
  if (!align) return cellLines.map(cs => cs.map(c => c.t.trim()));
  // cluster cell start positions into column anchors
  const xs = cellLines.flat().map(c => c.x).sort((a, b) => a - b);
  const anchors = [];
  for (const x of xs) { if (!anchors.length || x - anchors[anchors.length - 1].max > 6) anchors.push({ min: x, max: x }); else anchors[anchors.length - 1].max = x; }
  return cellLines.map(cs => {
    const row = new Array(anchors.length).fill('');
    for (const c of cs) {
      let k = anchors.findIndex(a => c.x >= a.min - 0.5 && c.x <= a.max + 0.5); if (k < 0) k = 0;
      row[k] = row[k] ? row[k] + ' ' + c.t.trim() : c.t.trim();
    }
    while (row.length && row[row.length - 1] === '') row.pop();
    return row;
  });
}

/* Image -> Excel (OCR) */
tool({ id: 'image-to-excel', cat: 'convert', ic: 'OCR', name: 'Image to Excel (OCR)', desc: 'Read tables from screenshots, photos or scans and turn them into a spreadsheet.',
  render(root) {
    let file = null; const info = h('span', { class: 'fchip' }); const preview = h('div');
    const zone = dropZone({ accept: 'image/*', hint: 'PNG, JPG, WEBP, BMP', onFiles: fs => {
      file = fs[0]; info.textContent = `${file.name} · ${fmtBytes(file.size)}`;
      preview.replaceChildren(h('img', { src: URL.createObjectURL(file), alt: 'Selected image', style: 'max-width:100%;max-height:260px;margin-top:12px;border-radius:10px;border:1px solid var(--border)' }));
    } });
    const lang = select([['eng', 'English'], ['hin', 'Hindi'], ['spa', 'Spanish'], ['fra', 'French'], ['deu', 'German'], ['por', 'Portuguese'], ['ita', 'Italian'], ['chi_sim', 'Chinese (Simplified)'], ['jpn', 'Japanese'], ['ara', 'Arabic']], 'eng');
    const split = select([['spaces', '2+ spaces, tabs or | separate columns'], ['single', 'Any space separates columns'], ['none', 'One column per line']], 'spaces');
    const res = resultPanel();
    const run = () => guard('Loading OCR engine…', async () => {
      if (!file) return toast('Choose an image first.');
      await loadScript(LIBS.tesseract);
      const worker = await Tesseract.createWorker(lang.value, 1, { logger: m => { if (m.status) $('#busyMsg').textContent = `${m.status}${m.progress ? ' ' + Math.round(m.progress * 100) + '%' : ''}`; } });
      const { data } = await worker.recognize(file); await worker.terminate();
      const re = split.value === 'spaces' ? /\s*\|\s*|\t+| {2,}/ : split.value === 'single' ? /\s+/ : null;
      const rows = data.text.split('\n').map(l => l.trim()).filter(Boolean).map(l => (re ? l.split(re) : [l]).map(c => c.trim()).filter((c, i, a) => c || i < a.length - 1));
      res.show(stats([['lines', rows.length], ['confidence', Math.round(data.confidence) + '%']]), renderTable(rows, { header: false }),
        h('p', { class: 'note' }, 'OCR is never perfect — check numbers before relying on them.'),
        actions(btn('Download .xlsx', () => saveXlsx([{ name: 'OCR', rows: rows.map(r => r.map(autoType)) }], baseName(file.name) + '.xlsx')),
                btn('Download .csv', () => saveText(toCSV(rows), baseName(file.name) + '.csv', 'text/csv', true), true)));
    });
    root.append(h('div', { class: 'flow' }, step(1, 'Choose an image', zone, h('div', { class: 'file-row' }, info), preview),
      step(2, 'Options', h('div', { class: 'opts' }, field('Language', lang), field('Columns', split)), actions(btn('Extract table', run))), res.el));
  } });

/* Compare */
tool({ id: 'compare-excel', cat: 'data', ic: 'Δ', name: 'Compare Excel Files', desc: 'Spot changed cells, added rows and removed rows between two versions of a sheet.',
  render(root) {
    const res = resultPanel('Differences'); const opts = h('div', null, h('p', { class: 'muted' }, 'Load both files to see options.'));
    const a = tableInput({ label: 'Original file (A)', onChange: () => drawOpts() });
    const b = tableInput({ label: 'New file (B)', onChange: () => drawOpts() });
    function drawOpts() {
      if (!a.rows || !b.rows) return;
      const hA = a.headers, hB = new Set(b.headers.map(x => x.toLowerCase()));
      const common = hA.map((x, i) => [String(i), x]).filter(([, x]) => hB.has(x.toLowerCase()));
      const key = select([['', 'Row position (row 1 vs row 1)'], ...common], '');
      const ic = check('Ignore case', false), tr = check('Ignore leading/trailing spaces', true), hideSame = check('Hide unchanged rows', true);
      opts.replaceChildren(h('div', { class: 'opts' }, field('Match rows by', key, 'grow'), ic, tr, hideSame),
        actions(btn('Compare', () => guard('Comparing…', async () => compare(key.value === '' ? null : hA[+key.value], { ic: ic.on, tr: tr.on, hideSame: hideSame.on })))));
    }
    function compare(keyName, o) {
      const norm = v => { let s = str(v); if (o.tr) s = s.trim(); if (o.ic) s = s.toLowerCase(); return s; };
      const hA = a.headers, hB = b.headers;
      const cols = [...hA]; const lower = new Set(hA.map(x => x.toLowerCase()));
      for (const x of hB) if (!lower.has(x.toLowerCase())) cols.push(x);
      const ia = cols.map(c => hA.findIndex(x => x.toLowerCase() === c.toLowerCase()));
      const ib = cols.map(c => hB.findIndex(x => x.toLowerCase() === c.toLowerCase()));
      const pick = (r, map) => map.map(i => i < 0 ? '' : (r[i] ?? ''));
      const A = a.rows.slice(1), B = b.rows.slice(1);
      let pairs = [];
      if (keyName) {
        const ka = hA.findIndex(x => x === keyName), kb = hB.findIndex(x => x.toLowerCase() === keyName.toLowerCase());
        const mapB = new Map(); B.forEach((r, i) => { const k = norm(r[kb]); if (!mapB.has(k)) mapB.set(k, i); });
        const usedB = new Set();
        A.forEach(r => { const k = norm(r[ka]); const j = mapB.get(k); if (j != null && !usedB.has(j)) { usedB.add(j); pairs.push([r, B[j]]); } else pairs.push([r, null]); });
        B.forEach((r, j) => { if (!usedB.has(j)) pairs.push([null, r]); });
      } else {
        for (let i = 0; i < Math.max(A.length, B.length); i++) pairs.push([A[i] || null, B[i] || null]);
      }
      const out = [['Status', ...cols]], marks = [null], changes = [['Row', keyName || 'Row #', 'Column', 'Old value (A)', 'New value (B)']];
      let added = 0, removed = 0, changed = 0, same = 0;
      pairs.forEach(([ra, rb], n) => {
        if (ra && !rb) { removed++; out.push(['Removed', ...pick(ra, ia)]); marks.push({ row: 'del' }); return; }
        if (!ra && rb) { added++; out.push(['Added', ...pick(rb, ib)]); marks.push({ row: 'add' }); return; }
        const va = pick(ra, ia), vb = pick(rb, ib); const diffCols = new Set();
        cols.forEach((c, i) => { if (norm(va[i]) !== norm(vb[i])) { diffCols.add(i + 1); changes.push([n + 2, keyName ? str(ra[hA.indexOf(keyName)]) : n + 1, c, va[i], vb[i]]); } });
        if (diffCols.size) { changed++; out.push(['Changed', ...cols.map((_, i) => diffCols.has(i + 1) ? `${str(va[i])} → ${str(vb[i])}` : vb[i])]); marks.push({ cells: diffCols }); }
        else { same++; if (!o.hideSame) { out.push(['Same', ...vb]); marks.push({}); } }
      });
      res.show(stats([['changed rows', changed, 'chg'], ['added rows', added, 'add'], ['removed rows', removed, 'del'], ['unchanged', same], ['changed cells', changes.length - 1, 'chg']]),
        out.length > 1 ? renderTable(out, { limit: 500, rowClass: r => marks[r]?.row, cellClass: (r, c) => marks[r]?.cells?.has(c) ? 'chg' : null }) : h('p', null, '✅ The files are identical with these settings.'),
        actions(btn('Download diff report (.xlsx)', () => saveXlsx([{ name: 'Summary', rows: out }, { name: 'Cell changes', rows: changes }], 'comparison.xlsx'))));
    }
    root.append(h('div', { class: 'flow' }, step(1, 'Add both versions', h('div', { class: 'pairgrid' }, h('div', null, h('p', { class: 'sub' }, h('b', null, 'A'), 'Original'), a.el), h('div', null, h('p', { class: 'sub' }, h('b', null, 'B'), 'New version'), b.el))), step(2, 'Options', opts), res.el));
  } });

/* Remove duplicates */
tool({ id: 'remove-duplicates', cat: 'data', ic: 'DUP', name: 'Remove Duplicates', desc: 'Find and remove duplicate rows based on all columns or just the ones you pick.',
  render(root) {
    const res = resultPanel(); const opts = h('div', null, h('p', { class: 'muted' }, 'Load a file to see options.'));
    const inp = tableInput({ onChange: st => {
      const cols = colChooser(st.headers, { selected: st.headers.map((_, i) => i) });
      const ic = check('Ignore case', true), tr = check('Ignore extra spaces', true);
      const keep = select([['first', 'Keep first occurrence'], ['last', 'Keep last occurrence']], 'first');
      opts.replaceChildren(h('p', { class: 'note', style: 'margin-top:0' }, 'Rows are duplicates when these columns match:'), cols.el,
        h('div', { class: 'opts', style: 'margin-top:14px' }, field('Which to keep', keep), ic, tr),
        actions(btn('Remove duplicates', () => guard('Finding duplicates…', async () => {
          const c = cols.value; if (!c.length) return toast('Pick at least one column.');
          const norm = v => { let s = str(v); if (tr.on) s = s.trim().replace(/\s+/g, ' '); if (ic.on) s = s.toLowerCase(); return s; };
          const body = st.rows.slice(1); const order = keep.value === 'last' ? [...body.keys()].reverse() : [...body.keys()];
          const seen = new Set(), keepIdx = new Set(), dupIdx = [];
          for (const i of order) { const k = c.map(j => norm(body[i][j])).join('\u0001'); if (seen.has(k)) dupIdx.push(i); else { seen.add(k); keepIdx.add(i); } }
          const head = st.rows[0] || [];
          const clean = [head, ...body.filter((_, i) => keepIdx.has(i))];
          const dups = [head, ...dupIdx.sort((x, y) => x - y).map(i => body[i])];
          const base = baseName(st.book.name);
          res.show(stats([['rows before', body.length], ['duplicates removed', dupIdx.length, 'del'], ['rows after', clean.length - 1, 'add']]), renderTable(clean),
            actions(btn('Download cleaned file', () => saveXlsx([{ name: st.sheetName, rows: clean }], `${base}-deduped.xlsx`)),
                    dupIdx.length ? btn('Download removed duplicates', () => saveXlsx([{ name: 'Duplicates', rows: dups }], `${base}-duplicates.xlsx`), true) : null,
                    btn('Download as CSV', () => saveText(toCSV(clean), `${base}-deduped.csv`, 'text/csv', true), true)));
        }))));
    } });
    root.append(h('div', { class: 'flow' }, step(1, 'Choose a spreadsheet', inp.el), step(2, 'Duplicate rules', opts), res.el));
  } });

/* Lookups (VLOOKUP / XLOOKUP / HLOOKUP / INDEX-MATCH / multi-criteria / fuzzy) */
function lookupTool(root, preset) {
  const res = resultPanel(); const opts = h('div', null, h('p', { class: 'muted' }, 'Load both tables to see options.'));
  const a = tableInput({ label: 'Main table (rows you want to fill in)', onChange: () => drawOpts() });
  const b = tableInput({ label: 'Lookup table (where the answers live)', onChange: () => drawOpts() });
  let orient = preset.orient || 'cols';
  function drawOpts() {
    if (!a.rows || !b.rows) return;
    const A = orient === 'rows' ? transpose(a.rows) : a.rows, B = orient === 'rows' ? transpose(b.rows) : b.rows;
    const hA = headersOf(A), hB = headersOf(B);
    const orientSel = select([['cols', 'Vertical — headers in first row (VLOOKUP/XLOOKUP)'], ['rows', 'Horizontal — headers in first column (HLOOKUP)']], orient, () => { orient = orientSel.value; drawOpts(); });
    const pairs = h('div', { class: 'pairs' }); const pairSels = [];
    const addPair = () => {
      const guess = hB.findIndex(x => x.toLowerCase() === (hA[pairSels.length] || hA[0]).toLowerCase());
      const sa = colSelect(hA, Math.min(pairSels.length, hA.length - 1)); const sb = colSelect(hB, guess >= 0 ? guess : 0);
      const pair = { sa, sb }; pairSels.push(pair);
      const row = h('div', { class: 'pair' }, h('span', { class: 'muted' }, pairSels.length === 1 ? 'Match' : 'and'), sa, h('span', { class: 'muted' }, 'with'), sb,
        pairSels.length > 1 ? h('button', { type: 'button', class: 'btn ghost sm', onclick: () => { pairSels.splice(pairSels.indexOf(pair), 1); row.remove(); } }, 'Remove') : null);
      pairs.append(row);
    };
    // try to auto-match a common column first
    addPair(); for (let i = 1; i < (preset.criteria || 1); i++) addPair();
    const common = hA.findIndex(x => hB.some(y => y.toLowerCase() === x.toLowerCase()));
    if (common >= 0) { pairSels[0].sa.value = String(common); pairSels[0].sb.value = String(hB.findIndex(y => y.toLowerCase() === hA[common].toLowerCase())); }
    const ret = colChooser(hB, { selected: hB.length > 1 ? [hB.length - 1 === +pairSels[0].sb.value ? 0 : hB.length - 1] : [0] });
    const mode = select([['exact', 'Exact match'], ['fuzzy', 'Fuzzy match (similar text)']], preset.fuzzy ? 'fuzzy' : 'exact', () => (thField.hidden = mode.value !== 'fuzzy'));
    const th = h('input', { type: 'number', min: 50, max: 100, value: 80, style: 'width:90px' }); const thField = field('Min. similarity %', th); thField.hidden = mode.value !== 'fuzzy';
    const multi = select([['first', 'First match'], ['last', 'Last match'], ['all', 'All matches (comma-joined)'], ['count', 'Count of matches']], 'first');
    const nf = h('input', { type: 'text', value: '#N/A', style: 'width:110px' });
    const ic = check('Ignore case & extra spaces', true);
    opts.replaceChildren(
      h('div', { class: 'opts' }, field('Table orientation', orientSel, 'grow')),
      h('p', { class: 'note' }, 'Match criteria (main table column ↔ lookup table column):'), pairs,
      actions(btn('+ Add criterion (multi-criteria)', addPair, true)),
      h('p', { class: 'note' }, 'Columns to bring back from the lookup table:'), ret.el,
      h('div', { class: 'opts', style: 'margin-top:14px' }, field('Match type', mode), thField, field('When several match', multi), field('If not found', nf), ic),
      actions(btn('Run lookup', () => guard('Looking up…', () => runLookup(A, B, hA, hB, {
        pairs: pairSels.map(p => [+p.sa.value, +p.sb.value]), ret: ret.value, fuzzy: mode.value === 'fuzzy', th: (+th.value || 80) / 100, multi: multi.value, nf: nf.value, ic: ic.on,
      })))));
  }
  async function runLookup(A, B, hA, hB, o) {
    if (!o.ret.length) return toast('Pick at least one column to return.');
    const norm = v => { let s = str(v); if (o.ic) s = s.trim().replace(/\s+/g, ' ').toLowerCase(); return s; };
    const keyOf = (r, side) => o.pairs.map(p => norm(r[p[side]])).join(' | ');
    const index = new Map();
    B.slice(1).forEach(r => { const k = keyOf(r, 1); if (!index.has(k)) index.set(k, []); index.get(k).push(r); });
    const keys = [...index.keys()];
    const outHead = [...hA, ...o.ret.map(i => hB[i])]; if (o.fuzzy) outHead.push('Match score', 'Matched value');
    const out = [outHead]; let found = 0, fuzzyFound = 0;
    const body = A.slice(1);
    for (let n = 0; n < body.length; n++) {
      const r = body[n]; const k = keyOf(r, 0); let hits = index.get(k); let score = hits ? 1 : 0, matched = hits ? k : '';
      if (!hits && o.fuzzy && k.replace(/ \| /g, '')) {
        let best = null, bs = 0;
        for (const cand of keys) { if (Math.abs(cand.length - k.length) / Math.max(cand.length, k.length, 1) > 1 - o.th + 1e-9) continue; const s = similarity(k, cand); if (s > bs) { bs = s; best = cand; } }
        if (best != null && bs >= o.th - 1e-9) { hits = index.get(best); score = bs; matched = best; fuzzyFound++; }
      }
      if (n % 300 === 299) { $('#busyMsg').textContent = `Looking up… ${n + 1} / ${body.length}`; await tick(); }
      const vals = o.ret.map(i => {
        if (!hits) return o.multi === 'count' ? 0 : o.nf;
        if (o.multi === 'count') return hits.length;
        if (o.multi === 'all') return [...new Set(hits.map(x => str(x[i])))].join(', ');
        return (o.multi === 'last' ? hits[hits.length - 1] : hits[0])[i] ?? '';
      });
      if (hits) found++;
      const row = [...Array.from({ length: hA.length }, (_, i) => r[i] ?? ''), ...vals];
      if (o.fuzzy) row.push(hits ? Math.round(score * 100) + '%' : '', matched);
      out.push(row);
    }
    const final = orient === 'rows' ? transpose(out) : out;
    const name = baseName(a.book.name);
    res.show(stats([['rows', body.length], ['matched', found, 'add'], ...(o.fuzzy ? [['fuzzy matches', fuzzyFound, 'chg']] : []), ['not found', body.length - found, 'del']]),
      renderTable(final, { header: orient !== 'rows' }),
      actions(btn('Download result (.xlsx)', () => saveXlsx([{ name: 'Lookup result', rows: final }], `${name}-lookup.xlsx`)),
              btn('Download as CSV', () => saveText(toCSV(final), `${name}-lookup.csv`, 'text/csv', true), true)));
  }
  root.append(h('div', { class: 'flow' }, step(1, 'Add your two tables', h('div', { class: 'pairgrid' }, h('div', null, h('p', { class: 'sub' }, h('b', null, 'A'), 'Main table'), a.el), h('div', null, h('p', { class: 'sub' }, h('b', null, 'B'), 'Lookup table'), b.el)),
    h('p', { class: 'note' }, 'Tip: both tables in one workbook? Upload it twice and pick a different sheet on each side.')),
    step(2, 'Lookup settings', opts), res.el));
}
tool({ id: 'vlookup', cat: 'data', ic: 'VL', name: 'Excel VLOOKUP', desc: 'Pull matching values from a second table into your sheet — no formulas, no #REF errors.', render: r => lookupTool(r, {}) });
tool({ id: 'xlookup', cat: 'data', ic: 'XL', name: 'XLOOKUP / INDEX-MATCH', desc: 'Look up any column in any direction and return one or many result columns.', render: r => lookupTool(r, {}) });
tool({ id: 'hlookup', cat: 'data', ic: 'HL', name: 'HLOOKUP Tool', desc: 'Lookups for horizontal tables where headers run down the first column.', render: r => lookupTool(r, { orient: 'rows' }) });
tool({ id: 'multi-criteria-lookup', cat: 'data', ic: 'MC', name: 'Multi-Criteria Lookup', desc: 'Match rows on two or more columns at once, e.g. Name + Date.', render: r => lookupTool(r, { criteria: 2 }) });
tool({ id: 'fuzzy-lookup', cat: 'data', ic: '≈', name: 'Excel Fuzzy Lookup', desc: 'Match names that are almost the same — typos, abbreviations and spacing differences.', render: r => lookupTool(r, { fuzzy: true }) });

/* Combine columns */
tool({ id: 'combine-columns', cat: 'data', ic: 'A+B', name: 'Combine Columns', desc: 'Join two or more columns into one with any separator — e.g. First + Last name.',
  render(root) {
    const res = resultPanel(); const opts = h('div', null, h('p', { class: 'muted' }, 'Load a file to see options.'));
    const inp = tableInput({ onChange: st => {
      const cols = colChooser(st.headers, { ordered: true });
      const sep = h('input', { type: 'text', value: ' ', style: 'width:90px' });
      const presets = h('div', { class: 'chips' }, [[' ', 'Space'], [', ', 'Comma'], [' - ', 'Dash'], ['', 'None'], ['\n', 'New line']].map(([v, l]) => h('button', { type: 'button', class: 'chip', onclick: () => (sep.value = v) }, l)));
      const name = h('input', { type: 'text', value: 'Combined' });
      const skip = check('Skip empty cells', true), remove = check('Remove original columns', false);
      opts.replaceChildren(h('p', { class: 'note', style: 'margin-top:0' }, 'Click columns in the order to combine them:'), cols.el,
        h('div', { class: 'opts', style: 'margin-top:14px' }, field('Separator', sep), field('Quick pick', presets), field('New column name', name), skip, remove),
        actions(btn('Combine columns', () => {
          const c = cols.value; if (c.length < 2) return toast('Pick at least two columns.');
          const s = sep.value.replace(/\\n/g, '\n');
          const w = st.headers.length; const drop = new Set(remove.on ? c : []);
          const out = st.rows.map((r, ri) => {
            const base = Array.from({ length: w }, (_, i) => r[i] ?? '').filter((_, i) => !drop.has(i));
            if (ri === 0) return [...st.headers.filter((_, i) => !drop.has(i)), name.value || 'Combined'];
            const parts = c.map(i => str(r[i])); return [...base, (skip.on ? parts.filter(p => p.trim() !== '') : parts).join(s)];
          });
          res.show(renderTable(out), actions(btn('Download .xlsx', () => saveXlsx([{ name: st.sheetName, rows: out }], baseName(st.book.name) + '-combined.xlsx')),
            btn('Download .csv', () => saveText(toCSV(out), baseName(st.book.name) + '-combined.csv', 'text/csv', true), true)));
        })));
    } });
    root.append(h('div', { class: 'flow' }, step(1, 'Choose a spreadsheet', inp.el), step(2, 'Columns to combine', opts), res.el));
  } });

/* Mail merge */
tool({ id: 'mail-merge', cat: 'data', ic: '✉', name: 'Mail Merge', desc: 'Generate personalised letters, emails or messages from a spreadsheet and a template.',
  render(root) {
    const res = resultPanel(); const opts = h('div', null, h('p', { class: 'muted' }, 'Load a data file to see options.'));
    const inp = tableInput({ label: 'Drop your data (one row per recipient)', onChange: st => {
      const hd = st.headers;
      const ta = h('textarea', { style: 'min-height:200px', value: `Dear {{${hd.find(x => /name/i.test(x)) || hd[0]}}},\n\nThank you for being with us.\n\nBest regards,\nTeam` });
      const insert = t => { const s = ta.selectionStart ?? ta.value.length; ta.value = ta.value.slice(0, s) + t + ta.value.slice(ta.selectionEnd ?? s); ta.focus(); ta.selectionStart = ta.selectionEnd = s + t.length; };
      const fname = h('input', { type: 'text', value: `{{${hd[0]}}}` });
      const ext = select([['txt', '.txt'], ['html', '.html'], ['md', '.md']], 'txt');
      const outMode = select([['zip', 'One file per row (ZIP)'], ['one', 'Everything in one file'], ['xlsx', 'Spreadsheet with a "Message" column']], 'zip');
      const fill = r => tpl => tpl.replace(/\{\{\s*([^}]+?)\s*\}\}/g, (m, k) => { const i = hd.findIndex(x => x.toLowerCase() === k.toLowerCase()); return i >= 0 ? str(r[i]) : m; });
      opts.replaceChildren(
        h('p', { class: 'note', style: 'margin-top:0' }, 'Click a field to insert it at the cursor:'),
        h('div', { class: 'chips', style: 'margin-bottom:10px' }, hd.map(x => h('button', { type: 'button', class: 'chip', onclick: () => insert(`{{${x}}}`) }, x))), ta,
        h('div', { class: 'opts', style: 'margin-top:14px' }, field('Output', outMode), field('File name pattern', fname), field('Extension', ext)),
        actions(btn('Preview & generate', () => {
          const body = st.rows.slice(1).filter(r => r.some(v => str(v) !== ''));
          const docs = body.map(r => ({ name: fill(r)(fname.value) || 'document', text: fill(r)(ta.value) }));
          const download = () => guard('Generating…', async () => {
            if (outMode.value === 'zip') return saveZip(docs.map(d => ({ name: `${d.name}.${ext.value}`, data: d.text })), 'mail-merge.zip');
            if (outMode.value === 'one') return saveText(docs.map(d => d.text).join('\n\n' + '-'.repeat(40) + '\n\n'), `mail-merge.${ext.value}`);
            saveXlsx([{ name: 'Merged', rows: [[...hd, 'Message'], ...body.map((r, i) => [...hd.map((_, j) => r[j] ?? ''), docs[i].text])] }], 'mail-merge.xlsx');
          });
          res.show(stats([['documents', docs.length]]),
            ...docs.slice(0, 3).map(d => h('div', { style: 'margin-bottom:12px' }, h('div', { class: 'note' }, `${d.name}.${ext.value}`), h('pre', { class: 'out' }, d.text))),
            docs.length > 3 ? h('p', { class: 'note' }, `…and ${docs.length - 3} more.`) : null, actions(btn('Download', download)));
        })));
    } });
    root.append(h('div', { class: 'flow' }, step(1, 'Data source', inp.el), step(2, 'Template', opts), res.el));
  } });

/* Compress */
tool({ id: 'compress-excel', cat: 'other', ic: 'ZIP', name: 'Compress Excel', desc: 'Shrink bloated XLSX/XLS files by rewriting them with maximum compression.',
  render(root) {
    const files = filesInput({ accept: '.xlsx,.xlsm,.xls,.xlsb,.ods', hint: 'XLSX, XLSM, XLS, XLSB, ODS' });
    const res = resultPanel();
    const run = () => guard('Compressing…', async () => {
      if (!files.files.length) return toast('Add a file first.');
      const results = [];
      for (const f of files.files) {
        const wb = XLSX.read(await readAs(f, 'buffer'), { type: 'array', cellDates: true });
        const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array', compression: true });
        results.push({ f, out, before: f.size, after: out.byteLength });
      }
      const name = r => baseName(r.f.name) + '-compressed.xlsx';
      res.show(h('div', { class: 'warn' }, 'Compression keeps your data and formulas but drops formatting, charts, images and macros. Keep your original.'),
        h('div', { class: 'tbl-wrap' }, h('table', { class: 'tbl' }, h('thead', null, h('tr', null, ['File', 'Before', 'After', 'Saved', ''].map(x => h('th', null, x)))),
          h('tbody', null, results.map(r => h('tr', null, h('td', null, r.f.name), h('td', null, fmtBytes(r.before)), h('td', null, fmtBytes(r.after)),
            h('td', null, r.after < r.before ? Math.round((1 - r.after / r.before) * 100) + '%' : 'Already optimal'),
            h('td', null, h('button', { class: 'btn sm', onclick: () => saveBlob(new Blob([r.out]), name(r)) }, 'Download'))))))),
        results.length > 1 ? actions(btn('Download all (ZIP)', () => saveZip(results.map(r => ({ name: name(r), data: r.out })), 'compressed.zip'))) : null);
    });
    root.append(h('div', { class: 'flow' }, step(1, 'Add workbooks', files.el, actions(btn('Compress', run))), res.el));
  } });

/* ---------- icons ---------- */
const ICONS = {
  upload: '<path d="M12 16V4m0 0-4.5 4.5M12 4l4.5 4.5"/><path d="M4 15v3a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-3"/>',
  check: '<path d="m5 12.5 4.5 4.5L19 7.5"/>',
  arrow: '<path d="M7 17 17 7M9 7h8v8"/>',
  lock: '<rect x="5" y="11" width="14" height="9" rx="2"/><path d="M8 11V8a4 4 0 0 1 8 0v3"/>',
  bolt: '<path d="M13 3 5 13.5h6L10 21l8-10.5h-6z"/>',
  user: '<circle cx="12" cy="8" r="3.5"/><path d="M5 20a7 7 0 0 1 14 0"/>',
  eye: '<path d="M2.5 12S6 5.5 12 5.5 21.5 12 21.5 12 18 18.5 12 18.5 2.5 12 2.5 12z"/><circle cx="12" cy="12" r="2.8"/>',
  merge: '<path d="M6 4v5a5 5 0 0 0 5 5h7m0 0-3.5-3.5M18 14l-3.5 3.5"/><path d="M6 20v-6"/>',
  stack: '<rect x="4" y="4" width="16" height="5" rx="1.5"/><rect x="4" y="10.5" width="16" height="4" rx="1.5"/><rect x="4" y="16" width="16" height="4" rx="1.5"/>',
  split: '<rect x="3.5" y="4" width="7" height="16" rx="1.5"/><rect x="13.5" y="4" width="7" height="16" rx="1.5"/>',
  csv: '<path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z"/><path d="M14 3v5h5M8.5 13h7M8.5 16.5h7M12 11v8"/>',
  table: '<rect x="3.5" y="4.5" width="17" height="15" rx="2"/><path d="M3.5 9.5h17M3.5 14.5h17M9.5 4.5v15"/>',
  braces: '<path d="M8 4H7a2 2 0 0 0-2 2v3.5L3.5 12 5 14.5V18a2 2 0 0 0 2 2h1M16 4h1a2 2 0 0 1 2 2v3.5l1.5 2.5-1.5 2.5V18a2 2 0 0 1-2 2h-1"/>',
  json2: '<path d="M4 7h10m0 0-3-3m3 3-3 3"/><rect x="9" y="12" width="11" height="8" rx="1.5"/><path d="M9 16h11M14.5 12v8"/>',
  pdf: '<path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z"/><path d="M14 3v5h5"/><path d="M8.5 17v-4h1.5a1.25 1.25 0 0 1 0 2.5H8.5M13 13v4h1a2 2 0 0 0 0-4z"/>',
  image: '<rect x="3.5" y="4.5" width="17" height="15" rx="2"/><circle cx="9" cy="10" r="1.7"/><path d="m20.5 16-4.5-4.5L6 19.5"/>',
  diff: '<path d="M7 4v10M2 9h10M14 18h8M17 5v6"/><circle cx="17" cy="15" r="0" />',
  dedupe: '<rect x="8" y="8" width="12" height="12" rx="2"/><path d="M16 8V6a2 2 0 0 0-2-2H6a2 2 0 0 0-2 2v8a2 2 0 0 0 2 2h2"/><path d="m11 14 2 2 4-4"/>',
  search: '<circle cx="11" cy="11" r="6"/><path d="m20 20-4.2-4.2"/>',
  target: '<circle cx="12" cy="12" r="8"/><circle cx="12" cy="12" r="4"/><path d="M12 2v3M12 19v3M2 12h3M19 12h3"/>',
  rows: '<path d="M4 6h16M4 12h16M4 18h16"/><path d="M9 3v18"/>',
  filter: '<path d="M4 5h16l-6 7.5V19l-4 1.5v-8z"/>',
  wave: '<path d="M3 9c3-3 6 3 9 0s6 3 9 0M3 15c3-3 6 3 9 0s6 3 9 0"/>',
  join: '<rect x="3" y="7" width="6" height="10" rx="1.5"/><rect x="15" y="7" width="6" height="10" rx="1.5"/><path d="M9 12h6"/>',
  mail: '<rect x="3" y="5" width="18" height="14" rx="2"/><path d="m3.5 6.5 8.5 6.5 8.5-6.5"/>',
  shrink: '<path d="M4 14h6v6M20 10h-6V4M14 10l7-7M3 21l7-7"/>',
};
function icon(name) { const s = document.createElementNS('http://www.w3.org/2000/svg', 'svg'); s.setAttribute('viewBox', '0 0 24 24'); s.innerHTML = ICONS[name] || ICONS.table; return s; }
const TOOL_ICON = {
  'excel-viewer': 'eye', 'merge-excel': 'merge', 'csv-merger': 'stack', 'excel-splitter': 'split', 'excel-to-csv': 'csv', 'csv-to-excel': 'table',
  'excel-to-json': 'braces', 'json-to-excel': 'json2', 'pdf-to-excel': 'pdf', 'image-to-excel': 'image', 'compare-excel': 'diff', 'remove-duplicates': 'dedupe',
  vlookup: 'search', xlookup: 'target', hlookup: 'rows', 'multi-criteria-lookup': 'filter', 'fuzzy-lookup': 'wave', 'combine-columns': 'join', 'mail-merge': 'mail', 'compress-excel': 'shrink',
};
const tIcon = t => icon(TOOL_ICON[t.id]);

/* file handed over from the home drop stage to the next tool that can use it */
const PENDING = {
  file: null,
  take(accept) {
    const f = this.file; if (!f || !accept) return null;
    const ok = accept.split(',').some(a => { a = a.trim().toLowerCase(); return a.endsWith('/*') ? f.type.startsWith(a.slice(0, -1)) : f.name.toLowerCase().endsWith(a); });
    if (!ok) return null;
    this.file = null; return f;
  },
};
function suggestFor(file) {
  const n = file.name.toLowerCase();
  if (/\.pdf$/.test(n)) return ['pdf-to-excel'];
  if (file.type.startsWith('image/')) return ['image-to-excel'];
  if (/\.json$/.test(n)) return ['json-to-excel'];
  if (/\.(csv|tsv|txt)$/.test(n)) return ['csv-to-excel', 'excel-viewer', 'remove-duplicates', 'vlookup', 'excel-splitter', 'compare-excel', 'combine-columns', 'excel-to-json', 'mail-merge'];
  if (/\.(xlsx|xlsm|xlsb|xls|ods)$/.test(n)) return ['excel-viewer', 'remove-duplicates', 'vlookup', 'excel-to-csv', 'excel-splitter', 'compare-excel', 'combine-columns', 'excel-to-json', 'mail-merge', 'compress-excel'];
  return [];
}

/* ---------- pages ---------- */
const FEATURED = { 'merge-excel': 'hero wide', vlookup: 'wide', 'compare-excel': '', 'remove-duplicates': '' };
const ORDER = ['merge-excel', 'vlookup', 'compare-excel', 'remove-duplicates', 'pdf-to-excel', 'excel-splitter', 'fuzzy-lookup', 'excel-viewer'];
function tile(t, extra = '') {
  return h('a', { class: 'tile ' + extra, href: '#/' + t.id }, h('span', { class: 'tic' }, tIcon(t)), h('span', { class: 'go' }, icon('arrow')),
    h('h3', null, t.name), h('p', null, t.desc), h('span', { class: 'tag' }, CATS.find(c => c.id === t.cat).name.toLowerCase()));
}
let homeFilter = 'all';
function home(app) {
  const sugg = h('div', { class: 'suggest', hidden: true });
  const zone = dropZone({ accept: '', label: 'Drop any file to get started', hint: 'Excel, CSV, JSON, PDF or an image — we’ll suggest what you can do',
    onFiles: fs => {
      const f = fs[0]; const ids = suggestFor(f);
      PENDING.file = f;
      sugg.hidden = false;
      sugg.replaceChildren(h('h4', null, 'What should we do with ', h('span', null, f.name), '?'),
        ids.length ? h('div', { class: 'pills' }, ids.map(id => { const t = TOOLS.find(x => x.id === id); return h('a', { class: 'pill', href: '#/' + id }, tIcon(t), t.name); }))
                   : h('p', { class: 'muted' }, 'That file type isn’t supported yet. Try a spreadsheet, CSV, JSON, PDF or image.'));
    } });
  const grid = h('div', { class: 'bento' });
  const seg = h('div', { class: 'seg' });
  const paint = () => {
    seg.replaceChildren(...[['all', 'All'], ...CATS.map(c => [c.id, c.name])].map(([id, l]) => h('button', { class: homeFilter === id ? 'on' : '', onclick: () => { homeFilter = id; paint(); } }, l)));
    const list = homeFilter === 'all' ? [...TOOLS].sort((a, b) => (ORDER.indexOf(a.id) + 1 || 99) - (ORDER.indexOf(b.id) + 1 || 99)) : TOOLS.filter(t => t.cat === homeFilter);
    grid.replaceChildren(...list.map(t => tile(t, homeFilter === 'all' ? FEATURED[t.id] || '' : '')));
  };
  paint();
  app.replaceChildren(
    h('section', { class: 'home-top' },
      h('div', { class: 'home-copy' },
        h('span', { class: 'eyebrow' }, h('i'), `${TOOLS.length} tools · runs offline`),
        h('h1', null, 'Spreadsheet chores, ', h('em', null, 'done in seconds.')),
        h('p', null, 'Merge, split, compare, look up and convert Excel files without formulas, installs or uploads. Your data stays on your device.'),
        h('div', { class: 'facts' }, [['lock', 'Never uploaded'], ['bolt', 'Instant results'], ['user', 'No sign-up']].map(([i, t]) => h('span', null, icon(i), t)))),
      h('div', { class: 'stage' }, h('div', { class: 'stage-in' }, zone, sugg))),
    h('div', { class: 'filters' }, h('h2', null, 'Toolbox'), seg),
    grid);
  document.title = 'Excel tools · gumber.ai';
}
function buildNav(activeId) {
  $('#sideNav').replaceChildren(...CATS.map(c => h('div', { class: 'nav-group' }, h('h6', null, c.name),
    TOOLS.filter(t => t.cat === c.id).map(t => h('a', { class: 'nav-item' + (t.id === activeId ? ' on' : ''), href: '#/' + t.id }, tIcon(t), t.name)))));
}
function route() {
  const app = $('#app'); const id = location.hash.replace(/^#\/?/, '');
  const t = TOOLS.find(x => x.id === id);
  closeMenu(); buildNav(t && t.id); window.scrollTo(0, 0);
  if (!t) return home(app);
  const cat = CATS.find(c => c.id === t.cat);
  const related = TOOLS.filter(x => x.cat === t.cat && x.id !== t.id).slice(0, 4);
  const body = h('div');
  app.replaceChildren(
    h('header', { class: 'tool-head' }, h('span', { class: 'tic' }, tIcon(t)),
      h('div', null, h('div', { class: 'crumb' }, h('a', { href: '#/' }, 'tools'), ' / ', cat.name.toLowerCase()),
        h('h1', null, t.name), h('p', null, t.desc), h('span', { class: 'lock' }, icon('lock'), 'Processed locally — nothing leaves your browser'))),
    body,
    related.length ? h('section', { class: 'more' }, h('h3', null, 'Keep going'), h('div', { class: 'bento' }, related.map(r => tile(r)))) : null);
  document.title = `${t.name} · excel.gumber.ai`;
  t.render(body);
}

/* command palette */
let palSel = 0, palItems = [];
function openPalette() { const p = $('#palette'); p.hidden = false; const i = $('#palInput'); i.value = ''; drawPalette(); i.focus(); }
function closePalette() { $('#palette').hidden = true; }
function drawPalette() {
  const q = $('#palInput').value.trim().toLowerCase();
  palItems = TOOLS.filter(t => !q || (t.name + ' ' + t.desc + ' ' + t.id).toLowerCase().includes(q));
  palSel = Math.min(palSel, Math.max(0, palItems.length - 1));
  $('#palList').replaceChildren(...(palItems.length ? palItems.map((t, i) => h('li', { class: i === palSel ? 'sel' : '', onmouseenter: () => { palSel = i; drawPalette(); }, onclick: () => go(t) },
    h('span', { class: 'tic' }, tIcon(t)), h('div', null, h('b', null, t.name), h('span', null, t.desc)))) : [h('li', { class: 'empty' }, 'No tools match.')]));
  $('#palList .sel')?.scrollIntoView({ block: 'nearest' });
}
function go(t) { closePalette(); location.hash = '#/' + t.id; }
function initPalette() {
  $('#findBtn').addEventListener('click', openPalette); $('#mPalette').addEventListener('click', openPalette);
  $('#palInput').addEventListener('input', () => { palSel = 0; drawPalette(); });
  $('#palette').addEventListener('click', e => { if (e.target.id === 'palette') closePalette(); });
  document.addEventListener('keydown', e => {
    if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 'k') { e.preventDefault(); $('#palette').hidden ? openPalette() : closePalette(); return; }
    if ($('#palette').hidden) return;
    if (e.key === 'Escape') closePalette();
    else if (e.key === 'ArrowDown') { e.preventDefault(); palSel = Math.min(palSel + 1, palItems.length - 1); drawPalette(); }
    else if (e.key === 'ArrowUp') { e.preventDefault(); palSel = Math.max(palSel - 1, 0); drawPalette(); }
    else if (e.key === 'Enter' && palItems[palSel]) go(palItems[palSel]);
  });
}
/* mobile menu */
function closeMenu() { $('#side').classList.remove('open'); $('#scrim').classList.remove('open'); }
function initMenu() {
  $('#menuBtn').addEventListener('click', () => { $('#side').classList.add('open'); $('#scrim').classList.add('open'); });
  $('#scrim').addEventListener('click', closeMenu);
}

/* theme */
function initTheme() {
  let saved = null; try { saved = localStorage.getItem('theme'); } catch (e) {}
  if (saved) document.documentElement.dataset.theme = saved;
  $('#themeBtn').addEventListener('click', () => {
    const cur = document.documentElement.dataset.theme || (matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light');
    const next = cur === 'dark' ? 'light' : 'dark';
    document.documentElement.dataset.theme = next; try { localStorage.setItem('theme', next); } catch (e) {}
  });
}

window.addEventListener('hashchange', route);
initTheme(); initPalette(); initMenu(); $('#yr').textContent = new Date().getFullYear();
if (!window.XLSX) { $('#app').append(h('div', { class: 'warn' }, 'Could not load the spreadsheet engine. Check your connection and reload.')); }
route();
