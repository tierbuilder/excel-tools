// SEO + AI content for every page. Single source of truth for titles, descriptions, how-tos and FAQs.
export const SITE = {
  url: 'https://exceltools.gumber.ai',
  name: 'Excel Tools by gumber.ai',
  brand: 'Excel Tools by gumber.ai',
  tagline: 'Free, private Excel & CSV tools that run in your browser',
  description: 'Free online Excel and CSV tools — merge, split, compare, convert, VLOOKUP, remove duplicates and more. Everything runs in your browser, so files are never uploaded. No sign-up.',
};

export const CATS = [
  { id: 'convert', name: 'Convert' },
  { id: 'combine', name: 'Merge & Split' },
  { id: 'data', name: 'Clean, Compare & Lookup' },
  { id: 'other', name: 'View & Optimize' },
];

// Shared limits shown on every tool page, the About page and llms.txt.
export const LIMITS = [
  'Everything runs in your browser’s memory, so the practical file size depends on your device. As a rough guide: up to ~50 MB or ~500,000 rows works well on a modern laptop; phones and tablets handle much less.',
  'Very large files can make the tab slow or crash. If that happens, split the file first or use a desktop computer.',
  'Outputs keep values and formulas’ results but not formatting, charts, images, pivot tables or macros.',
  'Password-protected workbooks can’t be opened.',
  'PDF to Excel works on digital PDFs only; scanned pages need the Image to Excel (OCR) tool. OCR accuracy depends on image quality.',
  'Fuzzy lookups compare every unmatched value against the lookup table, so very large tables (tens of thousands of rows on both sides) can take a while.',
];

export const DISCLAIMER = 'Excel Tools by gumber.ai is a hobby project and is completely free to use. It is provided “as is”, without warranty of any kind. Always double-check results before relying on them, and keep a copy of your original files.';

export const HOME_FAQ = [
  ['Are these Excel tools really free?', 'Yes. Excel Tools by gumber.ai is a free hobby project — no sign-up, no watermark, no paid tier and no file count limits.'],
  ['Are my files uploaded to a server?', 'No. Every tool runs entirely in your web browser using JavaScript. Your spreadsheets never leave your device, which also makes the tools work on confidential data.'],
  ['Which file formats are supported?', 'XLSX, XLSM, XLSB, XLS, ODS, CSV, TSV and TXT spreadsheets, plus JSON, PDF (digital) and images (PNG, JPG, WEBP) for the conversion tools.'],
  ['How big can my file be?', 'It depends on your device’s memory. Files up to about 50 MB or 500,000 rows usually work well on a laptop; phones handle less. Split very large files first.'],
  ['Do I need Microsoft Excel installed?', 'No. The tools work in any modern browser — Chrome, Edge, Safari or Firefox — on Windows, Mac, Linux, Android and iOS.'],
];

/** id → page content. name/desc must match the TOOLS list in site/app.js. */
export const TOOLS = [
  {
    id: 'excel-to-csv', cat: 'convert', name: 'Excel to CSV', desc: 'Convert XLSX/XLS sheets to CSV with your choice of delimiter. Export one sheet or all.',
    title: 'Excel to CSV Converter — Free, Online, No Upload',
    meta: 'Convert XLSX, XLS or ODS to CSV online for free. Choose comma, semicolon, tab or pipe, export one or all sheets, UTF-8 ready. Files stay in your browser.',
    intro: 'Turn any Excel worksheet into a clean CSV file. Pick the delimiter your system expects, add a UTF-8 BOM so accented characters open correctly in Excel, and export a single sheet or every sheet at once as a ZIP.',
    steps: ['Drop your Excel file (XLSX, XLS, XLSM, ODS) onto the page.', 'Choose the sheet and the delimiter — comma, semicolon, tab or pipe.', 'Click “Convert to CSV” and download the file (or a ZIP of all sheets).'],
    faq: [['Can I convert all sheets at once?', 'Yes. Choose “All sheets (ZIP)” and every worksheet is saved as its own CSV inside one ZIP file.'], ['Why do accents look wrong when I open the CSV in Excel?', 'Excel needs a UTF-8 byte-order mark to detect the encoding. Keep the “Add UTF-8 BOM” option ticked.'], ['Are formulas kept?', 'CSV can only store values, so each formula is saved as its calculated result.']],
  },
  {
    id: 'csv-to-excel', cat: 'convert', name: 'CSV to Excel', desc: 'Turn CSV, TSV or delimited text into a proper .xlsx workbook. Auto-detects delimiters and numbers.',
    title: 'CSV to Excel Converter (XLSX) — Free & Private',
    meta: 'Convert CSV, TSV or TXT files to Excel XLSX online. Auto-detects delimiters, converts numbers, keeps leading zeros, and merges several CSVs into tabs.',
    intro: 'Convert comma, semicolon, tab or pipe-separated text into a real Excel workbook. Numbers become numbers, codes with leading zeros (like 007 or ZIP codes) stay as text, and several CSV files can be combined as separate tabs.',
    steps: ['Add one or more CSV, TSV or TXT files.', 'Leave the delimiter on “Auto-detect” or pick one manually.', 'Click “Convert to Excel” and download the .xlsx workbook.'],
    faq: [['Will leading zeros be removed?', 'No. Values such as 00123 are kept as text so the zeros are preserved.'], ['Can I combine several CSVs into one workbook?', 'Yes. Add multiple files and each one becomes its own sheet.'], ['Which delimiters are supported?', 'Comma, semicolon, tab and pipe — detected automatically.']],
  },
  {
    id: 'excel-to-json', cat: 'convert', name: 'Excel to JSON', desc: 'Convert spreadsheets to JSON arrays or objects. Supports nested keys like "address.city".',
    title: 'Excel to JSON Converter — XLSX/CSV to JSON Online',
    meta: 'Convert Excel or CSV to JSON online for free. Array of objects or arrays, nested keys from dotted headers, all sheets at once, pretty or minified.',
    intro: 'Convert spreadsheet rows into JSON for APIs, apps and config files. The first row becomes the keys; headers like “address.city” become nested objects. Export one sheet or every sheet keyed by name.',
    steps: ['Drop an Excel or CSV file.', 'Choose array of objects or array of arrays, and whether to nest dotted headers.', 'Click “Convert to JSON”, then copy or download the result.'],
    faq: [['How are nested objects created?', 'Use dots in column headers, e.g. “user.name” and “user.email” become { "user": { "name", "email" } }.'], ['How are dates exported?', 'Dates are written as ISO strings such as 2026-01-31.'], ['Can I skip empty cells?', 'Yes, tick “Omit empty cells”.']],
  },
  {
    id: 'json-to-excel', cat: 'convert', name: 'JSON to Excel', desc: 'Paste or upload JSON and get a clean spreadsheet. Nested objects are flattened into columns.',
    title: 'JSON to Excel Converter — Flatten JSON to XLSX',
    meta: 'Convert JSON to Excel online. Paste or upload JSON; nested objects are flattened into columns and arrays of objects become rows. Free, private, no upload.',
    intro: 'Paste an API response or upload a .json file and get a tidy spreadsheet. Nested objects become dotted columns (address.city), simple arrays are joined, and an object of arrays becomes one sheet per key.',
    steps: ['Paste JSON into the box or drop a .json file.', 'Click “Convert to Excel”.', 'Review the preview and download the .xlsx file.'],
    faq: [['What JSON shapes are supported?', 'An array of objects, an array of arrays, a single object, or an object whose values are arrays (each becomes a sheet).'], ['What happens to nested arrays?', 'Arrays of simple values are joined with commas; arrays of objects are kept as JSON text in the cell.'], ['Is there a size limit?', 'Only your browser’s memory — see the limitations below.']],
  },
  {
    id: 'pdf-to-excel', cat: 'convert', name: 'PDF to Excel', desc: 'Extract tables and text from digital PDFs into rows and columns.',
    title: 'PDF to Excel Converter — Extract Tables Free',
    meta: 'Convert PDF tables to Excel online for free. Extract rows and columns from digital PDFs, one sheet per page or combined. Processed in your browser.',
    intro: 'Pull tables out of bank statements, invoices and reports. Text is grouped into rows by position and aligned into columns, so most tabular PDFs come out ready to sort and filter.',
    steps: ['Choose a digital (text-based) PDF.', 'Pick one sheet per page or all pages in one sheet; adjust the column gap if columns merge or split.', 'Click “Extract to Excel” and download.'],
    faq: [['Does it work on scanned PDFs?', 'Not directly. Scanned pages are images — take a screenshot and use the Image to Excel (OCR) tool instead.'], ['Columns are merged or split incorrectly. What can I do?', 'Lower the “Column gap” to split more, raise it to merge more, or turn off column alignment.'], ['Is my PDF uploaded?', 'No, it is parsed in your browser with pdf.js.']],
  },
  {
    id: 'image-to-excel', cat: 'convert', name: 'Image to Excel (OCR)', desc: 'Read tables from screenshots, photos or scans and turn them into a spreadsheet.',
    title: 'Image to Excel (OCR) — Convert Screenshot to Table',
    meta: 'Convert images, screenshots and scans to Excel with free OCR. Supports English, Hindi, Spanish, French, German and more. Runs locally in your browser.',
    intro: 'Turn a screenshot, photo or scanned table into editable rows and columns. Text recognition runs on your device with Tesseract OCR and supports many languages.',
    steps: ['Drop a PNG, JPG or WEBP image.', 'Choose the language and how columns are separated.', 'Click “Extract table”, check the result, and download XLSX or CSV.'],
    faq: [['How accurate is the OCR?', 'Clear, high-resolution screenshots work best. Photos with blur, skew or shadows reduce accuracy — always check numbers.'], ['Why is the first run slow?', 'The OCR engine and language data (a few MB) are downloaded once, then cached by your browser.'], ['Which languages are supported?', 'English, Hindi, Spanish, French, German, Portuguese, Italian, Chinese (Simplified), Japanese and Arabic.']],
  },
  {
    id: 'merge-excel', cat: 'combine', name: 'Merge Excel Files', desc: 'Combine multiple workbooks into one sheet or one workbook. Columns are matched by header name.',
    title: 'Merge Excel Files Online — Combine Workbooks Free',
    meta: 'Merge multiple Excel and CSV files into one sheet or one workbook. Columns are matched by header, with an optional source-file column. Free, no upload.',
    intro: 'Combine monthly reports, exports from different teams, or many CSVs into a single file. Rows are stacked and columns are lined up by header name — even when files have columns in a different order.',
    steps: ['Add two or more Excel or CSV files.', 'Choose “Stack all rows into one sheet” or “Keep each sheet as a separate tab”.', 'Click “Merge files” and download the result.'],
    faq: [['What if my files have different columns?', 'All columns are kept. Rows from files missing a column get an empty cell there.'], ['Can I tell which file each row came from?', 'Yes, keep “Add a Source column” ticked.'], ['Can I merge all sheets inside each workbook?', 'Yes, choose “All sheets of each file”.']],
  },
  {
    id: 'csv-merger', cat: 'combine', name: 'CSV Merger', desc: 'Join many CSV files into a single CSV, lining up columns by header.',
    title: 'CSV Merger — Combine Multiple CSV Files Online',
    meta: 'Merge many CSV files into one CSV online. Columns are aligned by header name and an optional source column is added. Free and private.',
    intro: 'Join any number of CSV exports into one file. Headers are matched by name (case-insensitive), so files with columns in a different order still line up.',
    steps: ['Add your CSV files.', 'Decide whether to add a “Source” column.', 'Click “Merge files” and download the combined CSV.'],
    faq: [['Do all CSVs need the same delimiter?', 'No, each file’s delimiter is detected separately.'], ['Is the header row repeated?', 'No, the output has a single header row.'], ['How many files can I merge?', 'As many as your browser’s memory allows.']],
  },
  {
    id: 'excel-splitter', cat: 'combine', name: 'Excel Splitter', desc: 'Split a workbook by sheet, by row count, into equal parts, or by the values in a column.',
    title: 'Excel Splitter — Split Excel by Rows, Sheets or Column',
    meta: 'Split large Excel or CSV files online: every N rows, into N equal parts, by unique column values, or one file per sheet. Download as ZIP or tabs.',
    intro: 'Break a large spreadsheet into manageable pieces. Split every 1,000 rows for imports, into equal parts for teammates, by region or category, or save each sheet as its own file.',
    steps: ['Drop the spreadsheet you want to split.', 'Choose a split mode and set N or the column.', 'Click “Split file”, then download a ZIP of files or one workbook with a tab per part.'],
    faq: [['Is the header row kept in every part?', 'Yes, by default the header is repeated in each file.'], ['Can I split by a column like “Country”?', 'Yes — choose “By unique values in a column” and every value gets its own file.'], ['What output formats are available?', 'A ZIP of XLSX files, a ZIP of CSV files, or a single workbook with one tab per part.']],
  },
  {
    id: 'compare-excel', cat: 'data', name: 'Compare Excel Files', desc: 'Spot changed cells, added rows and removed rows between two versions of a sheet.',
    title: 'Compare Two Excel Files — Find Differences Online',
    meta: 'Compare two Excel or CSV files and highlight changed cells, added rows and removed rows. Match rows by a key column. Download a diff report. Free.',
    intro: 'See exactly what changed between two versions of a spreadsheet. Match rows by position or by a key column such as ID, then review highlighted changes and export a report listing every changed cell.',
    steps: ['Upload the original (A) and the new version (B).', 'Choose how to match rows — by position or by a key column.', 'Click “Compare” and download the diff report.'],
    faq: [['Why use a key column?', 'If rows were inserted, deleted or re-sorted, matching by an ID column compares the right rows with each other.'], ['Can I ignore case and spaces?', 'Yes, both options are available.'], ['What does the report contain?', 'A summary sheet with every row’s status and a “Cell changes” sheet listing old and new values.']],
  },
  {
    id: 'remove-duplicates', cat: 'data', name: 'Remove Duplicates', desc: 'Find and remove duplicate rows based on all columns or just the ones you pick.',
    title: 'Remove Duplicates in Excel Online — Free Dedupe Tool',
    meta: 'Remove duplicate rows from Excel or CSV online. Dedupe on all columns or selected ones, ignore case and spaces, keep first or last. Download duplicates too.',
    intro: 'Clean up mailing lists, CRM exports and transaction logs. Decide which columns define a duplicate, ignore case and extra spaces, and keep the first or last occurrence.',
    steps: ['Drop your spreadsheet.', 'Select the columns that define a duplicate and set the matching options.', 'Click “Remove duplicates” and download the cleaned file (and optionally the removed rows).'],
    faq: [['Can I dedupe by email only?', 'Yes — select just the email column.'], ['Is “John ” the same as “john”?', 'With “Ignore case” and “Ignore extra spaces” on, yes.'], ['Can I see what was removed?', 'Yes, download the removed duplicates as a separate file.']],
  },
  {
    id: 'vlookup', cat: 'data', name: 'Excel VLOOKUP', desc: 'Pull matching values from a second table into your sheet — no formulas, no #REF errors.',
    title: 'Online VLOOKUP Tool — Match Two Excel Files Without Formulas',
    meta: 'Do a VLOOKUP between two Excel or CSV files online without formulas. Match on a column, bring back one or many columns, handle missing values. Free.',
    intro: 'Bring columns from one table into another by matching a shared value such as ID, SKU or email — without writing a formula. Return several columns at once and choose what to show when no match is found.',
    steps: ['Upload the main table and the lookup table (or the same file twice with different sheets).', 'Pick the matching columns and the columns to bring back.', 'Click “Run lookup” and download the result.'],
    faq: [['Can I return more than one column?', 'Yes, select as many return columns as you like.'], ['What happens when there are several matches?', 'Choose first match, last match, all matches joined, or the count.'], ['Does the lookup column have to be the first column?', 'No — any column can be matched, unlike Excel’s VLOOKUP.']],
  },
  {
    id: 'xlookup', cat: 'data', name: 'XLOOKUP / INDEX-MATCH', desc: 'Look up any column in any direction and return one or many result columns.',
    title: 'XLOOKUP & INDEX-MATCH Online — Lookup in Any Direction',
    meta: 'Free online XLOOKUP / INDEX-MATCH: match any column against any column and return values from left or right. Custom “not found” text. No upload.',
    intro: 'The flexibility of XLOOKUP and INDEX-MATCH without the formulas: match on any column and return columns to the left or right, with your own “not found” value.',
    steps: ['Upload the main table and the lookup table.', 'Choose the match column on each side and the columns to return.', 'Click “Run lookup” and download.'],
    faq: [['How is this different from the VLOOKUP tool?', 'They share the same engine; this page is set up for XLOOKUP / INDEX-MATCH style use.'], ['Can I look up to the left?', 'Yes, return columns can be anywhere in the lookup table.'], ['Is matching case-sensitive?', 'Not by default; you can switch it on.']],
  },
  {
    id: 'hlookup', cat: 'data', name: 'HLOOKUP Tool', desc: 'Lookups for horizontal tables where headers run down the first column.',
    title: 'HLOOKUP Online — Lookup in Horizontal Tables',
    meta: 'Free online HLOOKUP for horizontal Excel tables where headers are in the first column. Match and return rows without formulas, privately in your browser.',
    intro: 'For tables laid out sideways — with field names down the first column and records across. Match values across rows and bring back the rows you need.',
    steps: ['Upload the main and lookup tables (horizontal layout).', 'Pick the match row and the rows to return.', 'Run the lookup and download the result in the same horizontal layout.'],
    faq: [['What is a horizontal table?', 'One where each record is a column and each field is a row.'], ['Can I switch to vertical?', 'Yes, change “Table orientation” at any time.'], ['Is the result transposed?', 'No, it is returned in the same horizontal layout as your input.']],
  },
  {
    id: 'multi-criteria-lookup', cat: 'data', name: 'Multi-Criteria Lookup', desc: 'Match rows on two or more columns at once, e.g. Name + Date.',
    title: 'Multi-Criteria Lookup — Match on Multiple Columns',
    meta: 'Look up values that match on two or more columns at once (e.g. Name + Date or SKU + Store) between two Excel files. Free, private, no formulas.',
    intro: 'When one column isn’t unique, match on several together — like Customer + Month or SKU + Warehouse — and pull back the right values.',
    steps: ['Upload both tables.', 'Add a criterion for each column pair that must match.', 'Choose the return columns, run the lookup and download.'],
    faq: [['How many criteria can I add?', 'As many as you need.'], ['Do criteria need the same column names?', 'No, you pair any column in the main table with any column in the lookup table.'], ['Can I combine this with fuzzy matching?', 'Yes, set the match type to fuzzy.']],
  },
  {
    id: 'fuzzy-lookup', cat: 'data', name: 'Excel Fuzzy Lookup', desc: 'Match names that are almost the same — typos, abbreviations and spacing differences.',
    title: 'Excel Fuzzy Lookup Online — Match Similar Text',
    meta: 'Fuzzy match two Excel or CSV lists online. Match names and companies despite typos and spacing with an adjustable similarity threshold and match scores.',
    intro: 'Match company names, product titles or people even when they are spelled slightly differently. Set a minimum similarity and get a score and the matched value for every row so you can review the results.',
    steps: ['Upload the list to match and the reference list.', 'Pick the columns and set the minimum similarity (80% is a good start).', 'Run the lookup, review match scores, and download.'],
    faq: [['How is similarity measured?', 'With edit (Levenshtein) distance: 100% means identical, 80% means about one character in five differs.'], ['Exact matches are found first?', 'Yes, fuzzy matching only runs for values without an exact match.'], ['It is slow on big files — why?', 'Each unmatched value is compared with every lookup value. Clean obvious differences first or split the file.']],
  },
  {
    id: 'combine-columns', cat: 'data', name: 'Combine Columns', desc: 'Join two or more columns into one with any separator — e.g. First + Last name.',
    title: 'Combine Columns in Excel Online — Concatenate Free',
    meta: 'Combine two or more Excel columns into one with any separator — e.g. First + Last name or full addresses. Skip blanks, remove originals. Free.',
    intro: 'Join columns like First Name and Last Name, or Street, City and ZIP, into a single column. Choose the order, the separator and whether to skip blanks.',
    steps: ['Drop your spreadsheet.', 'Click the columns in the order you want them joined and choose a separator.', 'Click “Combine columns” and download.'],
    faq: [['Can I control the order?', 'Yes, the order you click columns is the order they are joined.'], ['Can I use a line break as separator?', 'Yes, choose “New line”.'], ['Are the original columns kept?', 'By default yes; tick “Remove original columns” to drop them.']],
  },
  {
    id: 'mail-merge', cat: 'data', name: 'Mail Merge', desc: 'Generate personalised letters, emails or messages from a spreadsheet and a template.',
    title: 'Mail Merge from Excel Online — Personalised Documents',
    meta: 'Mail merge from Excel or CSV online: write a template with {{fields}} and generate one personalised letter, email or message per row. Free, private.',
    intro: 'Write a template once with placeholders like {{Name}} and generate one personalised document per row — letters, emails, certificates text or WhatsApp messages.',
    steps: ['Upload your recipient list.', 'Write the template and insert fields by clicking them.', 'Preview, then download a ZIP of files, one combined file, or a spreadsheet with a Message column.'],
    faq: [['Can it send emails?', 'No, it generates the text. Use the spreadsheet output with your email tool’s own merge feature.'], ['Can I name each file after a column?', 'Yes, set the file name pattern, e.g. {{Name}}.'], ['Which output formats are available?', 'TXT, HTML or Markdown files, or an XLSX with a Message column.']],
  },
  {
    id: 'excel-viewer', cat: 'other', name: 'Online Excel Viewer', desc: 'Open XLSX, XLS, ODS and CSV files in your browser. Search and browse every sheet.',
    title: 'Online Excel Viewer — Open XLSX Without Excel',
    meta: 'Open and view XLSX, XLS, XLSM, ODS and CSV files online without Microsoft Excel. Browse sheets and search rows. Free, private, works on mobile.',
    intro: 'Open a spreadsheet on any device without installing Excel. Switch between sheets, search across rows and export to XLSX or CSV.',
    steps: ['Drop a spreadsheet file.', 'Pick a sheet and use the filter box to search.', 'Optionally download as XLSX or CSV.'],
    faq: [['Can I edit cells?', 'The viewer is read-only; use the other tools to transform data.'], ['Does it work on iPhone and Android?', 'Yes, in any modern mobile browser (for smaller files).'], ['Are formulas shown?', 'The calculated values are shown.']],
  },
  {
    id: 'compress-excel', cat: 'other', name: 'Compress Excel', desc: 'Shrink bloated XLSX/XLS files by rewriting them with maximum compression.',
    title: 'Compress Excel Files Online — Reduce XLSX Size',
    meta: 'Reduce Excel file size online by rewriting XLSX/XLS with maximum compression. See size before and after. Formatting is not kept. Free, private.',
    intro: 'Shrink bloated workbooks by rewriting the data with maximum compression. Best for data-heavy files that need to be emailed or uploaded.',
    steps: ['Add one or more workbooks.', 'Click “Compress”.', 'Compare sizes and download the smaller files.'],
    faq: [['What is removed?', 'Formatting, charts, images and macros are dropped — values and formulas are kept. Keep your original.'], ['Why didn’t my file get smaller?', 'Files that are mostly images or already optimised may not shrink further.'], ['Can I compress XLS files?', 'Yes, they are saved as modern .xlsx.']],
  },
];
