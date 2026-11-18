/* global Office, Excel, pdfjsLib */
"use strict";

pdfjsLib.GlobalWorkerOptions.workerSrc =
  "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";

const state = {
  files: [],        // File objects chosen by the user
  fileErrors: {},   // file name -> error message
  tables: [],       // { file, page, index, rows: string[][], selected }
  inExcel: false,
  busy: false,
};

const $ = (id) => document.getElementById(id);

Office.onReady((info) => {
  state.inExcel = info.host === Office.HostType.Excel;
  init();
});

/* ------------------------------------------------------------------ UI */

function init() {
  const input = $("fileInput");
  const drop = $("dropZone");

  input.addEventListener("change", () => { addFiles(input.files); input.value = ""; });
  drop.addEventListener("keydown", (e) => {
    if (e.key === "Enter" || e.key === " ") { e.preventDefault(); input.click(); }
  });
  ["dragenter", "dragover"].forEach((ev) =>
    drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.add("over"); }));
  ["dragleave", "drop"].forEach((ev) =>
    drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.remove("over"); }));
  drop.addEventListener("drop", (e) => addFiles(e.dataTransfer.files));

  $("extractBtn").addEventListener("click", runExtraction);
  $("importBtn").addEventListener("click", runImport);
  $("selectAll").addEventListener("click", () => setAllSelected(true));
  $("selectNone").addEventListener("click", () => setAllSelected(false));

  $("downloadBtn").addEventListener("click", runDownload);

  if (!state.inExcel) {
    // Opened in a normal browser (e.g. for Excel 2010): save an .xlsx file instead.
    $("importBtn").hidden = true;
    $("downloadBtn").hidden = false;
    $("outputTitle").textContent = "3. Save as Excel file";
    $("asTableLabel").textContent = "Add filter buttons to the header row";
  }
}

function addFiles(list) {
  for (const f of list) {
    const isPdf = /\.pdf$/i.test(f.name) || f.type === "application/pdf";
    if (!isPdf) continue;
    if (state.files.some((x) => x.name === f.name && x.size === f.size)) continue;
    state.files.push(f);
  }
  renderFiles();
}

function renderFiles() {
  const ul = $("fileList");
  ul.textContent = "";
  state.files.forEach((f, i) => {
    const li = document.createElement("li");
    if (state.fileErrors[f.name]) li.className = "error";
    const name = document.createElement("span");
    name.className = "name";
    name.textContent = f.name;
    name.title = state.fileErrors[f.name] || f.name;
    const rm = document.createElement("button");
    rm.className = "remove";
    rm.textContent = "×";
    rm.title = "Remove";
    rm.addEventListener("click", () => { state.files.splice(i, 1); renderFiles(); });
    li.append(name, rm);
    ul.append(li);
  });
  updateButtons();
}

function updateButtons() {
  $("extractBtn").disabled = state.busy || state.files.length === 0;
  const anySelected = state.tables.some((t) => t.selected);
  $("importBtn").disabled = state.busy || !anySelected || !state.inExcel;
  $("downloadBtn").disabled = state.busy || !anySelected;
}

function setStatus(msg, kind) {
  const s = $("status");
  s.textContent = msg || "";
  s.className = "status" + (kind ? " " + kind : "");
}

function setAllSelected(on) {
  state.tables.forEach((t) => { t.selected = on; });
  document.querySelectorAll("#results input[type=checkbox]").forEach((cb) => { cb.checked = on; });
  updateSummary();
  updateButtons();
}

function updateSummary() {
  const n = state.tables.length;
  const sel = state.tables.filter((t) => t.selected).length;
  $("resultsSummary").textContent = `${sel} of ${n} table${n === 1 ? "" : "s"} selected`;
}

function renderResults(notes) {
  const box = $("results");
  box.textContent = "";
  $("resultsSection").hidden = false;
  $("outputSection").hidden = state.tables.length === 0;

  for (const f of state.files) {
    const group = document.createElement("div");
    group.className = "file-group";
    const title = document.createElement("div");
    title.className = "file-title";
    title.textContent = f.name;
    group.append(title);

    const tables = state.tables.filter((t) => t.file === f.name);
    const note = notes[f.name];
    if (note || tables.length === 0) {
      const p = document.createElement("div");
      p.className = "note";
      p.textContent = note || "No tables found. Try a lower Column gap or Whole page mode.";
      group.append(p);
    }
    for (const t of tables) group.append(renderTableCard(t));
    box.append(group);
  }
  updateSummary();
}

function renderTableCard(t) {
  const card = document.createElement("div");
  card.className = "tbl";
  const head = document.createElement("div");
  head.className = "head";

  const label = document.createElement("label");
  const cb = document.createElement("input");
  cb.type = "checkbox";
  cb.checked = t.selected;
  cb.addEventListener("change", () => { t.selected = cb.checked; updateSummary(); updateButtons(); });
  const txt = document.createElement("span");
  txt.textContent = `Page ${t.page} · table ${t.index}`;
  const dims = document.createElement("span");
  dims.className = "dim";
  dims.textContent = `${t.rows.length} × ${t.rows[0] ? t.rows[0].length : 0}`;
  label.append(cb, txt, dims);

  const toggle = document.createElement("button");
  toggle.className = "link toggle";
  toggle.textContent = "Preview";
  head.append(label, toggle);
  card.append(head);

  let preview = null;
  toggle.addEventListener("click", () => {
    if (preview) { preview.remove(); preview = null; toggle.textContent = "Preview"; return; }
    preview = buildPreview(t.rows);
    card.append(preview);
    toggle.textContent = "Hide";
  });
  return card;
}

function buildPreview(rows) {
  const wrap = document.createElement("div");
  wrap.className = "preview";
  const table = document.createElement("table");
  const MAX = 10;
  rows.slice(0, MAX).forEach((r) => {
    const tr = document.createElement("tr");
    r.forEach((c) => {
      const td = document.createElement("td");
      td.textContent = c;
      td.title = c;
      tr.append(td);
    });
    table.append(tr);
  });
  wrap.append(table);
  if (rows.length > MAX) {
    const more = document.createElement("div");
    more.className = "more";
    more.textContent = `… ${rows.length - MAX} more rows`;
    wrap.append(more);
  }
  return wrap;
}

/* ---------------------------------------------------------- Extraction */

async function runExtraction() {
  let opts;
  try { opts = readOptions(); } catch (e) { setStatus(e.message, "error"); return; }

  state.busy = true;
  state.tables = [];
  state.fileErrors = {};
  updateButtons();
  const notes = {};

  for (const file of state.files) {
    try {
      const res = await extractTablesFromPdf(file, opts, (p, n) =>
        setStatus(`Reading ${file.name} (page ${p} of ${n})…`));
      res.tables.forEach((t) => { t.selected = true; state.tables.push(t); });
      if (res.pagesRead === 0) {
        notes[file.name] = "No pages matched the page range.";
      } else if (res.textlessPages === res.pagesRead) {
        notes[file.name] = "No text found. This looks like a scanned PDF; run OCR on it first.";
      } else if (res.textlessPages > 0) {
        notes[file.name] = `${res.textlessPages} page(s) had no text (scanned?) and were skipped.`;
      }
    } catch (e) {
      const msg = describePdfError(e);
      state.fileErrors[file.name] = msg;
      notes[file.name] = msg;
    }
  }

  state.busy = false;
  renderFiles();
  renderResults(notes);
  updateButtons();
  const n = state.tables.length;
  setStatus(n ? `Found ${n} table${n === 1 ? "" : "s"} in ${state.files.length} file(s).`
              : "No tables found.", n ? "ok" : "error");
}

function describePdfError(e) {
  if (e && e.name === "PasswordException") return "This PDF is password-protected.";
  if (e && e.name === "InvalidPDFException") return "This file is not a valid PDF.";
  return "Could not read this PDF: " + (e && e.message ? e.message : e);
}

function readOptions() {
  return {
    mode: $("mode").value,
    pages: parsePages($("pages").value),
    minRows: Math.max(1, parseInt($("minRows").value, 10) || 2),
    gapEm: Math.max(0.1, parseFloat($("gapEm").value) || 0.8),
    mergeWrapped: $("mergeWrapped").checked,
  };
}

function parsePages(s) {
  s = (s || "").trim();
  if (!s) return null;
  const ranges = [];
  for (const raw of s.split(",")) {
    const part = raw.trim();
    if (!part) continue;
    const m = part.match(/^(\d+)\s*(-)?\s*(\d+)?$/);
    if (!m) throw new Error(`Invalid page range "${part}". Use something like 1-3, 5`);
    const a = +m[1];
    const b = m[3] ? +m[3] : (m[2] ? Infinity : a);
    ranges.push([Math.min(a, b), Math.max(a, b)]);
  }
  return ranges.length ? ranges : null;
}

const pageWanted = (p, ranges) => !ranges || ranges.some(([a, b]) => p >= a && p <= b);

/**
 * Reads one PDF and returns the tables found on its pages.
 * file: anything with .name and .arrayBuffer() (File or Blob-like).
 */
async function extractTablesFromPdf(file, opts, onProgress) {
  const data = new Uint8Array(await file.arrayBuffer());
  const pdf = await pdfjsLib.getDocument({ data, isEvalSupported: false }).promise;
  const tables = [];
  let textlessPages = 0;
  let pagesRead = 0;
  try {
    for (let p = 1; p <= pdf.numPages; p++) {
      if (!pageWanted(p, opts.pages)) continue;
      pagesRead++;
      if (onProgress) onProgress(p, pdf.numPages);
      const page = await pdf.getPage(p);
      const items = await getPageItems(page);
      page.cleanup();
      if (!items.length) { textlessPages++; continue; }

      const lines = buildLines(items, opts);
      const grids = opts.mode === "page"
        ? [linesToGrid(lines, opts)].filter((g) => g && g.length)
        : detectTables(lines, opts);
      grids.forEach((rows, i) => tables.push({ file: file.name, page: p, index: i + 1, rows }));
    }
  } finally {
    pdf.destroy();
  }
  return { tables, textlessPages, pagesRead };
}

/** Text pieces on a page in top-down page coordinates. */
async function getPageItems(page) {
  const vp = page.getViewport({ scale: 1 });
  const tc = await page.getTextContent();
  const out = [];
  for (const it of tc.items) {
    if (!it.str || !it.str.trim()) continue;
    const t = pdfjsLib.Util.transform(vp.transform, it.transform);
    // Skip rotated / vertical text (axis labels, watermarks).
    if (Math.abs(t[1]) > Math.abs(t[0]) * 0.1) continue;
    const size = Math.hypot(t[2], t[3]) || Math.abs(t[0]) || 10;
    const width = it.width > 0 ? it.width : it.str.length * size * 0.5;
    const charW = width / Math.max(1, it.str.length);
    // A single text run can hold several cells separated by 2+ spaces; split them.
    const re = /\S+(?: \S+)*/g;
    let m;
    while ((m = re.exec(it.str))) {
      out.push({
        text: m[0],
        x0: t[4] + m.index * charW,
        x1: t[4] + (m.index + m[0].length) * charW,
        y: t[5],
        h: size,
      });
    }
  }
  return out;
}

function median(arr) {
  if (!arr.length) return 0;
  const s = [...arr].sort((a, b) => a - b);
  const m = s.length >> 1;
  return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2;
}

/** Groups text pieces into lines (same baseline), then into segments split at wide gaps. */
function buildLines(items, opts) {
  const sorted = [...items].sort((a, b) => a.y - b.y || a.x0 - b.x0);
  const lines = [];
  for (const it of sorted) {
    let line = null;
    for (let k = lines.length - 1; k >= Math.max(0, lines.length - 3); k--) {
      const tol = 0.45 * Math.min(lines[k].h, it.h);
      if (Math.abs(lines[k].y - it.y) <= tol) { line = lines[k]; break; }
    }
    if (!line) { line = { y: it.y, h: it.h, items: [] }; lines.push(line); }
    line.items.push(it);
    line.y += (it.y - line.y) / line.items.length;
    line.h = Math.max(line.h, it.h);
  }
  lines.sort((a, b) => a.y - b.y);

  for (const line of lines) {
    line.items.sort((a, b) => a.x0 - b.x0);
    const gapLimit = opts.gapEm * line.h;
    const segs = [];
    for (const it of line.items) {
      const last = segs[segs.length - 1];
      const gap = last ? it.x0 - last.x1 : Infinity;
      if (last && gap < gapLimit) {
        const space = gap > 0.15 * it.h ? " " : "";
        last.text += space + it.text;
        last.x1 = Math.max(last.x1, it.x1);
      } else {
        segs.push({ text: it.text, x0: it.x0, x1: it.x1 });
      }
    }
    line.segs = segs;
  }
  return lines;
}

/** Finds runs of consecutive multi-column lines and turns each run into a grid. */
function detectTables(lines, opts) {
  if (!lines.length) return [];
  const medH = median(lines.map((l) => l.h)) || 10;
  const spacings = [];
  for (let i = 1; i < lines.length; i++) {
    const d = lines[i].y - lines[i - 1].y;
    if (d > 0) spacings.push(d);
  }
  const lineGap = median(spacings) || medH * 1.4;
  const maxGap = Math.max(lineGap * 2.2, medH * 2.5);
  const MAX_SINGLE = 2; // single-cell lines allowed inside a table (section labels, wrapped text)

  const runs = [];
  let cur = null;
  let pending = [];
  const close = () => { if (cur) runs.push(cur); cur = null; pending = []; };

  // Once a table has a few rows, its own row spacing decides where it ends,
  // so a sentence between two tables does not glue them together.
  const rowLimit = (run) => {
    if (run.length < 2) return maxGap;
    const steps = [];
    for (let i = 1; i < run.length; i++) steps.push(run[i].y - run[i - 1].y);
    return Math.min(maxGap, Math.max(median(steps) * 1.7, medH * 1.3));
  };

  for (const ln of lines) {
    const prev = cur ? (pending.length ? pending[pending.length - 1] : cur[cur.length - 1]) : null;
    const near = prev && ln.y - prev.y <= rowLimit(cur);
    if (ln.segs.length >= 2) {
      if (cur && near) { cur.push(...pending, ln); pending = []; }
      else { close(); cur = [ln]; }
    } else if (cur && near && pending.length < MAX_SINGLE) {
      pending.push(ln);
    } else {
      close();
    }
  }
  close();

  const grids = [];
  for (const run of runs) {
    const multi = run.filter((l) => l.segs.length >= 2).length;
    if (multi < opts.minRows) continue;
    const grid = linesToGrid(run, opts);
    if (grid && grid.length && grid[0].length >= 2) grids.push(grid);
  }
  return grids;
}

/** Builds columns from the empty vertical gutters shared by the lines, then fills cells. */
function linesToGrid(lines, opts) {
  if (!lines.length) return null;
  const basis = lines.filter((l) => l.segs.length >= 2);
  const ref = basis.length ? basis : lines;

  let minX = Infinity, maxX = -Infinity;
  for (const l of lines) for (const s of l.segs) {
    minX = Math.min(minX, s.x0);
    maxX = Math.max(maxX, s.x1);
  }
  const x0 = Math.floor(minX);
  const width = Math.max(1, Math.ceil(maxX) - x0 + 1);

  // How many lines have text covering each x position (1pt bins).
  const cov = new Uint16Array(width);
  for (const l of ref) {
    const row = new Uint8Array(width);
    for (const s of l.segs) {
      const a = Math.max(0, Math.floor(s.x0) - x0);
      const b = Math.min(width, Math.ceil(s.x1) - x0);
      for (let x = a; x < b; x++) row[x] = 1;
    }
    for (let x = 0; x < width; x++) cov[x] += row[x];
  }

  const allow = Math.floor(ref.length * 0.1); // tolerate a few cells spilling across a gutter
  const medH = median(lines.map((l) => l.h)) || 10;
  const minGutter = Math.max(2, medH * 0.3);
  const bounds = [];
  for (let x = 0; x < width;) {
    if (cov[x] <= allow) {
      let e = x;
      while (e < width && cov[e] <= allow) e++;
      if (x > 0 && e < width && e - x >= minGutter) bounds.push(x0 + (x + e) / 2);
      x = e;
    } else {
      x++;
    }
  }

  const ncol = bounds.length + 1;
  const colOf = (s) => {
    let best = 0, bestOv = -Infinity;
    for (let c = 0; c < ncol; c++) {
      const a = c === 0 ? -Infinity : bounds[c - 1];
      const b = c === ncol - 1 ? Infinity : bounds[c];
      const ov = Math.min(s.x1, b) - Math.max(s.x0, a);
      if (ov > bestOv) { bestOv = ov; best = c; }
    }
    return best;
  };

  let rows = lines.map((l) => {
    const cells = new Array(ncol).fill("");
    for (const s of l.segs) {
      const c = colOf(s);
      cells[c] = cells[c] ? cells[c] + " " + s.text : s.text;
    }
    return { cells, y: l.y, h: l.h };
  });

  if (opts.mergeWrapped) {
    const merged = [];
    for (const r of rows) {
      const prev = merged[merged.length - 1];
      const filled = r.cells.filter(Boolean).length;
      const isContinuation = prev && !r.cells[0] &&
        filled <= Math.max(1, Math.floor(ncol / 2)) &&
        r.y - prev.y <= Math.max(prev.h, r.h) * 1.6;
      if (isContinuation) {
        r.cells.forEach((c, i) => { if (c) prev.cells[i] = prev.cells[i] ? prev.cells[i] + " " + c : c; });
        prev.y = r.y;
      } else {
        merged.push({ cells: [...r.cells], y: r.y, h: r.h });
      }
    }
    rows = merged;
  }

  const keep = [];
  for (let c = 0; c < ncol; c++) if (rows.some((r) => r.cells[c])) keep.push(c);
  return rows.map((r) => keep.map((c) => r.cells[c]));
}

/* --------------------------------------------------------------- Excel */

async function runImport() {
  const selected = state.tables.filter((t) => t.selected);
  if (!selected.length) return;
  const mode = document.querySelector("input[name=outMode]:checked").value;
  const o = {
    addSource: $("addSource").checked,
    convert: $("convertNumbers").checked,
    asTable: $("asTable").checked,
  };

  state.busy = true;
  updateButtons();
  setStatus("Writing to Excel…");
  try {
    const count = await importToExcel(selected, mode, o);
    setStatus(`Imported ${selected.length} table(s) into ${count} sheet(s).`, "ok");
  } catch (e) {
    console.error(e);
    setStatus("Import failed: " + (e && e.message ? e.message : e), "error");
  } finally {
    state.busy = false;
    updateButtons();
  }
}

async function importToExcel(tables, mode, o) {
  return Excel.run(async (ctx) => {
    const sheets = ctx.workbook.worksheets;
    sheets.load("items/name");
    await ctx.sync();
    const used = new Set(sheets.items.map((s) => s.name.toLowerCase()));
    const created = [];
    const newSheet = (base) => {
      const name = uniqueSheetName(base, used);
      used.add(name.toLowerCase());
      const sh = sheets.add(name);
      created.push(sh);
      return sh;
    };

    if (mode === "sheet") {
      for (const t of tables) {
        const sh = newSheet(`${fileStem(t.file).slice(0, 20)} p${t.page}-${t.index}`);
        await writeBlock(ctx, sh, 0, withSource(t, o.addSource), o);
      }
    } else if (mode === "stack") {
      const sh = newSheet("PDF Tables");
      let r = 0;
      for (const t of tables) {
        const title = sh.getRangeByIndexes(r, 0, 1, 1);
        title.values = [[`${t.file}  |  page ${t.page}  |  table ${t.index}`]];
        title.format.font.bold = true;
        r += 1;
        const rows = withSource(t, o.addSource);
        await writeBlock(ctx, sh, r, rows, o);
        r += rows.length + 2;
      }
    } else {
      const sh = newSheet("PDF Combined");
      await writeBlock(ctx, sh, 0, combineTables(tables, o.addSource), o);
    }

    for (const sh of created) {
      sh.getUsedRange().format.autofitColumns();
    }
    if (created.length) created[0].activate();
    await ctx.sync();
    return created.length;
  });
}

/* -------------------------------------------- .xlsx download (any Excel) */

function readOutputOptions() {
  return {
    mode: document.querySelector("input[name=outMode]:checked").value,
    addSource: $("addSource").checked,
    convert: $("convertNumbers").checked,
    asTable: $("asTable").checked,
  };
}

function runDownload() {
  const selected = state.tables.filter((t) => t.selected);
  if (!selected.length) return;
  try {
    const o = readOutputOptions();
    const wb = buildWorkbook(selected, o.mode, o);
    const today = new Date().toISOString().slice(0, 10);
    XLSX.writeFile(wb, `PDF tables ${today}.xlsx`, { compression: true });
    setStatus(`Saved ${selected.length} table(s) to "PDF tables ${today}.xlsx" in your Downloads folder.`, "ok");
  } catch (e) {
    console.error(e);
    setStatus("Could not create the Excel file: " + (e && e.message ? e.message : e), "error");
  }
}

/** Builds a SheetJS workbook with the same layout as the in-Excel import. */
function buildWorkbook(tables, mode, o) {
  const wb = XLSX.utils.book_new();
  const used = new Set();
  const addSheet = (base, blocks, filter) => {
    const name = uniqueSheetName(base, used);
    used.add(name.toLowerCase());
    XLSX.utils.book_append_sheet(wb, blocksToSheet(blocks, o, filter), name);
  };

  if (mode === "sheet") {
    for (const t of tables) {
      addSheet(`${fileStem(t.file).slice(0, 20)} p${t.page}-${t.index}`,
        [{ rows: withSource(t, o.addSource) }], o.asTable);
    }
  } else if (mode === "stack") {
    addSheet("PDF Tables", tables.map((t) => ({
      title: `${t.file}  |  page ${t.page}  |  table ${t.index}`,
      rows: withSource(t, o.addSource),
    })), false);
  } else {
    addSheet("PDF Combined", [{ rows: combineTables(tables, o.addSource) }], o.asTable);
  }
  return wb;
}

function blocksToSheet(blocks, o, filter) {
  const ws = {};
  const widths = [];
  let r = 0;
  let maxC = 0;
  const put = (row, col, cell) => {
    ws[XLSX.utils.encode_cell({ r: row, c: col })] = cell;
    const len = String(cell.w || cell.v).length;
    widths[col] = Math.max(widths[col] || 8, Math.min(50, len + 2));
    maxC = Math.max(maxC, col);
  };

  for (const b of blocks) {
    if (b.title) { put(r, 0, { t: "s", v: b.title }); r++; }
    const start = r;
    b.rows.forEach((row, i) => {
      const asNumber = o.convert && !(o.asTable && i === 0); // keep header row as text
      row.forEach((c, col) => {
        const cell = xlsxCell(c, asNumber);
        if (cell) put(r, col, cell);
      });
      r++;
    });
    if (filter && b.rows.length >= 2) {
      const width = Math.max(...b.rows.map((x) => x.length));
      ws["!autofilter"] = { ref: XLSX.utils.encode_range({ s: { r: start, c: 0 }, e: { r: r - 1, c: width - 1 } }) };
    }
    r += 2; // blank rows between stacked tables
  }

  ws["!ref"] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: Math.max(0, r - 3), c: maxC } });
  ws["!cols"] = widths.map((w) => ({ wch: w || 8 }));
  return ws;
}

/** Cell object for SheetJS. Text stays text, so no apostrophe guard is needed here. */
function xlsxCell(raw, convert) {
  if (typeof raw === "number") return { t: "n", v: raw };
  const s = String(raw == null ? "" : raw).trim();
  if (!s) return null;
  if (convert) {
    const n = parseNumber(s);
    if (n !== null) return { t: "n", v: n };
    const pct = s.match(/^([-+]?\d+(?:\.\d+)?)\s?%$/);
    if (pct) {
      const decimals = (pct[1].split(".")[1] || "").length;
      return { t: "n", v: parseFloat(pct[1]) / 100, z: decimals ? "0." + "0".repeat(decimals) + "%" : "0%" };
    }
  }
  return { t: "s", v: s };
}

function withSource(t, addSource) {
  if (!addSource) return t.rows;
  return t.rows.map((r, i) => (i === 0 ? ["Source file", "Page"] : [t.file, t.page]).concat(r));
}

/** Stacks tables under one header row; repeated header rows (e.g. on every page) are dropped. */
function combineTables(tables, addSource) {
  const width = Math.max(...tables.map((t) => (t.rows[0] ? t.rows[0].length : 0)), 1);
  const norm = (r) => r.map((c) => String(c).trim().toLowerCase()).join("\u0001");
  const header = tables[0].rows[0] || [];
  const headerKey = norm(header);
  const pad = (r) => { const a = r.slice(); while (a.length < width) a.push(""); return a; };

  const out = [(addSource ? ["Source file", "Page"] : []).concat(pad(header))];
  tables.forEach((t, ti) => {
    t.rows.forEach((r, ri) => {
      if (ti === 0 && ri === 0) return;
      if (norm(r) === headerKey) return;
      out.push((addSource ? [t.file, t.page] : []).concat(pad(r)));
    });
  });
  return out;
}

async function writeBlock(ctx, sheet, startRow, rows, o) {
  if (!rows.length) return;
  const width = Math.max(1, ...rows.map((r) => r.length));
  const values = rows.map((r, i) => {
    const asNumber = o.convert && !(o.asTable && i === 0); // keep header row as text
    const a = r.map((c) => toCell(c, asNumber));
    while (a.length < width) a.push("");
    return a;
  });

  const CHUNK = 2000;
  for (let i = 0; i < values.length; i += CHUNK) {
    const part = values.slice(i, i + CHUNK);
    sheet.getRangeByIndexes(startRow + i, 0, part.length, width).values = part;
    await ctx.sync();
  }

  const range = sheet.getRangeByIndexes(startRow, 0, values.length, width);
  if (o.asTable && values.length >= 2) {
    try {
      sheet.tables.add(range, true);
      await ctx.sync();
    } catch (e) {
      console.warn("Could not create Excel table", e);
    }
  } else {
    range.getRow(0).format.font.bold = true;
  }
}

/** Converts a cell's text into what Excel should store. */
function toCell(raw, convert) {
  if (typeof raw === "number") return raw;
  const s = String(raw == null ? "" : raw).trim();
  if (!s) return "";
  if (convert) {
    const n = parseNumber(s);
    if (n !== null) return n;
    if (/^[-+]?\d+(\.\d+)?\s?%$/.test(s)) return s.replace(/\s/g, ""); // Excel turns "12.5%" into a percentage
  }
  // A leading apostrophe makes Excel keep the text as typed (no formulas, dates or lost zeros).
  return /^[=+\-@]/.test(s) || /^[\d\s.,/:%$€£¥()-]+$/.test(s) ? "'" + s : s;
}

function parseNumber(s) {
  let t = s.replace(/\s+/g, "");
  let neg = false;
  if (/^\(.*\)$/.test(t)) { neg = true; t = t.slice(1, -1); }
  if (t.startsWith("-") || t.startsWith("−")) { neg = !neg; t = t.slice(1); }
  else if (t.startsWith("+")) t = t.slice(1);
  t = t.replace(/^[$€£¥]/, "");
  if (t.endsWith("-")) { neg = !neg; t = t.slice(0, -1); } // trailing minus (accounting style)
  if (!/^(\d{1,3}(,\d{3})+|\d+)(\.\d+)?$|^\.\d+$/.test(t)) return null;
  if (/^0\d/.test(t)) return null;                       // keep IDs like 00123 as text
  if (t.replace(/[,.]/g, "").length > 15) return null;   // avoid precision loss on long IDs
  const v = parseFloat(t.replace(/,/g, ""));
  if (!isFinite(v)) return null;
  return neg ? -v : v;
}

function fileStem(name) {
  return name.replace(/\.pdf$/i, "");
}

function uniqueSheetName(base, used) {
  let b = base.replace(/[\\/?*[\]:]/g, "_").replace(/^'+|'+$/g, "").trim() || "Sheet";
  b = b.slice(0, 31);
  let name = b;
  let n = 2;
  while (used.has(name.toLowerCase()) || name.toLowerCase() === "history") {
    const suffix = ` (${n++})`;
    name = b.slice(0, 31 - suffix.length) + suffix;
  }
  return name;
}

// Exposed for testing in a normal browser console.
window.PdfTableImporter = { extractTablesFromPdf, combineTables, toCell, parseNumber, parsePages, buildWorkbook };
