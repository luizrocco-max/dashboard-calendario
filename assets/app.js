/* ============================================================================
   Dashboard Calendário — app.js
   Lê uma planilha (xlsx/xls/csv), detecta colunas automaticamente e monta
   calendário + KPIs + gráficos. Tudo client-side; nada sai do navegador.
   ========================================================================== */
'use strict';

/* ----------------------------------------------------------------- helpers */
const $ = (sel, ctx = document) => ctx.querySelector(sel);
const $$ = (sel, ctx = document) => Array.from(ctx.querySelectorAll(sel));
const el = (tag, cls, txt) => {
  const n = document.createElement(tag);
  if (cls) n.className = cls;
  if (txt != null) n.textContent = txt;
  return n;
};
const NF = new Intl.NumberFormat('pt-BR');
const norm = (s) => String(s == null ? '' : s)
  .normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim();

const PT_MONTHS = ['janeiro','fevereiro','março','abril','maio','junho','julho',
  'agosto','setembro','outubro','novembro','dezembro'];
const PT_MONTHS_ABBR = ['jan','fev','mar','abr','mai','jun','jul','ago','set','out','nov','dez'];
const PT_DOW = ['Dom','Seg','Ter','Qua','Qui','Sex','Sáb'];

const SERIES = ['--series-1','--series-2','--series-3','--series-4',
  '--series-5','--series-6','--series-7','--series-8'];
const cssVar = (name) => getComputedStyle(document.documentElement).getPropertyValue(name).trim();

/* --------------------------------------------------------------- date utils */
function mkDate(y, m, d) {                       // m is 1-based
  const dt = new Date(y, m - 1, d);
  return isNaN(dt.getTime()) ? null : dt;
}
function excelSerialToDate(serial) {
  const days = Math.floor(serial);
  const ms = Math.round((serial - days) * 86400000);
  const utc = (days - 25569) * 86400000 + ms;    // 25569 = days 1899-12-30 → 1970-01-01
  const d = new Date(utc);
  if (isNaN(d.getTime())) return null;
  return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
}
function ptMonthIndex(name) {
  const n = norm(name);
  let i = PT_MONTHS.findIndex(m => norm(m) === n);
  if (i >= 0) return i + 1;
  i = PT_MONTHS_ABBR.findIndex(m => n.startsWith(m));
  return i >= 0 ? i + 1 : 0;
}
function parseDate(v) {
  if (v == null || v === '') return null;
  if (v instanceof Date) return isNaN(v.getTime()) ? null : new Date(v.getFullYear(), v.getMonth(), v.getDate());
  if (typeof v === 'number' && isFinite(v)) {
    if (v > 20 && v < 2958466) return excelSerialToDate(v);
    return null;
  }
  const s = String(v).trim();
  if (!s) return null;
  let m;
  // ISO: yyyy-mm-dd
  if ((m = s.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/))) return mkDate(+m[1], +m[2], +m[3]);
  // Brazilian: dd/mm/yyyy, dd-mm-yy, dd.mm.yyyy
  if ((m = s.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})/))) {
    let dd = +m[1], mm = +m[2], yy = +m[3];
    if (yy < 100) yy += yy < 70 ? 2000 : 1900;
    if (mm > 12 && dd <= 12) { const t = dd; dd = mm; mm = t; }   // was mm/dd
    return mkDate(yy, mm, dd);
  }
  // "24 de julho de 2026" / "24 jul 2026"
  if ((m = s.match(/(\d{1,2})\s*(?:de\s+)?([a-zç]{3,})\.?\s*(?:de\s+)?(\d{4})/i))) {
    const mo = ptMonthIndex(m[2]);
    if (mo) return mkDate(+m[3], mo, +m[1]);
  }
  const t = Date.parse(s);
  if (!isNaN(t)) { const d = new Date(t); return mkDate(d.getFullYear(), d.getMonth() + 1, d.getDate()); }
  return null;
}
const dayKey = (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
const monthKey = (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
const today0 = () => { const n = new Date(); return new Date(n.getFullYear(), n.getMonth(), n.getDate()); };
const sameDay = (a, b) => a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate();
const addDays = (d, n) => { const x = new Date(d); x.setDate(x.getDate() + n); return x; };
function fmtDate(d) {
  if (!(d instanceof Date) || isNaN(d.getTime())) return '';
  return `${String(d.getDate()).padStart(2, '0')}/${String(d.getMonth() + 1).padStart(2, '0')}/${d.getFullYear()}`;
}
function fmtCell(v) {
  if (v == null) return '';
  if (v instanceof Date) return fmtDate(v);
  return String(v);
}
function relativeLabel(d) {
  const diff = Math.round((d - today0()) / 86400000);
  if (diff === 0) return 'hoje';
  if (diff === 1) return 'amanhã';
  if (diff === -1) return 'ontem';
  if (diff > 1) return `em ${diff} dias`;
  return `há ${Math.abs(diff)} dias`;
}

/* ------------------------------------------------------------- app state */
const state = {
  fileName: '',
  sheetNames: [],
  sheetName: '',
  aoa: [],            // array-of-arrays of the active sheet
  headerRow: 0,
  headers: [],        // column labels
  mapping: { date: -1, title: -1, category: -1, end: -1 },
  events: [],         // built events (all)
  categories: [],     // ordered distinct categories
  catColor: new Map(),
  activeCats: new Set(),
  search: '',
  view: today0(),     // month being viewed (any day within it)
};

const LS_DATA = 'dc_data_v1';
const LS_THEME = 'dc_theme';

/* -------------------------------------------------------- column detection */
const RX_DATE = /(^|\s)(data|date|dia|quando|in[ií]cio|inicio|vencimento|prazo|realiza|agenda)/i;
const RX_TITLE = /(evento|t[ií]tulo|titulo|nome|atividade|descri|assunto|compromisso|tarefa|item|resumo|a[çc][aã]o|acao)/i;
const RX_CAT_STRONG = /(categoria|^tipo|\btipo\b|status|situa|classe|prioridade|n[ií]vel|nivel|est[aá]gio|estagio|\bfase\b|modalidade)/i;
const RX_CAT_WEAK = /(respons|setor|[aá]rea|area|grupo|equipe|depart|segmento|cliente|local|sala|unidade)/i;
const RX_CAT = new RegExp(RX_CAT_STRONG.source + '|' + RX_CAT_WEAK.source, 'i');
const RX_END = /(fim|t[ée]rmino|termino|final|end|at[ée]|conclus)/i;

function detectHeaderRow(aoa) {
  const limit = Math.min(aoa.length, 12);
  let best = 0, bestScore = -Infinity;
  for (let i = 0; i < limit; i++) {
    const row = aoa[i] || [];
    const nonEmpty = row.filter(c => c != null && String(c).trim() !== '');
    if (!nonEmpty.length) continue;
    const strCount = nonEmpty.filter(c => typeof c === 'string').length;
    const next = aoa[i + 1] || [];
    const nextNonEmpty = next.filter(c => c != null && String(c).trim() !== '').length;
    let score = nonEmpty.length + strCount * 0.6 + (nextNonEmpty > 0 ? 3 : -4);
    if (score > bestScore) { bestScore = score; best = i; }
  }
  return best;
}
function columnValues(aoa, headerRow, colIdx, max = 300) {
  const out = [];
  for (let i = headerRow + 1; i < aoa.length && out.length < max; i++) {
    const v = aoa[i] ? aoa[i][colIdx] : null;
    if (v != null && String(v).trim() !== '') out.push(v);
  }
  return out;
}
function dateScore(vals) {
  if (!vals.length) return 0;
  let ok = 0;
  for (const v of vals) if (parseDate(v)) ok++;
  return ok / vals.length;
}
function autoMap(aoa, headerRow, headers) {
  const nCols = headers.length;
  const info = [];
  for (let c = 0; c < nCols; c++) {
    const vals = columnValues(aoa, headerRow, c);
    const distinct = new Set(vals.map(v => fmtCell(v)));
    const strFrac = vals.length ? vals.filter(v => typeof v === 'string' && parseDate(v) == null).length / vals.length : 0;
    info.push({
      c, label: headers[c] || `Coluna ${c + 1}`,
      count: vals.length,
      dateScore: dateScore(vals),
      distinct: distinct.size,
      distinctRatio: vals.length ? distinct.size / vals.length : 1,
      strFrac,
      avgLen: vals.length ? vals.reduce((s, v) => s + fmtCell(v).length, 0) / vals.length : 0,
    });
  }
  const used = new Set();
  const pick = (fn) => {
    let best = null;
    for (const x of info) {
      if (used.has(x.c)) continue;
      const s = fn(x);
      if (s > 0 && (!best || s > best.s)) best = { c: x.c, s };
    }
    if (best) { used.add(best.c); return best.c; }
    return -1;
  };
  // date: highest dateScore >= .5, boosted by name
  const date = pick(x => x.dateScore >= 0.5 ? x.dateScore + (RX_DATE.test(x.label) ? 0.5 : 0) : 0);
  // end date: another date column with matching name
  const end = pick(x => (x.dateScore >= 0.5 && RX_END.test(x.label)) ? x.dateScore + 1 : 0);
  // title: prefer name match, else most free-text-ish (high distinct ratio, decent length)
  const title = pick(x => {
    if (x.dateScore >= 0.5) return 0;
    let s = 0;
    if (RX_TITLE.test(x.label)) s += 5;
    s += x.distinctRatio * 2 + Math.min(x.avgLen / 12, 2);
    return s;
  });
  // category: low-ish cardinality text column, prefer name match (strong > weak)
  const category = pick(x => {
    if (x.dateScore >= 0.5) return 0;
    const strong = RX_CAT_STRONG.test(x.label), weak = RX_CAT_WEAK.test(x.label);
    if (x.distinct < 2 || x.distinct > 40) return (strong || weak) && x.distinct >= 2 ? 3 : 0;
    let s = (1 - x.distinctRatio) * 2 + 1;
    if (strong) s += 6; else if (weak) s += 2.5;
    return s;
  });
  return { date, title, category, end };
}

/* ----------------------------------------------------------- load pipeline */
function readWorkbook(data, isText) {
  return XLSX.read(data, { type: isText ? 'string' : 'array', cellDates: true, dateNF: 'yyyy-mm-dd' });
}
function sheetToAoa(wb, name) {
  const ws = wb.Sheets[name];
  return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: false });
}
function loadFile(file) {
  const reader = new FileReader();
  const isCsv = /\.csv$/i.test(file.name);
  reader.onload = (e) => {
    try {
      const wb = isCsv
        ? readWorkbook(e.target.result, true)
        : readWorkbook(new Uint8Array(e.target.result), false);
      const sheetNames = wb.SheetNames.filter(n => {
        const a = sheetToAoa(wb, n);
        return a.some(r => r && r.some(c => c != null && String(c).trim() !== ''));
      });
      if (!sheetNames.length) throw new Error('A planilha está vazia.');
      state.fileName = file.name;
      state._wb = wb;
      state.sheetNames = sheetNames;
      activateSheet(sheetNames[0], true);
      persist();
    } catch (err) {
      console.error(err);
      toast('Não consegui ler esse arquivo. Verifique se é uma planilha válida (.xlsx, .xls ou .csv).', true);
    }
  };
  reader.onerror = () => toast('Erro ao ler o arquivo.', true);
  if (isCsv) reader.readAsText(file); else reader.readAsArrayBuffer(file);
}
function activateSheet(name, autoView) {
  state.sheetName = name;
  state.aoa = sheetToAoa(state._wb, name);
  state.headerRow = null;
  state.mapping = null;
  processData(autoView);
}
// Decide between the matrix-calendar format and the generic tabular format, then build events.
function processData(autoView) {
  state.matrix = detectMatrix(state.aoa);
  state.year = detectYear(state.aoa, state.sheetNames, state.fileName);
  if (state.matrix) {
    state.mapping = null;
    buildEventsMatrix();
  } else {
    if (state.headerRow == null) state.headerRow = detectHeaderRow(state.aoa);
    rebuildHeaders();
    if (!state.mapping) state.mapping = autoMap(state.aoa, state.headerRow, state.headers);
    buildEvents();
  }
  if (autoView) pickInitialView();
  renderAll();
  showDashboard();
  if (!state.events.length) {
    $('#mapping').open = true;
    toast('Não identifiquei datas na planilha. Ajuste as colunas em “Configurar colunas e cabeçalho”.', true);
  }
}
function rebuildHeaders() {
  const hrow = state.aoa[state.headerRow] || [];
  const nCols = state.aoa.reduce((m, r) => Math.max(m, r ? r.length : 0), 0);
  const headers = [];
  for (let c = 0; c < nCols; c++) {
    const h = hrow[c];
    headers.push(h != null && String(h).trim() !== '' ? String(h).trim() : `Coluna ${c + 1}`);
  }
  state.headers = headers;
}

/* ------------------------------------------------------------ build events */
function buildEvents() {
  const { date: dc, title: tc, category: cc, end: ec } = state.mapping;
  const events = [];
  for (let i = state.headerRow + 1; i < state.aoa.length; i++) {
    const row = state.aoa[i];
    if (!row) continue;
    const d = dc >= 0 ? parseDate(row[dc]) : null;
    if (!d) continue;
    const end = ec >= 0 ? parseDate(row[ec]) : null;
    const title = tc >= 0 && row[tc] != null && String(row[tc]).trim() !== ''
      ? String(row[tc]).trim() : '(sem título)';
    const category = cc >= 0 && row[cc] != null && String(row[cc]).trim() !== ''
      ? String(row[cc]).trim() : null;
    const raw = {};
    state.headers.forEach((h, c) => {
      if (row[c] == null || String(row[c]).trim() === '') return;
      if (c === dc && d) raw[h] = fmtDate(d);           // always show a real date, never a serial
      else if (c === ec && end) raw[h] = fmtDate(end);
      else raw[h] = fmtCell(row[c]);
    });
    events.push({ date: d, end: end && end >= d ? end : null, title, category, raw });
  }
  events.sort((a, b) => a.date - b.date);
  state.events = events;
  buildCategories();
}
function buildCategories() {
  const freq = new Map();
  for (const e of state.events) {
    const c = e.category || '—';
    freq.set(c, (freq.get(c) || 0) + 1);
  }
  const cats = Array.from(freq.entries()).sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0])).map(x => x[0]);
  state.categories = cats;
  state.catColor = new Map();
  cats.forEach((c, i) => {
    state.catColor.set(c, i < SERIES.length ? cssVar(SERIES[i]) : cssVar('--series-other'));
  });
  // keep active set (default all on for new categories)
  const prev = state.activeCats;
  state.activeCats = new Set(cats.filter(c => prev.size === 0 || prev.has(c) || !prev._init));
  if (prev.size === 0) { cats.forEach(c => state.activeCats.add(c)); }
  state.activeCats._init = true;
}
const catColorOf = (e) => state.catColor.get(e.category || '—') || cssVar('--series-1');

function pickInitialView() {
  const t = today0();
  const hasThisMonth = state.events.some(e => monthKey(e.date) === monthKey(t));
  if (hasThisMonth || !state.events.length) { state.view = t; return; }
  // jump to the month of the next upcoming event, else the last event
  const upcoming = state.events.find(e => e.date >= t);
  state.view = new Date((upcoming || state.events[state.events.length - 1]).date);
}

/* ---------------------------------------------- matrix (grid) calendar format
   Formato "planilha de parede": colunas EVENTOS | DATA | CONFEDERAÇÃO + colunas
   de dias 1..31, com blocos por mês e "X" marcando os dias de cada evento. */
const PT_MONTH_IDX = {
  janeiro: 1, fevereiro: 2, marco: 3, abril: 4, maio: 5, junho: 6,
  julho: 7, agosto: 8, setembro: 9, outubro: 10, novembro: 11, dezembro: 12,
};
const clampDay = (y, m, d) => Math.min(Math.max(d, 1), new Date(y, m, 0).getDate());
const mkClamped = (y, m, d) => new Date(y, m - 1, clampDay(y, m, d));

function detectMatrix(aoa) {
  for (let i = 0; i < Math.min(aoa.length, 25); i++) {
    const row = aoa[i] || [];
    let eventCol = -1, dataCol = -1, confCol = -1;
    row.forEach((c, ci) => {
      const n = norm(c);
      if (n === 'eventos' && eventCol < 0) eventCol = ci;
      if (n === 'data' && dataCol < 0) dataCol = ci;
      if ((n === 'confederacao' || n === 'confederacoes') && confCol < 0) confCol = ci;
    });
    if (eventCol >= 0 && dataCol >= 0) {
      const dayCols = {};
      row.forEach((c, ci) => { const v = Number(c); if (Number.isInteger(v) && v >= 1 && v <= 31 && ci > dataCol) dayCols[ci] = v; });
      if (Object.keys(dayCols).length >= 20) return { headerRow: i, eventCol, dataCol, confCol, dayCols };
    }
  }
  return null;
}
function detectYear(aoa, sheetNames, fileName) {
  const hay = [fileName || '', (sheetNames || []).join(' '), (aoa.slice(0, 3).flat() || []).map(c => c == null ? '' : String(c)).join(' ')].join(' ');
  const m = hay.match(/20\d{2}/);
  return m ? +m[0] : new Date().getFullYear();
}
function matrixSpan(text, month, year, row, dayCols) {
  text = (text || '').trim();
  let m;
  if ((m = text.match(/(\d{1,2})\s*\/\s*(\d{1,2}).*?\ba\b.*?(\d{1,2})\s*\/\s*(\d{1,2})/i)))
    return { start: mkClamped(year, +m[2], +m[1]), end: mkClamped(year, +m[4], +m[3]) };
  if ((m = text.match(/^(\d{1,2})\s*\/\s*(\d{1,2})$/))) { const d = mkClamped(year, +m[2], +m[1]); return { start: d, end: d }; }
  if (month) {
    if ((m = text.match(/^(\d{1,2})\s*a\s*(\d{1,2})$/i))) return { start: mkClamped(year, month, +m[1]), end: mkClamped(year, month, +m[2]) };
    const nums = (text.match(/\d{1,2}/g) || []).map(Number).filter(n => n >= 1 && n <= 31);
    if (nums.length) return { start: mkClamped(year, month, Math.min(...nums)), end: mkClamped(year, month, Math.max(...nums)) };
  }
  const marks = [];
  for (const ci in dayCols) { const v = (row[ci] == null ? '' : String(row[ci]).trim()); if (v && v !== '·') marks.push(dayCols[ci]); }
  if (month && marks.length) return { start: mkClamped(year, month, Math.min(...marks)), end: mkClamped(year, month, Math.max(...marks)) };
  return null;
}
function buildEventsMatrix() {
  const aoa = state.aoa, year = state.year, L = state.matrix;
  const events = [];
  let month = null;
  const codes = new Set();          // confederation codes seen (for inferring blanks)
  for (let i = 0; i < aoa.length; i++) {
    const row = aoa[i] || [];
    for (const c of row) { const mo = PT_MONTH_IDX[norm(c)]; if (mo) { month = mo; break; } }
    if (norm(row[L.eventCol]) === 'eventos') continue;                 // repeated header row
    const title = row[L.eventCol] == null ? '' : String(row[L.eventCol]).trim();
    const dataText = row[L.dataCol] == null ? '' : String(row[L.dataCol]).trim();
    if (!title && !dataText) continue;                                 // weekday / pure month row
    const span = matrixSpan(dataText, month, year, row, L.dayCols);
    if (!span) continue;
    let category = L.confCol >= 0 && row[L.confCol] != null && String(row[L.confCol]).trim() !== '' ? String(row[L.confCol]).trim() : null;
    if (category) codes.add(category.toUpperCase());
    events.push({ _row: i, date: span.start, end: span.end > span.start ? span.end : null, title: title || '(sem título)', category, dataText, month });
  }
  // infer blank categories from a code appearing literally in the title
  events.forEach(e => {
    if (e.category) return;
    const up = (e.title || '').toUpperCase();
    for (const code of codes) { if (new RegExp('\\b' + code.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '\\b').test(up)) { e.category = code; break; } }
  });
  // dedupe cross-month repeats (same event drawn in two month blocks)
  const seen = new Set();
  const deduped = [];
  for (const e of events) {
    const k = [e.title, e.category || '', +e.date, +(e.end || e.date)].join('|');
    if (seen.has(k)) continue;
    seen.add(k);
    const monthName = e.month ? PT_MONTHS[e.month - 1] : '';
    const period = fmtDate(e.date) + (e.end ? ' – ' + fmtDate(e.end) : '');
    e.raw = { 'Data (planilha)': e.dataText || period, 'Período': period, 'Mês': monthName ? monthName.charAt(0).toUpperCase() + monthName.slice(1) : '' };
    deduped.push(e);
  }
  deduped.sort((a, b) => a.date - b.date);
  state.events = deduped;
  buildCategories();
}

/* -------------------------------------------------------------- filtering */
function filteredEvents() {
  const q = norm(state.search);
  return state.events.filter(e => {
    if (state.categories.length && !state.activeCats.has(e.category || '—')) return false;
    if (q) {
      const hay = norm(e.title + ' ' + Object.values(e.raw).join(' '));
      if (!hay.includes(q)) return false;
    }
    return true;
  });
}
/* expand multi-day events into per-day occurrences within a bounded range */
function eventsByDay(events) {
  const map = new Map();
  const push = (d, e) => { const k = dayKey(d); if (!map.has(k)) map.set(k, []); map.get(k).push(e); };
  for (const e of events) {
    const spanDays = e.end ? Math.round((e.end - e.date) / 86400000) + 1 : 1;
    if (e.end && spanDays > 20) { push(e.date, e); continue; }   // mês inteiro: mostra só no dia de início
    let d = new Date(e.date);
    const last = e.end ? new Date(Math.min(e.end.getTime(), addDays(e.date, 60).getTime())) : e.date;
    while (d <= last) { push(d, e); d = addDays(d, 1); }
  }
  return map;
}

/* ================================================================= RENDER */
function renderAll() {
  renderFileChip();
  renderSheetSelect();
  renderMapping();
  renderCatChips();
  renderKPIs();
  renderCalendar();
  renderCharts();
  renderUpcoming();
}

function renderFileChip() {
  $('#file-chip-text').innerHTML = `<strong>${escapeHtml(state.fileName || 'planilha')}</strong>` +
    (state.events.length ? ` · ${NF.format(state.events.length)} eventos` : '');
}
function renderSheetSelect() {
  const sel = $('#sheet-select');
  if (state.sheetNames.length > 1) {
    sel.classList.remove('hidden');
    sel.innerHTML = state.sheetNames.map(n => `<option ${n === state.sheetName ? 'selected' : ''}>${escapeHtml(n)}</option>`).join('');
  } else {
    sel.classList.add('hidden');
  }
}

function renderMapping() {
  const body = $('#mapping-body');
  body.innerHTML = '';
  const summary = $('#mapping > summary');

  // Matrix (grid) format: columns are fixed by detection — show a note + year override.
  if (state.matrix) {
    if (summary) summary.lastChild.textContent = ' Formato calendário (matriz) detectado — ajustar ano';
    const note = el('div', 'map-field');
    note.style.gridColumn = '1 / -1';
    note.innerHTML = `<label>Formato detectado automaticamente</label><div style="font-size:.84rem;color:var(--text-secondary)">` +
      `Sua planilha usa o formato de calendário em matriz (colunas <strong>EVENTOS · DATA · CONFEDERAÇÃO</strong> + dias 1–31, com “X” marcando os dias). ` +
      `As datas vêm da coluna <strong>DATA</strong> e as categorias da <strong>CONFEDERAÇÃO</strong>. Eventos que cruzam o mês são unificados.</div>`;
    const yr = el('div', 'map-field');
    yr.appendChild(el('label', null, 'Ano do calendário'));
    const yi = el('input');
    yi.type = 'number'; yi.min = '2000'; yi.max = '2100'; yi.value = String(state.year);
    yi.addEventListener('change', () => {
      const v = +yi.value; if (v >= 2000 && v <= 2100) { state.year = v; buildEventsMatrix(); pickInitialView(); renderAll(); persist(); }
    });
    yr.appendChild(yi);
    body.appendChild(note);
    body.appendChild(yr);
    return;
  }
  if (summary) summary.lastChild.textContent = ' Configurar colunas e cabeçalho';

  const opts = (selected, allowEmpty) => {
    let h = allowEmpty ? `<option value="-1"${selected < 0 ? ' selected' : ''}>— nenhuma —</option>` : '';
    state.headers.forEach((lbl, i) => { h += `<option value="${i}"${i === selected ? ' selected' : ''}>${escapeHtml(lbl)}</option>`; });
    return h;
  };
  const field = (key, label, allowEmpty) => {
    const wrap = el('div', 'map-field');
    wrap.appendChild(el('label', null, label));
    const s = el('select');
    s.innerHTML = opts(state.mapping[key], allowEmpty);
    s.addEventListener('change', () => { state.mapping[key] = +s.value; buildEvents(); renderAll(); persist(); });
    wrap.appendChild(s);
    return wrap;
  };
  // header row selector
  const hr = el('div', 'map-field');
  hr.appendChild(el('label', null, 'Linha do cabeçalho'));
  const hi = el('input');
  hi.type = 'number'; hi.min = '1'; hi.value = String(state.headerRow + 1);
  hi.addEventListener('change', () => {
    const v = Math.max(1, Math.min(+hi.value || 1, state.aoa.length)) - 1;
    state.headerRow = v; rebuildHeaders();
    state.mapping = autoMap(state.aoa, state.headerRow, state.headers);
    buildEvents(); renderAll(); persist();
  });
  hr.appendChild(hi);

  body.appendChild(field('date', 'Coluna de data', false));
  body.appendChild(field('title', 'Coluna de título/evento', false));
  body.appendChild(field('category', 'Coluna de categoria', true));
  body.appendChild(field('end', 'Coluna de data final (opcional)', true));
  body.appendChild(hr);
}

function renderCatChips() {
  let row = $('#cat-chips');
  if (!row) {
    row = el('div', 'chip-row');
    row.id = 'cat-chips';
    row.style.marginBottom = '18px';
    $('#kpi-row').before(row);
  }
  row.innerHTML = '';
  if (state.categories.length <= 1 && (state.categories[0] === '—' || !state.categories.length)) {
    row.classList.add('hidden'); return;
  }
  row.classList.remove('hidden');
  const allOn = state.categories.every(c => state.activeCats.has(c));
  const toggleAll = el('button', 'filter-chip');
  toggleAll.textContent = allOn ? 'Limpar filtros' : 'Selecionar todos';
  toggleAll.addEventListener('click', () => {
    if (allOn) state.categories.forEach(c => state.activeCats.delete(c));
    else state.categories.forEach(c => state.activeCats.add(c));
    renderAll();
  });
  row.appendChild(toggleAll);
  state.categories.forEach(c => {
    const on = state.activeCats.has(c);
    const chip = el('button', 'filter-chip');
    chip.setAttribute('aria-pressed', on ? 'true' : 'false');
    const sw = el('span', 'swatch'); sw.style.background = state.catColor.get(c);
    chip.appendChild(sw);
    chip.appendChild(document.createTextNode(c === '—' ? 'Sem categoria' : c));
    chip.addEventListener('click', () => {
      if (state.activeCats.has(c)) state.activeCats.delete(c); else state.activeCats.add(c);
      renderAll();
    });
    row.appendChild(chip);
  });
}

/* ------------------------------------------------------------------ KPIs */
function renderKPIs() {
  const evs = filteredEvents();
  const t = today0();
  const inView = evs.filter(e => monthKey(e.date) === monthKey(state.view));
  const next7 = evs.filter(e => e.date >= t && e.date <= addDays(t, 7));
  // busiest day
  const byDay = new Map();
  evs.forEach(e => { const k = dayKey(e.date); byDay.set(k, (byDay.get(k) || 0) + 1); });
  let busyKey = null, busyN = 0;
  byDay.forEach((n, k) => { if (n > busyN) { busyN = n; busyKey = k; } });
  const busyDate = busyKey ? new Date(busyKey + 'T00:00:00') : null;

  const monthLbl = state.view.toLocaleDateString('pt-BR', { month: 'long', year: 'numeric' });
  const kpis = [
    { icon: 'calendar', label: 'Total de eventos', value: NF.format(evs.length), sub: state.events.length !== evs.length ? `de ${NF.format(state.events.length)} no total` : 'na planilha' },
    { icon: 'month', label: 'No mês exibido', value: NF.format(inView.length), sub: monthLbl },
    { icon: 'clock', label: 'Próximos 7 dias', value: NF.format(next7.length), sub: next7.length ? 'a partir de hoje' : 'nada agendado' },
    { icon: 'star', label: 'Dia mais movimentado', value: busyN ? NF.format(busyN) : '—', sub: busyDate ? `${busyDate.getDate()} de ${PT_MONTHS[busyDate.getMonth()]}` : '' },
  ];
  const icons = {
    calendar: '<rect x="3" y="4" width="18" height="18" rx="2"/><path d="M16 2v4M8 2v4M3 10h18"/>',
    month: '<rect x="3" y="4" width="18" height="18" rx="2"/><path d="M3 10h18M8 2v4M16 2v4"/><rect x="7" y="14" width="4" height="4" rx="1" fill="currentColor" stroke="none"/>',
    clock: '<circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/>',
    star: '<path d="m12 2 3 6.3 6.9 1-5 4.9 1.2 6.8L12 17.8 5.9 21l1.2-6.8-5-4.9 6.9-1z"/>',
  };
  const row = $('#kpi-row');
  row.innerHTML = '';
  kpis.forEach(k => {
    const c = el('div', 'kpi');
    c.innerHTML =
      `<div class="kpi-label"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">${icons[k.icon]}</svg>${k.label}</div>` +
      `<div class="kpi-value">${k.value}</div>` +
      `<div class="kpi-sub">${escapeHtml(k.sub)}</div>`;
    row.appendChild(c);
  });
}

/* -------------------------------------------------------------- calendar */
function renderCalendar() {
  // weekday header
  const dow = $('#cal-dow');
  if (!dow.childElementCount) PT_DOW.forEach(d => dow.appendChild(el('div', 'cal-dow', d)));

  $('#cal-title').textContent = state.view.toLocaleDateString('pt-BR', { month: 'long', year: 'numeric' });

  const grid = $('#cal-grid');
  grid.innerHTML = '';
  const year = state.view.getFullYear(), month = state.view.getMonth();
  const first = new Date(year, month, 1);
  const lead = first.getDay();                    // 0 = domingo
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  const totalCells = Math.ceil((lead + daysInMonth) / 7) * 7;
  const start = addDays(first, -lead);
  const byDay = eventsByDay(filteredEvents());
  const t = today0();

  for (let i = 0; i < totalCells; i++) {
    const d = addDays(start, i);
    const inMonth = d.getMonth() === month;
    const cell = el('button', 'cal-cell');
    if (!inMonth) cell.classList.add('other-month');
    if (sameDay(d, t)) cell.classList.add('today');
    cell.appendChild(el('span', 'cal-daynum', String(d.getDate())));

    const evs = byDay.get(dayKey(d)) || [];
    const box = el('div', 'cal-events');
    const shown = evs.slice(0, 3);
    shown.forEach(e => {
      const chip = el('div', 'cal-ev');
      chip.style.setProperty('--ev-color', catColorOf(e));
      const txt = el('span', 'cal-ev-text', e.title);
      chip.appendChild(txt);
      chip.title = `${e.title}${e.category ? ' · ' + e.category : ''}`;
      box.appendChild(chip);
    });
    if (evs.length > shown.length) box.appendChild(el('div', 'cal-more', `+${evs.length - shown.length} mais`));
    cell.appendChild(box);
    if (evs.length) cell.addEventListener('click', () => openDay(d, evs));
    else cell.classList.add('empty-cell');
    grid.appendChild(cell);
  }
}

/* ----------------------------------------------------------------- charts */
const SVGNS = 'http://www.w3.org/2000/svg';
const svgEl = (tag, attrs) => {
  const n = document.createElementNS(SVGNS, tag);
  for (const k in attrs) n.setAttribute(k, attrs[k]);
  return n;
};

function renderCharts() {
  const evs = filteredEvents();
  renderMonthlyChart(evs);
  if (state.categories.length && !(state.categories.length === 1 && state.categories[0] === '—')) {
    $('#chart-percat').closest('.card').querySelector('h3').textContent = 'Eventos por categoria';
    renderCategoryChart(evs);
  } else {
    $('#chart-percat').closest('.card').querySelector('h3').textContent = 'Eventos por dia da semana';
    renderWeekdayChart(evs);
  }
}

/* vertical bars — magnitude over time (single blue hue) */
function renderMonthlyChart(evs) {
  const host = $('#chart-permes');
  host.innerHTML = '';
  if (!evs.length) { host.appendChild(emptyNote('Sem eventos para exibir.')); return; }
  // group by year-month, keep chronological, cap to last 12 buckets that have data
  const map = new Map();
  evs.forEach(e => { const k = monthKey(e.date); map.set(k, (map.get(k) || 0) + 1); });
  let keys = Array.from(map.keys()).sort();
  // fill gaps between first and last month
  if (keys.length > 1) {
    const filled = [];
    let [y, m] = keys[0].split('-').map(Number);
    const [ey, em] = keys[keys.length - 1].split('-').map(Number);
    while (y < ey || (y === ey && m <= em)) {
      filled.push(`${y}-${String(m).padStart(2, '0')}`);
      m++; if (m > 12) { m = 1; y++; }
      if (filled.length > 36) break;
    }
    keys = filled;
  }
  if (keys.length > 12) keys = keys.slice(keys.length - 12);
  const data = keys.map(k => {
    const [y, m] = k.split('-').map(Number);
    return { key: k, label: `${PT_MONTHS_ABBR[m - 1]}/${String(y).slice(2)}`, full: `${PT_MONTHS[m - 1]} de ${y}`, value: map.get(k) || 0 };
  });
  $('#permes-sub').textContent = `${data.length} ${data.length === 1 ? 'mês' : 'meses'}`;

  const W = 620, H = 240, padL = 30, padR = 12, padT = 16, padB = 30;
  const iw = W - padL - padR, ih = H - padT - padB;
  const max = Math.max(1, ...data.map(d => d.value));
  const svg = svgEl('svg', { viewBox: `0 0 ${W} ${H}`, role: 'img', 'aria-label': 'Eventos por mês' });

  // gridlines + y labels (0, mid, max)
  const ticks = niceTicks(max, 4);
  ticks.forEach(tk => {
    const y = padT + ih - (tk / max) * ih;
    svg.appendChild(svgEl('line', { class: 'gridline', x1: padL, y1: y, x2: W - padR, y2: y }));
    const lab = svgEl('text', { class: 'axis-label', x: padL - 6, y: y + 3, 'text-anchor': 'end' });
    lab.textContent = NF.format(tk); svg.appendChild(lab);
  });
  svg.appendChild(svgEl('line', { class: 'baseline', x1: padL, y1: padT + ih, x2: W - padR, y2: padT + ih }));

  const slot = iw / data.length;
  const bw = Math.min(46, slot * 0.62);
  const color = cssVar('--seq-450');
  data.forEach((d, i) => {
    const x = padL + slot * i + (slot - bw) / 2;
    const h = (d.value / max) * ih;
    const y = padT + ih - h;
    const bar = svgEl('rect', { x, y, width: bw, height: Math.max(h, d.value ? 2 : 0), rx: 4, fill: color });
    bar.style.cursor = 'default';
    const tt = `${d.full}: ${NF.format(d.value)} evento${d.value === 1 ? '' : 's'}`;
    bar.appendChild(svgTitle(tt));
    attachTip(bar, () => tt);
    svg.appendChild(bar);
    // value on top when few bars
    if (data.length <= 12 && d.value) {
      const vl = svgEl('text', { class: 'val-label', x: x + bw / 2, y: y - 5, 'text-anchor': 'middle' });
      vl.textContent = NF.format(d.value); svg.appendChild(vl);
    }
    // x label (thin out if many)
    if (data.length <= 12 || i % 2 === 0) {
      const xl = svgEl('text', { class: 'bar-label', x: x + bw / 2, y: H - 10, 'text-anchor': 'middle' });
      xl.textContent = d.label; svg.appendChild(xl);
    }
  });
  host.appendChild(svg);
}

/* horizontal meter list — identity (per-category colors, direct labels) */
function renderCategoryChart(evs) {
  const host = $('#chart-percat');
  host.innerHTML = '';
  if (!evs.length) { host.appendChild(emptyNote('Sem eventos para exibir.')); return; }
  const map = new Map();
  evs.forEach(e => { const c = e.category || '—'; map.set(c, (map.get(c) || 0) + 1); });
  let rows = Array.from(map.entries()).map(([c, v]) => ({ cat: c, value: v, color: state.catColor.get(c) || cssVar('--series-other') }))
    .sort((a, b) => b.value - a.value);
  // fold overflow into "Outras"
  if (rows.length > 9) {
    const head = rows.slice(0, 8);
    const rest = rows.slice(8).reduce((s, r) => s + r.value, 0);
    head.push({ cat: 'Outras', value: rest, color: cssVar('--series-other') });
    rows = head;
  }
  $('#percat-sub').textContent = `${state.categories.filter(c => state.activeCats.has(c)).length} de ${state.categories.length}`;
  const max = Math.max(1, ...rows.map(r => r.value));
  const total = rows.reduce((s, r) => s + r.value, 0);
  buildMeterList(host, rows.map(r => ({
    label: r.cat === '—' ? 'Sem categoria' : r.cat, value: r.value, color: r.color, max, total,
  })));
}
function renderWeekdayChart(evs) {
  const host = $('#chart-percat');
  host.innerHTML = '';
  if (!evs.length) { host.appendChild(emptyNote('Sem eventos para exibir.')); return; }
  const counts = new Array(7).fill(0);
  evs.forEach(e => counts[e.date.getDay()]++);
  const names = ['Domingo', 'Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado'];
  const max = Math.max(1, ...counts);
  const total = counts.reduce((a, b) => a + b, 0);
  $('#percat-sub').textContent = '';
  buildMeterList(host, counts.map((v, i) => ({ label: names[i], value: v, color: cssVar('--seq-450'), max, total })));
}
function buildMeterList(host, rows) {
  const list = el('div');
  list.style.display = 'flex'; list.style.flexDirection = 'column'; list.style.gap = '13px';
  rows.forEach(r => {
    const item = el('div');
    const top = el('div');
    top.style.display = 'flex'; top.style.justifyContent = 'space-between';
    top.style.alignItems = 'baseline'; top.style.marginBottom = '5px'; top.style.gap = '10px';
    const left = el('div');
    left.style.display = 'flex'; left.style.alignItems = 'center'; left.style.gap = '7px';
    left.style.minWidth = '0';
    const sw = el('span'); sw.style.cssText = `width:10px;height:10px;border-radius:3px;flex:none;background:${r.color}`;
    const nm = el('span', null, r.label);
    nm.style.cssText = 'font-size:.86rem;white-space:nowrap;overflow:hidden;text-overflow:ellipsis';
    left.appendChild(sw); left.appendChild(nm);
    const val = el('span');
    val.style.cssText = 'font-weight:700;font-size:.86rem;flex:none';
    const pct = r.total ? Math.round((r.value / r.total) * 100) : 0;
    val.innerHTML = `${NF.format(r.value)} <span style="color:var(--text-muted);font-weight:500;font-size:.78rem">${pct}%</span>`;
    top.appendChild(left); top.appendChild(val);
    const track = el('div');
    track.style.cssText = 'height:8px;border-radius:5px;background:var(--surface-2);overflow:hidden';
    const fill = el('div');
    fill.style.cssText = `height:100%;border-radius:5px;background:${r.color};width:${(r.value / r.max) * 100}%;min-width:${r.value ? '4px' : '0'}`;
    track.appendChild(fill);
    item.appendChild(top); item.appendChild(track);
    item.title = `${r.label}: ${NF.format(r.value)} (${pct}%)`;
    list.appendChild(item);
  });
  host.appendChild(list);
}
function niceTicks(max, n) {
  const step = niceNum(max / n);
  const ticks = [];
  for (let v = 0; v <= max + 1e-9; v += step) ticks.push(Math.round(v));
  if (ticks[ticks.length - 1] < max) ticks.push(Math.ceil(max / step) * step);
  return Array.from(new Set(ticks));
}
function niceNum(x) {
  if (x <= 1) return 1;
  const exp = Math.floor(Math.log10(x));
  const f = x / Math.pow(10, exp);
  const nf = f < 1.5 ? 1 : f < 3 ? 2 : f < 7 ? 5 : 10;
  return nf * Math.pow(10, exp);
}
const svgTitle = (t) => { const n = svgEl('title'); n.textContent = t; return n; };

/* --------------------------------------------------------------- upcoming */
function renderUpcoming() {
  const host = $('#upcoming');
  host.innerHTML = '';
  const t = today0();
  const up = filteredEvents().filter(e => e.date >= t).slice(0, 8);
  if (!up.length) { host.appendChild(emptyNote('Nenhum evento a partir de hoje.')); return; }
  up.forEach(e => {
    const item = el('div', 'up-item');
    const date = el('div', 'up-date');
    date.innerHTML = `<div class="d">${e.date.getDate()}</div><div class="m">${PT_MONTHS_ABBR[e.date.getMonth()]}</div>`;
    const main = el('div', 'up-main');
    main.appendChild(el('div', 'up-title', e.title));
    const meta = el('div', 'up-meta');
    if (e.category) {
      const tag = el('span', 'up-tag');
      const sw = el('span', 'swatch'); sw.style.background = catColorOf(e);
      tag.appendChild(sw); tag.appendChild(document.createTextNode(e.category));
      meta.appendChild(tag);
    }
    const loc = firstExtra(e);
    if (loc) { if (e.category) meta.appendChild(document.createTextNode('·')); meta.appendChild(document.createTextNode(loc)); }
    main.appendChild(meta);
    const rel = el('div', 'up-rel', relativeLabel(e.date));
    item.appendChild(date); item.appendChild(main); item.appendChild(rel);
    item.style.cursor = 'pointer';
    item.addEventListener('click', () => openDay(e.date, filteredEvents().filter(x => sameDay(x.date, e.date) || (x.end && x.date <= e.date && x.end >= e.date))));
    host.appendChild(item);
  });
}
function firstExtra(e) {
  if (!state.mapping) return e.dataText || (e.raw ? Object.values(e.raw)[0] : '') || '';
  const used = new Set();
  [state.mapping.date, state.mapping.title, state.mapping.category, state.mapping.end].forEach(i => { if (i >= 0) used.add(state.headers[i]); });
  for (const [k, v] of Object.entries(e.raw)) if (!used.has(k)) return v;
  return '';
}

/* ----------------------------------------------------------------- drawer */
function openDay(d, evs) {
  $('#drawer-title').textContent = `${d.getDate()} de ${PT_MONTHS[d.getMonth()]}`;
  $('#drawer-sub').textContent = `${d.toLocaleDateString('pt-BR', { weekday: 'long' })} · ${NF.format(evs.length)} evento${evs.length === 1 ? '' : 's'}`;
  const body = $('#drawer-body');
  body.innerHTML = '';
  evs.forEach(e => {
    const card = el('div', 'ev-card');
    card.style.setProperty('--ev-color', catColorOf(e));
    card.appendChild(el('h4', null, e.title));
    if (e.category) {
      const tag = el('div', 'ev-tag');
      const sw = el('span', 'swatch'); sw.style.background = catColorOf(e);
      tag.appendChild(sw); tag.appendChild(document.createTextNode(e.category));
      card.appendChild(tag);
    }
    const titleHeader = state.mapping && state.mapping.title >= 0 ? state.headers[state.mapping.title] : null;
    const catHeader = state.mapping && state.mapping.category >= 0 ? state.headers[state.mapping.category] : null;
    Object.entries(e.raw).forEach(([k, v]) => {
      if (k === titleHeader || k === catHeader) return;
      const f = el('div', 'ev-field');
      f.appendChild(el('span', 'k', k));
      f.appendChild(el('span', 'v', v));
      card.appendChild(f);
    });
    body.appendChild(card);
  });
  $('#drawer').classList.add('open');
  $('#drawer').setAttribute('aria-hidden', 'false');
  $('#drawer-backdrop').classList.add('open');
}
function closeDrawer() {
  $('#drawer').classList.remove('open');
  $('#drawer').setAttribute('aria-hidden', 'true');
  $('#drawer-backdrop').classList.remove('open');
}

/* ------------------------------------------------------------------ tips */
const tipEl = () => $('#tooltip');
function attachTip(node, textFn) {
  node.addEventListener('mouseenter', () => { const t = tipEl(); t.textContent = textFn(); t.classList.add('show'); });
  node.addEventListener('mousemove', (e) => { const t = tipEl(); t.style.left = e.clientX + 'px'; t.style.top = e.clientY + 'px'; });
  node.addEventListener('mouseleave', () => tipEl().classList.remove('show'));
}

/* ------------------------------------------------------------- utilities */
function emptyNote(msg) { return el('div', 'empty-note', msg); }
function escapeHtml(s) { return String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }
let toastTimer;
function toast(msg, isError) {
  const t = $('#toast');
  t.textContent = msg; t.classList.toggle('error', !!isError); t.classList.add('show');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => t.classList.remove('show'), 4200);
}

function showDashboard() { $('#empty-state').classList.add('hidden'); $('#dashboard').classList.remove('hidden'); }
function showEmpty() { $('#dashboard').classList.add('hidden'); $('#empty-state').classList.remove('hidden'); }

/* ---------------------------------------------------------- persistence */
function persist() {
  try {
    const payload = {
      fileName: state.fileName, sheetNames: state.sheetNames, sheetName: state.sheetName,
      aoa: state.aoa.map(r => (r || []).map(c => c instanceof Date ? { __d: c.getTime() } : c)),
      headerRow: state.headerRow, mapping: state.mapping,
    };
    localStorage.setItem(LS_DATA, JSON.stringify(payload));
  } catch (e) { /* quota or serialization — ignore, session stays in-memory */ }
}
function restore() {
  let raw;
  try { raw = localStorage.getItem(LS_DATA); } catch (e) { return false; }
  if (!raw) return false;
  try {
    const p = JSON.parse(raw);
    state.fileName = p.fileName; state.sheetNames = p.sheetNames || [p.sheetName];
    state.sheetName = p.sheetName;
    state.aoa = p.aoa.map(r => (r || []).map(c => (c && typeof c === 'object' && '__d' in c) ? new Date(c.__d) : c));
    state.headerRow = p.headerRow != null ? p.headerRow : null;
    state.mapping = p.mapping || null;
    state._wb = null;
    processData(true);
    return true;
  } catch (e) { return false; }
}
function clearData() {
  try { localStorage.removeItem(LS_DATA); } catch (e) {}
  Object.assign(state, { fileName: '', sheetNames: [], sheetName: '', aoa: [], events: [], categories: [], _wb: null });
  state.activeCats = new Set();
  showEmpty();
}

/* ------------------------------------------------------------------ theme */
function applyTheme(mode) {
  if (mode === 'light' || mode === 'dark') document.documentElement.setAttribute('data-theme', mode);
  else document.documentElement.removeAttribute('data-theme');
  const dark = document.documentElement.getAttribute('data-theme') === 'dark' ||
    (!document.documentElement.hasAttribute('data-theme') && matchMedia('(prefers-color-scheme: dark)').matches);
  $('.ic-sun').classList.toggle('hidden', !dark);
  $('.ic-moon').classList.toggle('hidden', dark);
  // recolor categories (css vars differ per theme) and re-render charts/chips
  if (state.events.length) { buildCategories(); renderAll(); }
}
function initTheme() {
  let saved;
  try { saved = localStorage.getItem(LS_THEME); } catch (e) {}
  applyTheme(saved || 'auto');
}
function toggleTheme() {
  const dark = document.documentElement.getAttribute('data-theme') === 'dark' ||
    (!document.documentElement.hasAttribute('data-theme') && matchMedia('(prefers-color-scheme: dark)').matches);
  const next = dark ? 'light' : 'dark';
  try { localStorage.setItem(LS_THEME, next); } catch (e) {}
  applyTheme(next);
}

/* -------------------------------------------------------------- sample data */
function loadSample() {
  const base = new Date(2026, 6, 1);                // julho/2026
  const cats = ['Reunião', 'Entrega', 'Treinamento', 'Viagem', 'Feriado'];
  const resp = ['Ana', 'Bruno', 'Carla', 'Diego', 'Equipe'];
  const locais = ['Sala 3', 'Online', 'Matriz', 'Cliente', 'Auditório'];
  const titles = {
    'Reunião': ['Reunião de planejamento', 'Alinhamento semanal', 'Comitê de projetos', 'Review de resultados'],
    'Entrega': ['Entrega do relatório', 'Fechamento mensal', 'Envio da proposta', 'Deploy da versão'],
    'Treinamento': ['Treinamento de segurança', 'Workshop de dados', 'Onboarding', 'Capacitação'],
    'Viagem': ['Visita técnica', 'Viagem comercial', 'Feira do setor'],
    'Feriado': ['Feriado nacional', 'Recesso'],
  };
  const rows = [['Data', 'Evento', 'Categoria', 'Responsável', 'Local']];
  let seed = 7;
  const rnd = () => { seed = (seed * 9301 + 49297) % 233280; return seed / 233280; };
  for (let m = 0; m < 3; m++) {
    const n = 12 + Math.floor(rnd() * 6);
    for (let i = 0; i < n; i++) {
      const day = 1 + Math.floor(rnd() * 27);
      const cat = cats[Math.floor(rnd() * cats.length)];
      const opts = titles[cat];
      rows.push([
        new Date(base.getFullYear(), base.getMonth() + m, day),
        opts[Math.floor(rnd() * opts.length)],
        cat,
        resp[Math.floor(rnd() * resp.length)],
        locais[Math.floor(rnd() * locais.length)],
      ]);
    }
  }
  const ws = XLSX.utils.aoa_to_sheet(rows, { cellDates: true });
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Eventos');
  state.fileName = 'exemplo.xlsx'; state._wb = wb; state.sheetNames = ['Eventos'];
  activateSheet('Eventos', true); persist();
  toast('Carregado com dados de exemplo. Troque pela sua planilha quando quiser.');
}

/* ------------------------------------------------------------------- wiring */
function chooseFile() { $('#file-input').click(); }
function initEvents() {
  $('#file-input').addEventListener('change', (e) => { if (e.target.files[0]) loadFile(e.target.files[0]); e.target.value = ''; });
  $('#btn-choose').addEventListener('click', chooseFile);
  $('#btn-load').addEventListener('click', chooseFile);
  $('#btn-replace').addEventListener('click', chooseFile);
  $('#btn-sample').addEventListener('click', loadSample);
  $('#btn-theme').addEventListener('click', toggleTheme);
  $('#search').addEventListener('input', (e) => { state.search = e.target.value; renderKPIs(); renderCalendar(); renderCharts(); renderUpcoming(); });
  $('#sheet-select').addEventListener('change', (e) => { activateSheet(e.target.value, true); persist(); });
  $('#cal-prev').addEventListener('click', () => { state.view = new Date(state.view.getFullYear(), state.view.getMonth() - 1, 1); renderCalendar(); renderKPIs(); });
  $('#cal-next').addEventListener('click', () => { state.view = new Date(state.view.getFullYear(), state.view.getMonth() + 1, 1); renderCalendar(); renderKPIs(); });
  $('#cal-today').addEventListener('click', () => { state.view = today0(); renderCalendar(); renderKPIs(); });
  $('#drawer-close').addEventListener('click', closeDrawer);
  $('#drawer-backdrop').addEventListener('click', closeDrawer);
  document.addEventListener('keydown', (e) => { if (e.key === 'Escape') closeDrawer(); });

  // drag & drop on empty state
  const drop = $('#empty-state');
  ['dragenter', 'dragover'].forEach(ev => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.add('dragover'); }));
  ['dragleave', 'drop'].forEach(ev => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.remove('dragover'); }));
  drop.addEventListener('drop', (e) => { const f = e.dataTransfer.files[0]; if (f) loadFile(f); });
  // also accept a drop anywhere once data is loaded
  window.addEventListener('dragover', (e) => e.preventDefault());
  window.addEventListener('drop', (e) => {
    if (e.target.closest('#empty-state')) return;
    e.preventDefault();
    const f = e.dataTransfer && e.dataTransfer.files[0];
    if (f && /\.(xlsx|xls|csv|xlsm)$/i.test(f.name)) loadFile(f);
  });
}

/* --------------------------------------------------------------------- init */
initTheme();
initEvents();
restore();
