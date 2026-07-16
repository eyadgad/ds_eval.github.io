'use strict';

/* ============================================================
   DS Eval · Rubric Scorer
   Vanilla JS, no framework. Excel via SheetJS (lazy-loaded).
   ============================================================ */

const CATEGORIES = ['Content', 'Execution', 'Presentation & Formatting'];
const LS_KEY = 'ds_eval_rubric_v1';
const XLSX_LOCAL = 'vendor/xlsx.full.min.js';
const XLSX_CDN = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
const JSZIP_LOCAL = 'vendor/jszip.min.js';
const JSZIP_CDN = 'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js';
const MAX_MATCH_FILES = 4;     // how many candidate files to list under each rubric point
const MIN_WEIGHT = -10;
const MAX_WEIGHT = 10;

const state = {
  items: [],
  filters: { search: '', category: 'all', grok: 'all', claude: 'all' }
};

/* ---------- id ---------- */
let _seq = 0;
const uid = () => 'r' + Date.now().toString(36) + (_seq++).toString(36);

/* ---------- normalization ---------- */
function normCategory(c) {
  if (c == null) return 'Content';
  const s = String(c).trim().toLowerCase();
  if (s.startsWith('content')) return 'Content';
  if (s.startsWith('exec')) return 'Execution';
  if (s.startsWith('present') || s.includes('format')) return 'Presentation & Formatting';
  const hit = CATEGORIES.find((x) => x.toLowerCase() === s);
  return hit || 'Content';
}
function normWeight(w) {
  let n = parseInt(w, 10);
  if (isNaN(n)) n = 1;
  return Math.max(MIN_WEIGHT, Math.min(MAX_WEIGHT, n));
}
function normPass(v) {
  if (v === 1 || v === 0) return v;
  const s = String(v == null ? '' : v).trim().toLowerCase();
  if (['1', 'pass', 'p', 'true', 'yes', 'y', '✓', '✔'].includes(s)) return 1;
  // everything else (including blank / unknown) defaults to fail (red)
  return 0;
}
function makeItem(o) {
  o = o || {};
  return {
    id: uid(),
    text: o.text != null ? String(o.text) : '',
    category: normCategory(o.category),
    weight: normWeight(o.weight),
    grok: normPass(o.grok),
    claude: normPass(o.claude)
  };
}

/* ---------- JSON <-> model ---------- */
function itemsFromJSON(obj) {
  const arr = obj && Array.isArray(obj.items) ? obj.items : Array.isArray(obj) ? obj : [];
  return arr.map((it) => {
    const ca = it.criteriaAnnotations || it.criteria || {};
    return makeItem({
      text: it.text != null ? it.text : it.rule || it.description || '',
      category: ca.Rule_Category != null ? ca.Rule_Category : it.category || it.Rule_Category,
      weight: ca.Weight != null ? ca.Weight : it.weight || it.Weight,
      grok: it.grok != null ? it.grok : it.Grok,
      claude: it.claude != null ? it.claude : it.Claude
    });
  });
}
function itemsToJSON() {
  return {
    items: state.items.map((it, idx) => ({
      id: 'C' + (idx + 1),
      text: it.text,
      criteriaAnnotations: { Rule_Category: it.category, Weight: String(it.weight) },
      grok: it.grok === 1 ? '1' : '0',
      claude: it.claude === 1 ? '1' : '0'
    }))
  };
}

/* ---------- persistence ---------- */
function save() {
  try { localStorage.setItem(LS_KEY, JSON.stringify({ items: state.items })); } catch (e) {}
}
function loadSaved() {
  try {
    const r = JSON.parse(localStorage.getItem(LS_KEY));
    if (r && Array.isArray(r.items)) return r.items.map(makeItem);
  } catch (e) {}
  return null;
}
let _saveT;
function debouncedSave() { clearTimeout(_saveT); _saveT = setTimeout(save, 250); }

/* ---------- scoring ---------- */
function computeScores() {
  const items = state.items;
  const N = items.length;
  let totalWeight = 0, netWeight = 0, gPass = 0, cPass = 0, gW = 0, cW = 0;
  let diff = 0, leaked = 0, bothFail = 0;
  for (const it of items) {
    // A negative-weight (penalty) rule is a positive rule of magnitude |w| whose
    // "pass" is AVOIDING the behavior (score 0); a positive rule passes on score 1.
    const aw = Math.abs(it.weight);
    const gPassed = it.weight >= 0 ? it.grok === 1 : it.grok === 0;
    const cPassed = it.weight >= 0 ? it.claude === 1 : it.claude === 0;
    totalWeight += aw;
    netWeight += it.weight;
    if (gPassed) { gPass++; gW += aw; }
    if (cPassed) { cPass++; cW += aw; }
    if (gPassed) leaked++;          // Grok passed → leaked (bad for a hard eval)
    else if (cPassed) diff++;       // only Claude passed → differentiating (ideal)
    else bothFail++;                // neither passed
  }
  const pct = (num, den) => (den > 0 ? (num / den) * 100 : 0);
  const grokRate = pct(gPass, N);
  const claudeRate = pct(cPass, N);
  const grokWeighted = pct(gW, totalWeight);
  const claudeWeighted = pct(cW, totalWeight);
  const gap = claudeRate - grokRate;
  const weightedGap = claudeWeighted - grokWeighted;
  const targetHit = grokWeighted < 50 && claudeWeighted < 60;
  const strongDiff = claudeWeighted < 80 && weightedGap >= 25;
  return {
    N, totalWeight, netWeight, gPass, cPass, gW, cW,
    diff, leaked, bothFail,
    grokRate, claudeRate,
    grokW: grokWeighted, claudeW: claudeWeighted,
    gap, weightedGap, targetHit, strongDiff, accepted: targetHit || strongDiff
  };
}

function formatDecimal(value, digits = 2) {
  if (!Number.isFinite(value)) return '0';
  const sign = value < 0 ? '-' : '';
  const factor = Math.pow(10, digits);
  const scaled = Math.trunc(Math.abs(value) * factor);
  const whole = Math.floor(scaled / factor);
  const frac = String(scaled % factor).padStart(digits, '0').replace(/0+$/, '');
  return sign + whole + (frac ? '.' + frac : '');
}

function formatPercent(value) {
  return formatDecimal(value, 2) + '%';
}

function formatPointGap(value) {
  return (value >= 0 ? '+' : '') + formatDecimal(value, 2) + ' pp';
}

/* ---------- DOM refs ---------- */
const $ = (id) => document.getElementById(id);
const body = $('rubricBody');

/* ---------- summary render ---------- */
function renderSummary() {
  const s = computeScores();
  $('sumRules').textContent = s.N;
  $('sumWeight').textContent = s.totalWeight;
  $('sumWeight').title = 'Total achievable weight (sum of |weights|): ' + s.totalWeight + ' wt. Net rubric weight: ' + s.netWeight + ' wt.';

  $('sumGrokU').textContent = formatPercent(s.grokRate);
  $('sumGrokUsub').textContent = s.gPass + '/' + s.N + ' passed';
  $('sumGrokW').textContent = formatPercent(s.grokW);
  $('sumGrokWsub').textContent = s.gW + '/' + s.totalWeight + ' wt';
  $('sumClaudeU').textContent = formatPercent(s.claudeRate);
  $('sumClaudeUsub').textContent = s.cPass + '/' + s.N + ' passed';
  $('sumClaudeW').textContent = formatPercent(s.claudeW);
  $('sumClaudeWsub').textContent = s.cW + '/' + s.totalWeight + ' wt';
  const setGap = (id, val) => {
    const el = $(id);
    el.textContent = formatPointGap(val);
    el.style.color = val >= 25 ? 'var(--green)' : val < 0 ? 'var(--red)' : 'var(--text)';
  };
  setGap('sumGapU', s.gap);
  setGap('sumGapW', s.claudeW - s.grokW);

  const badge = $('decisionBadge');
  badge.textContent = s.accepted ? 'ACCEPTED' : 'REJECTED';
  badge.className = 'card-value badge ' + (s.accepted ? 'accepted' : 'rejected');
  const dCard = $('decisionCard');
  dCard.classList.toggle('ok', s.accepted);
  dCard.classList.toggle('bad', !s.accepted);
  const t = $('critTarget'), d = $('critDiff');
  t.className = 'crit-dot ' + (s.targetHit ? 'ok' : 'no');
  d.className = 'crit-dot ' + (s.strongDiff ? 'ok' : 'no');
  t.title = 'Weighted Target Hit: Grok <50% and Claude <60% (Grok ' + formatPercent(s.grokW) + ', Claude ' + formatPercent(s.claudeW) + ')';
  d.title = 'Weighted Strong Differentiation: Claude <80% and (Claude - Grok) >=25pp (gap ' + formatPointGap(s.weightedGap) + ')';
}

/* ---------- filtering ---------- */
function passFilter(it) {
  const f = state.filters;
  if (f.category !== 'all' && it.category !== f.category) return false;
  if (f.grok === 'pass' && it.grok !== 1) return false;
  if (f.grok === 'fail' && it.grok !== 0) return false;
  if (f.claude === 'pass' && it.claude !== 1) return false;
  if (f.claude === 'fail' && it.claude !== 0) return false;
  if (f.search && !it.text.toLowerCase().includes(f.search)) return false;
  return true;
}

/* ---------- circle (binary red <-> green) ---------- */
function circleClass(v) { return 'circle ' + (v === 1 ? 'pass' : 'fail'); }
function cycle(v) { return v === 1 ? 0 : 1; }
function passTitle(model, v) {
  return model + ': ' + (v === 1 ? 'Pass (click → Fail)' : 'Fail (click → Pass)');
}

/* ---------- table render ---------- */
function renderTable() {
  body.innerHTML = '';
  const visible = state.items.filter(passFilter);
  $('emptyState').hidden = state.items.length !== 0;
  $('addRowTr').hidden = false;
  $('visibleCount').textContent =
    state.items.length === 0 ? '' : 'Showing ' + visible.length + ' of ' + state.items.length;

  let displayIdx = 0;
  for (const it of visible) {
    displayIdx++;
    const tr = document.createElement('tr');
    tr.dataset.id = it.id;
    tr.draggable = true;

    // grip
    const tdGrip = document.createElement('td');
    tdGrip.className = 'grip';
    tdGrip.textContent = '⠿';
    tdGrip.title = 'Drag to reorder';
    tr.appendChild(tdGrip);

    // index (C-prefixed pill)
    const tdIdx = document.createElement('td');
    tdIdx.className = 'idx';
    const pill = document.createElement('span');
    pill.className = 'id-pill';
    pill.textContent = 'C' + displayIdx;
    tdIdx.appendChild(pill);
    tr.appendChild(tdIdx);

    // model toggle circles, lettered (A = Grok, B = Claude)
    const tdModels = document.createElement('td');
    const modelsWrap = document.createElement('div');
    modelsWrap.className = 'models-cell';
    for (const [model, letter, label] of [['grok', 'A', 'Grok'], ['claude', 'B', 'Claude']]) {
      const c = document.createElement('button');
      c.type = 'button';
      c.className = circleClass(it[model]);
      c.textContent = letter;
      c.title = passTitle(label, it[model]);
      c.addEventListener('click', () => {
        it[model] = cycle(it[model]);
        c.className = circleClass(it[model]);
        c.title = passTitle(label, it[model]);
        renderSummary(); save();
      });
      modelsWrap.appendChild(c);
    }
    tdModels.appendChild(modelsWrap);
    tr.appendChild(tdModels);

    // text (contenteditable)
    const tdText = document.createElement('td');
    const txt = document.createElement('div');
    txt.className = 'cell-text';
    txt.contentEditable = 'true';
    txt.spellcheck = false;
    txt.textContent = it.text;
    txt.addEventListener('input', () => { it.text = txt.textContent; debouncedSave(); });
    txt.addEventListener('blur', () => {
      it.text = txt.textContent;
      if (hasCorpus()) { analyzeItem(it); save(); renderTable(); }   // refresh matched-file hints only
      else save();
    });
    tdText.appendChild(txt);
    const mf = buildMatchEl(it);
    if (mf) tdText.appendChild(mf);
    tr.appendChild(tdText);

    // category (segmented toggle: Content / Execution / Style)
    const tdCat = document.createElement('td');
    const toggle = document.createElement('div');
    toggle.className = 'cat-toggle';
    for (const [value, label] of [['Content', 'Content'], ['Execution', 'Execution'], ['Presentation & Formatting', 'Style']]) {
      const b = document.createElement('button');
      b.type = 'button';
      b.className = 'seg' + (it.category === value ? ' active' : '');
      b.dataset.cat = value;
      b.textContent = label;
      b.addEventListener('click', () => {
        it.category = value;
        toggle.querySelectorAll('.seg').forEach((s) => s.classList.toggle('active', s.dataset.cat === value));
        save();
      });
      toggle.appendChild(b);
    }
    tdCat.appendChild(toggle);
    tr.appendChild(tdCat);

    // weight
    const tdWt = document.createElement('td');
    const wt = document.createElement('input');
    wt.className = 'wt-input';
    wt.type = 'number'; wt.min = String(MIN_WEIGHT); wt.max = String(MAX_WEIGHT); wt.step = '1';
    wt.title = 'Weight from -10 to 10';
    wt.value = it.weight;
    const commitWt = () => {
      const n = normWeight(wt.value);
      it.weight = n; wt.value = n; wt.classList.remove('invalid');
      renderSummary(); save();
    };
    wt.addEventListener('input', () => {
      const n = parseInt(wt.value, 10);
      wt.classList.toggle('invalid', isNaN(n) || n < MIN_WEIGHT || n > MAX_WEIGHT);
    });
    wt.addEventListener('change', commitWt);
    wt.addEventListener('blur', commitWt);
    tdWt.appendChild(wt);
    tr.appendChild(tdWt);

    // actions
    const tdAct = document.createElement('td');
    const wrap = document.createElement('div');
    wrap.className = 'row-actions';
    const dup = document.createElement('button');
    dup.className = 'icon-btn';
    dup.title = 'Duplicate rule';
    dup.setAttribute('aria-label', 'Duplicate rule');
    dup.innerHTML = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="9" y="9" width="11" height="11" rx="2"/><path d="M5 15V7a2 2 0 0 1 2-2h8"/></svg>';
    dup.addEventListener('click', () => duplicateItem(it.id));
    const del = document.createElement('button');
    del.className = 'icon-btn del';
    del.title = 'Delete rule';
    del.setAttribute('aria-label', 'Delete rule');
    del.innerHTML = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 6h18"/><path d="M8 6V4a1 1 0 0 1 1-1h6a1 1 0 0 1 1 1v2"/><path d="M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"/><path d="M10 11v6M14 11v6"/></svg>';
    del.addEventListener('click', () => deleteItem(it.id));
    wrap.appendChild(dup); wrap.appendChild(del);
    tdAct.appendChild(wrap);
    tr.appendChild(tdAct);

    attachDrag(tr);
    body.appendChild(tr);
  }
  renderSummary();
}

/* ---------- drag & drop reorder ---------- */
let dragId = null;
function attachDrag(tr) {
  tr.addEventListener('dragstart', (e) => {
    dragId = tr.dataset.id;
    tr.classList.add('dragging');
    e.dataTransfer.effectAllowed = 'move';
    try { e.dataTransfer.setData('text/plain', dragId); } catch (_) {}
  });
  tr.addEventListener('dragend', () => {
    tr.classList.remove('dragging');
    document.querySelectorAll('.drop-target').forEach((r) => r.classList.remove('drop-target'));
    dragId = null;
  });
  tr.addEventListener('dragover', (e) => {
    e.preventDefault();
    e.dataTransfer.dropEffect = 'move';
    if (tr.dataset.id !== dragId) tr.classList.add('drop-target');
  });
  tr.addEventListener('dragleave', () => tr.classList.remove('drop-target'));
  tr.addEventListener('drop', (e) => {
    e.preventDefault();
    tr.classList.remove('drop-target');
    const targetId = tr.dataset.id;
    if (!dragId || dragId === targetId) return;
    const from = state.items.findIndex((x) => x.id === dragId);
    const to = state.items.findIndex((x) => x.id === targetId);
    if (from < 0 || to < 0) return;
    const [moved] = state.items.splice(from, 1);
    state.items.splice(to, 0, moved);
    save(); renderTable();
  });
}

/* ---------- mutations ---------- */
function addItem(atEnd) {
  const it = makeItem({ text: '', category: 'Content', weight: 3 });
  if (atEnd) state.items.push(it); else state.items.unshift(it);
  if (hasCorpus()) analyzeItem(it);
  save(); renderTable();
  const row = body.querySelector('tr[data-id="' + it.id + '"]');
  if (row) {
    row.scrollIntoView({ block: 'nearest' });
    const cell = row.querySelector('.cell-text');
    if (cell) cell.focus();
  }
}
function duplicateItem(id) {
  const i = state.items.findIndex((x) => x.id === id);
  if (i < 0) return;
  const src = state.items[i];
  const copy = makeItem({ text: src.text, category: src.category, weight: src.weight, grok: src.grok, claude: src.claude });
  state.items.splice(i + 1, 0, copy);
  if (hasCorpus()) analyzeItem(copy);
  save(); renderTable();
  toast('Rule duplicated', 'ok');
}
function deleteItem(id) {
  const i = state.items.findIndex((x) => x.id === id);
  if (i < 0) return;
  const removed = state.items[i];
  state.items.splice(i, 1);
  matchIndex.delete(id);
  save(); renderTable();
  toastUndo('Rule deleted', () => {
    state.items.splice(Math.min(i, state.items.length), 0, removed);
    if (hasCorpus()) analyzeItem(removed);
    save(); renderTable();
    toast('Rule restored', 'ok');
  });
}

/* ---------- toast ---------- */
let _toastT;
function toast(msg, kind) {
  const el = $('toast');
  el.textContent = msg;
  el.className = 'toast' + (kind ? ' ' + kind : '');
  el.hidden = false;
  clearTimeout(_toastT);
  _toastT = setTimeout(() => { el.hidden = true; }, 2600);
}
// toast with an inline action (used for undoing a delete)
function toastUndo(msg, onAction, actionLabel) {
  const el = $('toast');
  el.className = 'toast';
  el.textContent = msg + ' ';
  const btn = document.createElement('button');
  btn.type = 'button';
  btn.className = 'toast-action';
  btn.textContent = actionLabel || 'Undo';
  btn.addEventListener('click', () => {
    el.hidden = true;
    clearTimeout(_toastT);
    onAction();
  });
  el.appendChild(btn);
  el.hidden = false;
  clearTimeout(_toastT);
  _toastT = setTimeout(() => { el.hidden = true; }, 6000);
}

/* ---------- file download ---------- */
function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

/* ---------- JSON import/export ---------- */
function exportJSON() {
  const blob = new Blob([JSON.stringify(itemsToJSON(), null, 2)], { type: 'application/json' });
  downloadBlob(blob, 'rubric.json');
  toast('Exported rubric.json', 'ok');
}
async function copyJSON() {
  const text = JSON.stringify(itemsToJSON(), null, 2);
  try {
    if (navigator.clipboard && navigator.clipboard.writeText) {
      await navigator.clipboard.writeText(text);
    } else {
      const ta = document.createElement('textarea');
      ta.value = text;
      ta.setAttribute('readonly', '');
      ta.style.position = 'fixed';
      ta.style.top = '-9999px';
      document.body.appendChild(ta);
      ta.select();
      const ok = document.execCommand('copy');
      ta.remove();
      if (!ok) throw new Error('copy command failed');
    }
    toast('Copied rubric JSON', 'ok');
  } catch (e) {
    toast('Copy failed: ' + e.message, 'err');
  }
}
function importJSONText(text) {
  let obj;
  try { obj = JSON.parse(text); } catch (e) { toast('Invalid JSON file', 'err'); return; }
  const items = itemsFromJSON(obj);
  if (!items.length) { toast('No rubric items found in JSON', 'err'); return; }
  state.items = items;
  if (hasCorpus()) analyzeAll(); else { save(); renderTable(); }
  toast('Imported ' + items.length + ' rules from JSON', 'ok');
}

/* ---------- Paste (JSON / Excel-TSV / plain text) ---------- */
function itemsFromRows(rows) {
  rows = rows.filter((r) => r.some((c) => String(c).trim() !== ''));
  if (!rows.length) return [];
  const first = rows[0].map((c) => String(c).trim().toLowerCase());
  const hints = ['rule', 'text', 'description', 'category', 'weight', 'wt', 'grok', 'claude', 'opus'];
  // Treat row 1 as a header only if it really looks like one (short, ≥2 header-ish cells),
  // so a prose first row is never mistaken for headers and dropped.
  const looksHeader = first.length >= 2 &&
    first.every((c) => c.length <= 24) &&
    first.filter((c) => hints.some((h) => c === h || c.includes(h))).length >= 2;
  let header = null, body = rows;
  if (looksHeader) { header = first; body = rows.slice(1); }

  let iText = 0, iCat = -1, iWt = -1, iGrok = -1, iClaude = -1;
  if (header) {
    const find = (cands) => header.findIndex((h) => cands.some((c) => h === c || h.includes(c)));
    const t = find(['rule', 'text', 'description']); iText = t >= 0 ? t : 0;
    iCat = find(['category', 'rule cat']);
    iWt = find(['weight', 'wt']);
    iGrok = find(['grok']);
    iClaude = find(['claude', 'opus']);
  } else {
    // positional fallback matches the Export order: text, category, weight, grok, claude
    const n = rows[0].length;
    if (n >= 2) iCat = 1;
    if (n >= 3) iWt = 2;
    if (n >= 4) iGrok = 3;
    if (n >= 5) iClaude = 4;
  }
  return body
    .map((r) => makeItem({
      text: r[iText],
      category: iCat >= 0 ? r[iCat] : 'Content',
      weight: iWt >= 0 ? r[iWt] : 1,
      grok: iGrok >= 0 ? r[iGrok] : null,
      claude: iClaude >= 0 ? r[iClaude] : null
    }))
    .filter((it) => it.text.trim() !== '');
}

function parsePastedText(raw) {
  const text = String(raw == null ? '' : raw).replace(/\r\n?/g, '\n').trim();
  if (!text) return { items: [], fmt: null };

  // 1) JSON — export format { items: [...] } or a bare array
  if (text[0] === '{' || text[0] === '[') {
    try {
      const items = itemsFromJSON(JSON.parse(text));
      if (items.length) return { items, fmt: 'JSON' };
    } catch (e) { /* not valid JSON — fall through */ }
  }

  const lines = text.split('\n').filter((l) => l.trim() !== '');

  // 2) Tab-delimited table — what Excel/Sheets put on the clipboard when you copy cells
  if (text.includes('\t')) {
    const items = itemsFromRows(lines.map((l) => l.split('\t')));
    if (items.length) return { items, fmt: 'Excel' };
  }

  // 3) Plain text — one rule per non-empty line
  const items = lines.map((l) => makeItem({ text: l.trim() })).filter((it) => it.text !== '');
  return { items, fmt: 'text' };
}

function applyPaste(text) {
  const { items, fmt } = parsePastedText(text);
  if (!items.length) { toast('No rubric rows found in the pasted text', 'err'); return false; }
  state.items = items;
  if (hasCorpus()) analyzeAll(); else { save(); renderTable(); }
  toast('Pasted ' + items.length + ' rule' + (items.length === 1 ? '' : 's') + ' from ' + fmt + ' — replaced existing', 'ok');
  return true;
}

let _pasteArmed = false;
function armManualPaste() {
  _pasteArmed = true;
  toast('Press Ctrl+V (⌘V) to paste your rubric', 'ok');
  setTimeout(() => { _pasteArmed = false; }, 20000);
}
async function pasteRubric() {
  // Preferred path: async Clipboard API (Chromium/Safari over https or localhost, with this click as the gesture)
  if (navigator.clipboard && navigator.clipboard.readText) {
    try {
      const text = await navigator.clipboard.readText();
      if (text && text.trim()) { applyPaste(text); return; }
    } catch (e) { /* permission denied / unsupported (e.g. Firefox) — fall back to manual Ctrl+V */ }
  }
  armManualPaste();
}

/* ---------- Excel (SheetJS, lazy) ---------- */
function loadScript(src) {
  return new Promise((resolve, reject) => {
    const s = document.createElement('script');
    s.src = src; s.onload = resolve; s.onerror = () => reject(new Error('load failed: ' + src));
    document.head.appendChild(s);
  });
}
async function ensureXLSX() {
  if (window.XLSX) return;
  try { await loadScript(XLSX_LOCAL); } catch (e) { /* fall through to CDN */ }
  if (!window.XLSX) await loadScript(XLSX_CDN);
  if (!window.XLSX) throw new Error('SheetJS unavailable');
}

const HDR = { rule: 'Rule', cat: 'Rule Category', wt: 'Weight', grok: 'Grok', claude: 'Claude' };

function rowsForExcel() {
  const pf = (v) => (v === 1 ? 'Pass' : v === 0 ? 'Fail' : '');
  return state.items.map((it) => ({
    [HDR.rule]: it.text,
    [HDR.cat]: it.category,
    [HDR.wt]: it.weight,
    [HDR.grok]: pf(it.grok),
    [HDR.claude]: pf(it.claude)
  }));
}
async function exportExcel() {
  try {
    await ensureXLSX();
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(rowsForExcel(), {
      header: [HDR.rule, HDR.cat, HDR.wt, HDR.grok, HDR.claude]
    });
    ws['!cols'] = [{ wch: 80 }, { wch: 26 }, { wch: 8 }, { wch: 10 }, { wch: 10 }];
    XLSX.utils.book_append_sheet(wb, ws, 'Rubric');
    XLSX.writeFile(wb, 'rubric.xlsx');
    toast('Exported rubric.xlsx', 'ok');
  } catch (e) { toast('Excel export failed: ' + e.message, 'err'); }
}
function pickHeader(keys, candidates) {
  // returns the actual key in the row matching any candidate (case-insensitive, trimmed)
  const map = {};
  keys.forEach((k) => { map[String(k).trim().toLowerCase()] = k; });
  for (const c of candidates) { if (map[c]) return map[c]; }
  // partial contains match
  for (const lk in map) {
    if (candidates.some((c) => lk.includes(c))) return map[lk];
  }
  return null;
}
async function importExcelFile(file) {
  try {
    await ensureXLSX();
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: 'array' });
    // choose first sheet whose header row has a text-ish col and a weight col
    let chosen = null, rows = null;
    for (const name of wb.SheetNames) {
      const r = XLSX.utils.sheet_to_json(wb.Sheets[name], { defval: '' });
      if (!r.length) continue;
      const keys = Object.keys(r[0]);
      const hasText = pickHeader(keys, ['rule', 'text', 'description']) &&
        !(keys.length === 1);
      const hasWt = pickHeader(keys, ['weight']);
      if (hasText && hasWt) { chosen = name; rows = r; break; }
      if (!chosen) { chosen = name; rows = r; } // fallback to first non-empty
    }
    if (!rows || !rows.length) { toast('No rows found in Excel', 'err'); return; }
    const keys = Object.keys(rows[0]);
    const kText = pickHeader(keys, ['rule', 'text', 'description']);
    const kCat = pickHeader(keys, ['rule category', 'rule_category', 'category', 'rule cat']);
    const kWt = pickHeader(keys, ['weight', 'wt']);
    const kGrok = pickHeader(keys, ['grok']);
    const kClaude = pickHeader(keys, ['claude', 'opus']);
    const items = rows
      .map((r) => makeItem({
        text: kText ? r[kText] : '',
        category: kCat ? r[kCat] : 'Content',
        weight: kWt ? r[kWt] : 1,
        grok: kGrok ? r[kGrok] : null,
        claude: kClaude ? r[kClaude] : null
      }))
      .filter((it) => it.text.trim() !== '');
    if (!items.length) { toast('No rubric rows with text found', 'err'); return; }
    state.items = items;
    if (hasCorpus()) analyzeAll(); else { save(); renderTable(); }
    toast('Imported ' + items.length + ' rules from Excel', 'ok');
  } catch (e) { toast('Excel import failed: ' + e.message, 'err'); }
}

/* ============================================================
   Response review — find likely files for each rubric point in the
   uploaded Grok (A) / Claude (B) response archives (.zip).

   Pipeline:  rubric text → keywords (stop-words & boilerplate
   verbs removed) → matched against every text file in the zip
   → TF-IDF-weighted coverage. Scores are never changed here;
   the user checks the files and scores manually.
   ============================================================ */

// In-memory only (zips can be large — not persisted to localStorage).
const corpora = { grok: null, claude: null };       // model → { name, files:[{path,text}], reCache }
const matchIndex = new Map();                        // item.id → { grok:{...}, claude:{...} }

// Words carrying no discriminating signal — dropped before matching.
const STOPWORDS = new Set(('a an the and or of to in on for with that this is are be been being was were ' +
  'by as at it its from into their they them these those if then else than so such via per each any all ' +
  'both which who whom whose what when where why how not no nor only also very more most less least over ' +
  'under about above below between among within across against during after before while until once here ' +
  'there out up down off but one two three exact given whether based using used use other same each every ' +
  'data set sets number numbers value values').split(/\s+/));

// Rubric boilerplate — verbs/nouns that describe *the act of answering*, not the subject matter.
const BOILERPLATE = new Set(('response responses answer answers reply replies submission output outputs ' +
  'result results include includes included including provide provides provided contain contains containing ' +
  'state states stated says report reports reported show shows shown present presents presented describe ' +
  'describes described explain explains explained mention mentions mentioned ensure ensures ensured make ' +
  'makes made produce produces produced return returns returned generate generates generated create creates ' +
  'created deliver delivers delivered should must shall will would could can may might correctly explicitly').split(/\s+/));

// rubric text → ranked, de-duplicated keyword list
function tokenizeRubric(text) {
  const out = new Set();
  String(text || '').toLowerCase().split(/[^a-z0-9]+/).forEach((raw) => {
    if (!raw || raw.length < 3) return;
    if (/^\d+$/.test(raw)) return;                       // pure numbers
    if (STOPWORDS.has(raw) || BOILERPLATE.has(raw)) return;
    let w = raw;
    // light singularisation so "clients"→"client", "models"→"model", "computes"→"compute"
    if (w.length > 4 && w.endsWith('s') && !w.endsWith('ss')) w = w.slice(0, -1);
    if (w.length >= 3 && !STOPWORDS.has(w) && !BOILERPLATE.has(w)) out.add(w);
  });
  return [...out];
}

function baseName(path) {
  const p = String(path).replace(/\\/g, '/').replace(/\/+$/, '');
  const i = p.lastIndexOf('/');
  return i >= 0 ? p.slice(i + 1) : p;
}

function escapeRe(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }

// Match one keyword set against one corpus.
// coverage = (distinct keywords found in ≥1 file) / (keywords);  files ranked by Σ idf.
function matchAgainstCorpus(keywords, corpus) {
  if (!corpus || !corpus.files.length || !keywords.length) {
    return { score: 0, matched: 0, total: keywords.length, files: [] };
  }
  const df = {}; keywords.forEach((k) => { df[k] = 0; });
  const perFile = [];
  for (const f of corpus.files) {
    const kws = [];
    for (const kw of keywords) {
      let re = corpus.reCache[kw];
      if (!re) { re = corpus.reCache[kw] = new RegExp('\\b' + escapeRe(kw)); }
      if (re.test(f.text)) { kws.push(kw); df[kw]++; }
    }
    if (kws.length) perFile.push({ path: f.path, kws });
  }
  const N = corpus.files.length;
  const idf = {}; keywords.forEach((k) => { idf[k] = Math.log(1 + N / (1 + df[k])); });
  const matched = new Set();
  perFile.forEach((pf) => pf.kws.forEach((k) => matched.add(k)));
  perFile.forEach((pf) => { pf.score = pf.kws.reduce((s, k) => s + (idf[k] || 0.1), 0); });
  perFile.sort((a, b) => b.score - a.score || b.kws.length - a.kws.length);
  const coverage = matched.size / keywords.length;
  return {
    score: coverage,
    matched: matched.size,
    total: keywords.length,
    files: perFile.slice(0, MAX_MATCH_FILES).map((p) => ({ name: baseName(p.path), path: p.path, kws: p.kws }))
  };
}

// Find likely files for one rule against whatever corpora are loaded.
// This intentionally does not update A/B pass-fail scores.
function analyzeItem(it) {
  const kws = tokenizeRubric(it.text);
  const res = {};
  for (const model of ['grok', 'claude']) {
    const c = corpora[model];
    if (!c) continue;
    const r = matchAgainstCorpus(kws, c);
    res[model] = r;
  }
  if (Object.keys(res).length) matchIndex.set(it.id, res);
  else matchIndex.delete(it.id);
}

function hasCorpus() { return !!(corpora.grok || corpora.claude); }

function analyzeAll() {
  if (!hasCorpus()) return;
  matchIndex.clear();
  state.items.forEach(analyzeItem);
  save();
  renderTable();
}

// Build a small "matched files" block to hang under a rubric point.
function buildMatchEl(it) {
  const res = matchIndex.get(it.id);
  if (!res) return null;
  const wrap = document.createElement('div');
  wrap.className = 'match-files';
  let any = false;
  for (const [model, letter] of [['grok', 'A'], ['claude', 'B']]) {
    const r = res[model];
    if (!r) continue;
    any = true;
    const line = document.createElement('div');
    line.className = 'mf-line ' + model;

    const tag = document.createElement('span');
    tag.className = 'mf-tag';
    tag.textContent = letter;
    line.appendChild(tag);

    if (r.files.length) {
      r.files.forEach((f) => {
        const chip = document.createElement('button');
        chip.type = 'button';
        chip.className = 'mf-file';
        chip.textContent = f.name;
        chip.title = 'View ' + f.path + (f.kws && f.kws.length ? ' — matched: ' + f.kws.join(', ') : '');
        chip.addEventListener('click', () => openFileModal(model, f.path, f.kws));
        line.appendChild(chip);
      });
    } else {
      const none = document.createElement('span');
      none.className = 'mf-none';
      none.textContent = 'no matching file';
      line.appendChild(none);
    }

    const sc = document.createElement('span');
    sc.className = 'mf-score';
    sc.textContent = Math.round(r.score * 100) + '% match';
    sc.title = r.matched + '/' + r.total + ' keyword' + (r.total === 1 ? '' : 's') +
      ' found. Review the files and score this rule manually.';
    line.appendChild(sc);

    wrap.appendChild(line);
  }
  return any ? wrap : null;
}

/* ---------- file viewer (click a matched file → see its content) ---------- */
function escapeHtml(s) {
  return String(s).replace(/[&<>]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]));
}
// escape the file text, then wrap each matched keyword (and its inflection) in <mark>
function highlightContent(raw, kws) {
  let html = escapeHtml(raw);
  if (kws && kws.length) {
    const alt = kws.slice().sort((a, b) => b.length - a.length).map(escapeRe).join('|');
    try {
      const re = new RegExp('\\b(' + alt + ')[a-z0-9]*', 'gi');
      html = html.replace(re, '<mark>$&</mark>');
    } catch (e) { /* keep unhighlighted on bad regex */ }
  }
  return html;
}
let _modalLastFocus = null;
function openFileModal(model, path, kws) {
  const c = corpora[model];
  if (!c) return;
  const file = c.files.find((f) => f.path === path);
  if (!file) return;
  _modalLastFocus = document.activeElement;
  const tag = $('fileModalTag');
  tag.textContent = model === 'grok' ? 'A' : 'B';
  tag.className = 'modal-tag ' + model;
  $('fileModalTitle').textContent = baseName(path);
  const who = model === 'grok' ? 'Grok (A)' : 'Claude (B)';
  const kwNote = kws && kws.length ? ' · ' + kws.length + ' keyword' + (kws.length === 1 ? '' : 's') + ' matched' : '';
  $('fileModalMeta').textContent = who + ' · ' + path + kwNote;
  $('fileModalContent').innerHTML = highlightContent(file.raw != null ? file.raw : '', kws);
  $('fileModalContent').scrollTop = 0;
  $('fileModal').hidden = false;
  document.body.classList.add('modal-open');
  $('fileModalClose').focus();
}
function closeFileModal() {
  const m = $('fileModal');
  if (m.hidden) return;
  m.hidden = true;
  document.body.classList.remove('modal-open');
  if (_modalLastFocus && _modalLastFocus.focus) _modalLastFocus.focus();
}

/* ---------- zip ingestion (JSZip, lazy) ---------- */
async function ensureJSZip() {
  if (window.JSZip) return;
  try { await loadScript(JSZIP_LOCAL); } catch (e) { /* fall through to CDN */ }
  if (!window.JSZip) await loadScript(JSZIP_CDN);
  if (!window.JSZip) throw new Error('JSZip unavailable');
}

// readable text formats we care to index (code, notebooks, docs, data)
const TEXT_EXT = /\.(txt|md|markdown|rst|tex|py|ipynb|js|mjs|cjs|ts|jsx|tsx|json|jsonl|csv|tsv|html?|css|scss|java|c|cc|cpp|cxx|h|hpp|cs|go|rs|rb|php|sh|bash|zsh|bat|ps1|sql|ya?ml|toml|ini|cfg|conf|properties|r|rmd|m|jl|scala|kt|kts|swift|pl|lua|xml|svg|env|gradle|dockerfile|makefile|log)$/i;
const SKIP_PATH = /(^|\/)(__MACOSX|\.git|\.svn|node_modules|\.venv|venv|env|\.idea|\.vscode|\.ipynb_checkpoints|dist|build|__pycache__)\//i;
const MAX_FILE_BYTES = 2000000;   // cap any single file's contribution

async function buildCorpusFromZip(file) {
  await ensureJSZip();
  const zip = await JSZip.loadAsync(await file.arrayBuffer());
  const entries = Object.keys(zip.files).map((k) => zip.files[k])
    .filter((e) => !e.dir && !SKIP_PATH.test('/' + e.name) && TEXT_EXT.test(e.name));
  const files = [];
  for (const e of entries) {
    let raw;
    try { raw = await e.async('string'); } catch (_) { continue; }
    if (raw.length > MAX_FILE_BYTES) raw = raw.slice(0, MAX_FILE_BYTES);
    // `text` (lowercased, filename-prefixed) is for matching; `raw` (original) is for the viewer
    files.push({ path: e.name, raw, text: (e.name + '\n' + raw).toLowerCase() });
  }
  return { name: file.name, files, reCache: {} };
}

function updateUploadStatus(model) {
  const btn = $(model === 'grok' ? 'btnUploadGrok' : 'btnUploadClaude');
  if (!btn) return;
  const c = corpora[model];
  const name = model === 'grok' ? 'Grok' : 'Claude';
  if (c) {
    btn.textContent = '✓ ' + name + ' · ' + c.files.length;
    btn.classList.add('loaded');
    btn.title = c.name + ' — ' + c.files.length + ' text file' + (c.files.length === 1 ? '' : 's') +
      ' indexed for manual review. Scores are not changed. Click to replace.';
  } else {
    btn.textContent = '↑ ' + name + ' .zip';
    btn.classList.remove('loaded');
  }
}

async function handleZipUpload(file, model) {
  if (!file) return;
  const label = model === 'grok' ? 'Grok (A)' : 'Claude (B)';
  try {
    toast('Reading ' + file.name + ' …', 'ok');
    const corpus = await buildCorpusFromZip(file);
    if (!corpus.files.length) { toast('No readable text files found in ' + file.name, 'err'); return; }
    corpora[model] = corpus;
    updateUploadStatus(model);
    analyzeAll();
    toast(label + ': indexed ' + corpus.files.length +
      ' file' + (corpus.files.length === 1 ? '' : 's') + ' for manual review', 'ok');
  } catch (e) {
    toast(label + ' upload failed: ' + e.message, 'err');
  }
}

/* ---------- wiring ---------- */
function wire() {
  $('btnAddBottom').addEventListener('click', () => addItem(true));
  $('btnCopy').addEventListener('click', copyJSON);
  $('btnExportJSON').addEventListener('click', exportJSON);
  $('btnExportExcel').addEventListener('click', exportExcel);

  $('btnImportJSON').addEventListener('click', () => $('fileJSON').click());
  $('btnImportExcel').addEventListener('click', () => $('fileExcel').click());
  $('fileJSON').addEventListener('change', (e) => {
    const f = e.target.files[0]; if (!f) return;
    const reader = new FileReader();
    reader.onload = () => importJSONText(reader.result);
    reader.readAsText(f);
    e.target.value = '';
  });
  $('fileExcel').addEventListener('change', (e) => {
    const f = e.target.files[0]; if (!f) return;
    importExcelFile(f);
    e.target.value = '';
  });

  $('btnUploadGrok').addEventListener('click', () => $('fileGrokZip').click());
  $('btnUploadClaude').addEventListener('click', () => $('fileClaudeZip').click());
  $('fileGrokZip').addEventListener('change', (e) => {
    const f = e.target.files[0]; handleZipUpload(f, 'grok'); e.target.value = '';
  });
  $('fileClaudeZip').addEventListener('change', (e) => {
    const f = e.target.files[0]; handleZipUpload(f, 'claude'); e.target.value = '';
  });

  $('btnPaste').addEventListener('click', pasteRubric);
  document.addEventListener('paste', (e) => {
    if (!_pasteArmed) return;
    const t = e.target;
    if (t && (t.isContentEditable || t.tagName === 'INPUT' || t.tagName === 'TEXTAREA')) return;
    _pasteArmed = false;
    const text = (e.clipboardData || window.clipboardData).getData('text');
    if (text && text.trim()) { e.preventDefault(); applyPaste(text); }
    else toast('Clipboard had no text', 'err');
  });

  $('search').addEventListener('input', (e) => {
    state.filters.search = e.target.value.trim().toLowerCase();
    renderTable();
  });
  document.querySelectorAll('.pill-group').forEach((group) => {
    const key = group.dataset.filter; // 'category' | 'grok' | 'claude'
    group.addEventListener('click', (e) => {
      const btn = e.target.closest('.pill');
      if (!btn) return;
      state.filters[key] = btn.dataset.val;
      group.querySelectorAll('.pill').forEach((p) => p.classList.toggle('active', p === btn));
      renderTable();
    });
  });
  $('btnClear').addEventListener('click', () => {
    state.filters = { search: '', category: 'all', grok: 'all', claude: 'all' };
    $('search').value = '';
    document.querySelectorAll('.pill-group .pill').forEach((p) => {
      p.classList.toggle('active', p.dataset.val === 'all');
    });
    renderTable();
  });

  $('fileModalClose').addEventListener('click', closeFileModal);
  $('fileModal').addEventListener('click', (e) => {
    if (e.target && e.target.hasAttribute('data-close')) closeFileModal();
  });

  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') { closeFileModal(); return; }
    if ((e.ctrlKey || e.metaKey) && e.shiftKey && (e.key === 'k' || e.key === 'K')) {
      e.preventDefault();
      addItem();
    }
  });
}

/* ---------- seed ---------- */
function seed() {
  return itemsFromJSON({
    items: [
      { text: 'The response trains a centralized baseline model on the pooled data from all clients.', criteriaAnnotations: { Rule_Category: 'Execution', Weight: '9' }, grok: 0, claude: 1 },
      { text: 'The response computes the difference in the headline metric between the centralized and decentralized models.', criteriaAnnotations: { Rule_Category: 'Execution', Weight: '9' }, grok: 0, claude: 1 },
      { text: 'The response states explicitly whether the hypothesis is supported or rejected based on the measured difference.', criteriaAnnotations: { Rule_Category: 'Content', Weight: '10' }, grok: 0, claude: 0 },
      { text: 'The response partitions the training data across the exact number of clients stated in the prompt.', criteriaAnnotations: { Rule_Category: 'Content', Weight: '8' }, grok: 0, claude: 0 },
      { text: 'The response delivers a report.tex file.', criteriaAnnotations: { Rule_Category: 'Presentation & Formatting', Weight: '4' }, grok: 0, claude: 0 }
    ]
  });
}

/* ---------- init ---------- */
function init() {
  wire();
  const saved = loadSaved();
  state.items = saved && saved.length ? saved : seed();
  renderTable();
}
document.addEventListener('DOMContentLoaded', init);
