// app.js — main UI logic

const state = {
  currentProjectId: null,
  currentTab: 'basic', // 'basic' | 'air' | 'water' | 'geo' | 'noise' | 'eco'
  editingProjectId: null, // set when project modal is in "edit" mode
  importCatKey: null,
  importParsed: null, // { headers, rows, ... } (generic import)
  importMode: null, // 'generic' | 'smart'
  smartResult: null, // { rows, matchedSheets, skippedSheets, sites } from SmartParse.parseWorkbook
  itemSelection: null, // Set of currently-checked item values (both import modes)
  batchQueue: null, // [{ catKey, result }] pending confirmation, when batch-importing
  batchQueueTotal: 0,
  periodFilter: {}, // { [catKey]: '115年第1季' | '' (全部) } — which period is being viewed/edited
  importPeriod: '', // the period label chosen for the batch currently being imported
  columnFilters: {}, // { [catKey]: { [fieldKey]: Set of allowed display values } } — Excel-style AutoFilter
  columnSort: {}, // { [catKey]: { fieldKey, direction: 'asc'|'desc' } } — Excel-style column sort
  undoStack: {}, // { [`${projectId}::${catKey}`]: [{ description, rows, batches }] } — in-memory only, cleared on page reload (matches typical app undo behavior; deliberately not persisted to localStorage)
  redoStack: {}, // same shape — holds states undone away from, so they can be reapplied; cleared for a key whenever a fresh destructive action is pushed onto its undo stack (standard undo/redo semantics: a new change invalidates the old "future")
};

// ---------- reporting period (年/季) ----------
// The official filing is submitted one quarter at a time (the template filenames
// even say "115Q1" etc.), but nothing in the row data itself says which quarter it
// belongs to. This tags each imported batch with a period label — guessed from the
// sampling dates but always shown to the person to confirm/adjust — so data from
// different quarters doesn't get silently merged together at export time.
function guessPeriodFromRows(rows) {
  const dates = rows.map(r => r['日期(起)']).filter(Boolean).sort();
  if (dates.length === 0) return '';
  const mid = dates[Math.floor(dates.length / 2)]; // representative date, robust to a few outliers
  const m = mid.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!m) return '';
  const rocYear = parseInt(m[1], 10) - 1911;
  const quarter = Math.ceil(parseInt(m[2], 10) / 3);
  return `${rocYear}年第${quarter}季`;
}
/** Parses the canonical "115年第1季" string into { rocYear, quarter }, or null. */
function parsePeriodString(str) {
  const m = String(str || '').match(/^(\d+)年第(\d)季$/);
  return m ? { rocYear: m[1], quarter: m[2] } : null;
}
/** Accepts shorthand like "115Q1" / "115-Q1" / "115 Q1" and normalizes it to the
 *  canonical "115年第1季" form; anything else (including already-canonical strings
 *  or free text) passes through unchanged. */
function normalizePeriodShorthand(str) {
  const s = String(str || '').trim();
  const m = s.match(/^(\d+)\s*-?\s*[Qq](\d)$/);
  return m ? `${m[1]}年第${m[2]}季` : s;
}
function getKnownPeriods(project, catKey) {
  const rows = DataStore.getData(project.id, catKey);
  const periods = [...new Set(rows.map(r => r._period).filter(Boolean))];
  // sort roughly chronologically by the leading ROC year + quarter number embedded in the label
  return periods.sort();
}

// ---------- undo/redo (in-memory, per project+category, not persisted to localStorage) ----------
const UNDO_STACK_LIMIT = 10;
function undoKey(projectId, catKey) { return `${projectId}::${catKey}`; }
function deepCopy(v) { return JSON.parse(JSON.stringify(v)); }
/** Snapshots the category's CURRENT full row list — and its import-batch history
 *  list, since deleting a batch also removes its entry there — before a destructive
 *  change, so both can be restored together and stay consistent. Call this right
 *  before the change, not after. A fresh destructive action also clears any pending
 *  redo history for this key — standard undo/redo semantics: once you make a new
 *  change, the "future" that redo pointed to no longer makes sense to reapply. */
function pushUndoSnapshot(projectId, catKey, description) {
  const key = undoKey(projectId, catKey);
  if (!state.undoStack[key]) state.undoStack[key] = [];
  const rows = DataStore.getData(projectId, catKey);
  const batches = DataStore.getImportBatches(projectId, catKey);
  const presets = getItemPresets(catKey);
  state.undoStack[key].push({ description, rows: deepCopy(rows), batches: deepCopy(batches), presets: deepCopy(presets) });
  if (state.undoStack[key].length > UNDO_STACK_LIMIT) state.undoStack[key].shift();
  state.redoStack[key] = [];
}
function peekUndo(projectId, catKey) {
  const stack = state.undoStack[undoKey(projectId, catKey)];
  return stack && stack.length > 0 ? stack[stack.length - 1] : null;
}
function peekRedo(projectId, catKey) {
  const stack = state.redoStack[undoKey(projectId, catKey)];
  return stack && stack.length > 0 ? stack[stack.length - 1] : null;
}
/** Undo: restores the state from BEFORE the most recent tracked action, and pushes
 *  what was live just now onto the redo stack (tagged with the same description,
 *  since redoing this entry means "reapply that same action") so it can be brought
 *  back with redo. Also restores the "常用測項新增" preset list (getItemPresets) to
 *  whatever it was before — that action auto-saves a preset as a side effect, so
 *  undoing it needs to remove that preset too, not just the rows it added; every
 *  other undoable action's presets snapshot is identical before/after, so this is a
 *  no-op for them. */
function popUndoAndRestore(projectId, catKey) {
  const key = undoKey(projectId, catKey);
  const stack = state.undoStack[key];
  if (!stack || stack.length === 0) return false;
  const { description, rows: rowsBefore, batches: batchesBefore, presets: presetsBefore } = stack.pop();
  const rowsNow = DataStore.getData(projectId, catKey);
  const batchesNow = DataStore.getImportBatches(projectId, catKey);
  const presetsNow = getItemPresets(catKey);
  if (!state.redoStack[key]) state.redoStack[key] = [];
  state.redoStack[key].push({ description, rows: deepCopy(rowsNow), batches: deepCopy(batchesNow), presets: deepCopy(presetsNow) });
  DataStore.saveData(projectId, catKey, rowsBefore);
  if (batchesBefore) DataStore.saveImportBatches(projectId, catKey, batchesBefore);
  if (presetsBefore) saveItemPresets(catKey, presetsBefore);
  return true;
}
/** Redo: reapplies the most recently undone action, and pushes the state being
 *  moved away from back onto the undo stack, so the redo itself can be undone
 *  again if needed. */
function popRedoAndRestore(projectId, catKey) {
  const key = undoKey(projectId, catKey);
  const stack = state.redoStack[key];
  if (!stack || stack.length === 0) return false;
  const { description, rows: rowsAfter, batches: batchesAfter, presets: presetsAfter } = stack.pop();
  const rowsNow = DataStore.getData(projectId, catKey);
  const batchesNow = DataStore.getImportBatches(projectId, catKey);
  const presetsNow = getItemPresets(catKey);
  if (!state.undoStack[key]) state.undoStack[key] = [];
  state.undoStack[key].push({ description, rows: deepCopy(rowsNow), batches: deepCopy(batchesNow), presets: deepCopy(presetsNow) });
  DataStore.saveData(projectId, catKey, rowsAfter);
  if (batchesAfter) DataStore.saveImportBatches(projectId, catKey, batchesAfter);
  if (presetsAfter) saveItemPresets(catKey, presetsAfter);
  return true;
}

/**
 * Site aliases need a key that survives across quarters — the raw site "code" a
 * report embeds includes the report's own sequence number (e.g. "13545N7-01" for
 * one report vs "13545N8-01" for the very next one, same physical location), which
 * breaks any memory keyed on it the moment a new report comes in — confirmed with
 * real data: every quarter's remembered 管制區/環境音量標準/座標 etc. was silently
 * lost because the lookup key itself changed. The location's own descriptive text
 * (as printed in the report) is what's actually stable season to season.
 *
 * The key also folds in the row's actual 檢測類別 whenever the category has one
 * (every category except 生態) — the most robust scoping, since one physical
 * location can legitimately carry more than one classification within a single
 * import (most obviously noise: a 環境噪音/道路交通噪音 sub-report and a separate
 * 振動 sub-report share a location but need different 管制標準/環境音量標準
 * remembered). Falls back to the bare location text when there's no category field
 * to key on (生態) or no sample row to read one from.
 */
function siteAliasKey(site, result, cat) {
  const sampleRow = result.rows[site.rowIndices[0]];
  const category = (sampleRow && sampleRow['檢測類別']) || '';
  return category ? `${site.rawLocation}::${category}` : (site.rawLocation || '');
}

/**
 * Re-renders the current tab but keeps the data table's scroll position where it
 * was — used after a sync confirmation (coordinates/date-time/category) so the
 * person doesn't get bounced back to the top of a long table and lose their place
 * while working through the rest of the rows.
 */
function renderContentPreservingScroll() {
  const oldWrap = document.querySelector('.table-wrap');
  const scrollTop = oldWrap ? oldWrap.scrollTop : 0;
  const scrollLeft = oldWrap ? oldWrap.scrollLeft : 0;
  const pageScrollY = window.scrollY;
  renderContent();
  const newWrap = document.querySelector('.table-wrap');
  if (newWrap) {
    newWrap.scrollTop = scrollTop;
    newWrap.scrollLeft = scrollLeft;
  }
  window.scrollTo(0, pageScrollY);
}

function renderPeriodPicker(containerId, rows) {
  const guess = guessPeriodFromRows(rows);
  if (!state.importPeriod) state.importPeriod = guess;
  const container = document.getElementById(containerId);

  // Flag when the batch's sampling dates span more than one 年/季 — this usually
  // means a report carries a few rows of older reference/historical data alongside
  // this quarter's real submission (confirmed against a real ground-truth filing:
  // a handful of rows from a prior year got silently excluded from the official
  // submission). The system can't know WHICH rows belong in this filing, but it can
  // reliably flag "these don't all agree on the same quarter" so the person notices
  // and can exclude the odd ones out via the row-detail checklist.
  const periodCounts = {};
  rows.forEach(r => {
    const p = guessPeriodFromRows([r]);
    if (p) periodCounts[p] = (periodCounts[p] || 0) + 1;
  });
  const periodEntries = Object.entries(periodCounts).sort((a, b) => b[1] - a[1]);
  const outlierWarning = periodEntries.length > 1
    ? `⚠️ 偵測到本次資料的採樣日期橫跨不同期別：${periodEntries.map(([p, c]) => `${escapeHtml(p)}（${c}筆）`).join('、')}。如果其中有不屬於本次要送件季度的資料（例如報告裡夾帶的舊資料或參考值），建議到下方「詳細資料列表」展開後取消勾選排除，避免誤送。`
    : '';

  // Warn (not block) when the entered/guessed period already has data in this
  // category — likely a corrected re-import of the same quarter's file. The actual
  // per-row diff-and-choose-which-version handling already happens automatically
  // via the conflict-resolution step after confirming, so this is just visibility:
  // the person doesn't have to guess that re-importing will behave sensibly.
  const project = getImportProject();
  const catKey = state.importCatKey;
  let duplicateWarning = '';
  if (project && catKey && state.importPeriod && getKnownPeriods(project, catKey).includes(state.importPeriod)) {
    duplicateWarning = `ℹ️ 「${escapeHtml(state.importPeriod)}」這個期別先前已經有資料了。若這是同一期的修正版，請放心繼續匯入——確認時系統會自動比對每一筆的差異，列出讓您選擇要保留哪個版本；若只是重複匯入同一份原始檔案，內容完全相同的部分會自動忽略，不會產生重複資料。`;
  }

  const parsed = parsePeriodString(normalizePeriodShorthand(state.importPeriod)) || {};

  container.innerHTML = `
    <label class="period-picker-label">
      本批資料屬於哪一期？
      <input type="number" id="periodRocYearInput" min="1" max="200" placeholder="民國年" value="${escapeAttr(parsed.rocYear || '')}" style="width:64px">
      年第
      <select id="periodQuarterInput" style="width:56px">
        <option value="">-</option>
        ${[1, 2, 3, 4].map(q => `<option value="${q}" ${String(q) === parsed.quarter ? 'selected' : ''}>${q}</option>`).join('')}
      </select>
      季
    </label>
    <label class="period-picker-label">
      或直接輸入：
      <input type="text" id="importPeriodInput" value="${escapeAttr(state.importPeriod)}" placeholder="例：115年第1季 或 115Q1" style="width:160px">
    </label>
    <span class="hint">${guess ? `系統依照資料中的採樣日期猜測為「${escapeHtml(guess)}」，如不正確請直接修改。` : '系統無法從資料中判斷期別，請填上方民國年+季別，或直接輸入文字（含 115Q1 這種簡寫）。'}</span>
    ${outlierWarning ? `<div class="warning" style="width:100%;margin-top:8px">${outlierWarning}</div>` : ''}
    ${duplicateWarning ? `<div class="warning" style="width:100%;margin-top:8px;background:#e8f0fe;border-color:#a8c7fa">${duplicateWarning}</div>` : ''}
  `;

  const yearInput = document.getElementById('periodRocYearInput');
  const quarterInput = document.getElementById('periodQuarterInput');
  const textInput = document.getElementById('importPeriodInput');

  const applyFromDropdown = () => {
    if (yearInput.value && quarterInput.value) {
      state.importPeriod = `${yearInput.value}年第${quarterInput.value}季`;
      textInput.value = state.importPeriod;
    }
  };
  yearInput.addEventListener('input', applyFromDropdown);
  quarterInput.addEventListener('change', applyFromDropdown);

  textInput.addEventListener('input', (e) => { state.importPeriod = e.target.value; });
  textInput.addEventListener('blur', (e) => {
    const normalized = normalizePeriodShorthand(e.target.value);
    state.importPeriod = normalized;
    e.target.value = normalized;
    const p = parsePeriodString(normalized);
    if (p) { yearInput.value = p.rocYear; quarterInput.value = p.quarter; }
  });
}

// ---------- helpers ----------
/** Lightweight, non-blocking notification — fades in, sits for a few seconds, fades
 *  out. Used for reminders that matter but shouldn't force a click to dismiss every
 *  single time (e.g. every export), unlike alert()/confirm(). */
function showToast(message, durationMs = 7000) {
  let container = document.getElementById('toastContainer');
  if (!container) {
    container = document.createElement('div');
    container.id = 'toastContainer';
    document.body.appendChild(container);
  }
  const toast = document.createElement('div');
  toast.className = 'toast';
  toast.innerHTML = message;
  container.appendChild(toast);
  requestAnimationFrame(() => toast.classList.add('toast-visible'));
  setTimeout(() => {
    toast.classList.remove('toast-visible');
    setTimeout(() => toast.remove(), 400);
  }, durationMs);
}

/** Excel-style value comparison for column sort: numeric when both sides parse as
 *  numbers, locale-aware text compare otherwise; blanks always sort to the end
 *  regardless of direction (the caller's direction flip is applied on top of this,
 *  so blanks-last needs to be encoded as a value that survives negation). */
function compareForSort(a, b) {
  const av = (a ?? '').toString().trim();
  const bv = (b ?? '').toString().trim();
  if (av === '' && bv === '') return 0;
  if (av === '') return 1;
  if (bv === '') return -1;
  const an = Number(av), bn = Number(bv);
  if (!isNaN(an) && !isNaN(bn)) return an - bn;
  return av.localeCompare(bv, 'zh-Hant');
}

/** Normalizes a location name for "are these the same site, just typed
 *  differently" comparison: strips all whitespace (incl. full-width), and
 *  collapses full-width/half-width variants of parentheses/dashes/commas to one
 *  form. Two names that are identical after this but NOT identical as typed are
 *  almost certainly the same physical site with a formatting difference — the
 *  exact class of mismatch that broke cross-season memory in earlier testing
 *  (地點名稱不完全一致 due to full/half-width punctuation or stray spaces). */
function normalizeLocationForFuzzyMatch(s) {
  return String(s || '')
    .replace(/[\s\u3000]/g, '')
    .replace(/[（(]/g, '(').replace(/[）)]/g, ')')
    .replace(/[－—–ー-]/g, '-')
    .replace(/[，,]/g, ',')
    .toLowerCase();
}
function levenshteinDistance(a, b) {
  const m = a.length, n = b.length;
  if (m === 0) return n;
  if (n === 0) return m;
  const dp = Array.from({ length: m + 1 }, () => new Array(n + 1).fill(0));
  for (let i = 0; i <= m; i++) dp[i][0] = i;
  for (let j = 0; j <= n; j++) dp[0][j] = j;
  for (let i = 1; i <= m; i++) {
    for (let j = 1; j <= n; j++) {
      dp[i][j] = a[i - 1] === b[j - 1] ? dp[i - 1][j - 1] : 1 + Math.min(dp[i - 1][j], dp[i][j - 1], dp[i - 1][j - 1]);
    }
  }
  return dp[m][n];
}
/** Looks for a historically-known location name that's a near-miss for `rawLoc` —
 *  not an exact match (nothing to flag there), but close enough that it's probably
 *  the same site typed slightly differently rather than a genuinely new site.
 *  Two tiers: (1) identical once whitespace/punctuation differences are normalized
 *  away — very high confidence, this is almost certainly the same site; (2) a
 *  small edit distance relative to length — catches likely single-character typos,
 *  lower confidence but still worth flagging. Returns null if rawLoc already
 *  exactly matches a known location (nothing to ask about) or nothing is close. */
function findSimilarHistoricalLocation(rawLoc, historicalLocs) {
  if (!rawLoc || !historicalLocs || historicalLocs.length === 0) return null;
  if (historicalLocs.includes(rawLoc)) return null;
  const normRaw = normalizeLocationForFuzzyMatch(rawLoc);
  const normMatch = historicalLocs.find(h => normalizeLocationForFuzzyMatch(h) === normRaw);
  if (normMatch) return { match: normMatch, confidence: 'high' };
  let best = null, bestDist = Infinity;
  historicalLocs.forEach(h => {
    const dist = levenshteinDistance(rawLoc, h);
    const threshold = Math.max(1, Math.floor(Math.max(rawLoc.length, h.length) * 0.2));
    if (dist <= threshold && dist < bestDist) { best = h; bestDist = dist; }
  });
  return best ? { match: best, confidence: 'low' } : null;
}

function escapeHtml(s) {
  return String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}
function escapeAttr(s) { return escapeHtml(s); }

// Date/time conversion lives in DateTimeUtil (js/datetime.js) — ONE implementation
// shared by the importers, the grid, the preview tables and the exporter, so a value
// can't be interpreted one way when it's read and another way when it's shown. These
// thin wrappers keep the existing call sites in this file readable.
function toDateInputValue(v) {
  const iso = DateTimeUtil.toISODate(v);
  return /^\d{4}-\d{2}-\d{2}$/.test(iso) ? iso : '';
}
/** ISO "YYYY-MM-DD" (internal storage/export format) -> "YYYY/MM/DD" for display.
 *  Date fields show a calendar date and nothing else — never a time, never seconds. */
function toDateDisplayValue(v) {
  return DateTimeUtil.toDisplayDate(toDateInputValue(v) || v);
}
/** Accepts "2026/5/12", "2026-5-12", "20260512", "115/06/25", "115年6月25日" and even a
 *  raw Excel serial, and normalizes to the canonical ISO "YYYY-MM-DD" used for
 *  storage/export/date-math. Unrecognized input is left as typed so the person can
 *  see and fix it rather than having it silently discarded. */
function normalizeDateString(raw) {
  return DateTimeUtil.toISODate(raw);
}
function toTimeInputValue(v) {
  const hms = DateTimeUtil.toHMS(v);
  return /^\d{2}:\d{2}:\d{2}$/.test(hms) ? hms : '';
}
/** Internal storage format -> "HH:MM" for display. Time fields show hours and minutes
 *  only — the official template's own time format is "h:mm" and no seconds value is
 *  ever meaningful for a sampling time. */
function toTimeDisplayValue(v) {
  return DateTimeUtil.toDisplayTime(toTimeInputValue(v) || v);
}
/** Accepts "1430", "14:30", "14:30:00", "143000", "10時30分", "2:30 PM" typed
 *  free-hand and normalizes to "HH:MM:00" (seconds always dropped). */
function normalizeTimeString(raw) {
  return DateTimeUtil.toHMS(raw);
}
function lookupUnit(code) {
  if (!code) return '';
  return UNIT_CODES[String(code).trim()] ? `代碼 ${code}：${UNIT_CODES[String(code).trim()]}` : '（找不到對應單位代碼，請確認）';
}
function lookupAgency(code) {
  if (!code) return '';
  const parts = String(code).split(';').map(s => s.trim()).filter(Boolean);
  return parts.map(c => AGENCY_CODES[c] ? `${c}：${AGENCY_CODES[c]}` : `${c}：（找不到對應機構代碼）`).join('　|　');
}

function getCurrentProject() {
  return DataStore.getProjects().find(p => p.id === state.currentProjectId) || null;
}

/** The project a currently-open import is locked to (captured in openImportModal at
 *  the moment the modal opened) — used by every import-flow function (preview
 *  render, confirm, history learning) instead of a live getCurrentProject() lookup,
 *  so switching projects in the sidebar while an import is still in progress can
 *  never cause the parsed file to land in the wrong project. */
function getImportProject() {
  return DataStore.getProjects().find(p => p.id === state.importProjectId) || null;
}

// ---------- project list ----------
function renderProjectList() {
  const list = document.getElementById('projectList');
  const projects = DataStore.getProjects();
  renderBackupReminder();
  if (projects.length === 0) {
    list.innerHTML = '<li class="hint" style="padding:10px">尚無計畫，請點上方「＋ 新增計畫」建立</li>';
    return;
  }
  list.innerHTML = projects.map(p => `
    <li class="project-item ${p.id === state.currentProjectId ? 'active' : ''}" data-id="${p.id}">
      <div class="project-item-main">
        <div class="p-code">${escapeHtml(p.code)}</div>
        <div class="p-name">${escapeHtml(p.name)}</div>
      </div>
      <button class="project-del-btn" data-id="${p.id}" title="刪除此計畫">🗑</button>
    </li>
  `).join('');
  list.querySelectorAll('.project-item-main').forEach(el => {
    el.addEventListener('click', () => selectProject(el.closest('.project-item').dataset.id));
  });
  list.querySelectorAll('.project-del-btn').forEach(btn => {
    btn.addEventListener('click', (e) => {
      e.stopPropagation();
      const project = DataStore.getProjects().find(p => p.id === btn.dataset.id);
      if (project) deleteProjectFlow(project);
    });
  });
}

/** Returns true if it's safe to change state.currentProjectId to `newProjectId`
 *  right now — false if the person was asked and chose to keep their in-progress
 *  import instead. Centralizes the "don't silently switch away from the project an
 *  open import is locked to" guard so EVERY code path that can change the active
 *  project — not just clicking a project in the sidebar, but also creating a brand
 *  new project while an import is open — gets the same protection. Data safety
 *  itself never depended on this (see getImportProject), but silently switching the
 *  screen out from under an in-progress import is still confusing and worth
 *  confirming rather than allowing unannounced. */
function confirmProjectSwitchIfImporting(newProjectId) {
  const importModal = document.getElementById('importModal');
  if (importModal && !importModal.classList.contains('hidden') && state.importProjectId && state.importProjectId !== newProjectId) {
    if (!confirm('目前正在匯入資料尚未完成，切換計畫會取消這次匯入。確定要放棄目前的匯入並切換嗎？')) return false;
    closeImportModal();
  }
  return true;
}
function selectProject(id) {
  if (!confirmProjectSwitchIfImporting(id)) return;
  state.currentProjectId = id;
  state.currentTab = 'basic';
  renderProjectList();
  renderContent();
}

// ---------- content area ----------
function renderContent() {
  const content = document.getElementById('content');
  const project = getCurrentProject();
  if (!project) {
    content.innerHTML = '<div class="empty-state"><p>👈 請先在左側建立或選擇一個計畫</p></div>';
    return;
  }

  const tabsHtml = `
    <button class="tab-btn ${state.currentTab === 'basic' ? 'active' : ''}" data-tab="basic">監測點基本資料</button>
    ${CATEGORY_ORDER.map(catKey => {
      const cat = CATEGORIES[catKey];
      const n = DataStore.getData(project.id, catKey).length;
      return `<button class="tab-btn ${state.currentTab === catKey ? 'active' : ''}" data-tab="${catKey}">${cat.label}${n ? `<span class="count">${n}</span>` : ''}</button>`;
    }).join('')}
  `;

  content.innerHTML = `
    <div class="project-header">
      <div>
        <h1>${escapeHtml(project.code)}　${escapeHtml(project.name)}</h1>
        <div class="sub">建立於 ${new Date(project.createdAt).toLocaleDateString('zh-TW')}</div>
      </div>
      <div class="project-header-actions">
        <button class="btn btn-ghost btn-sm" id="btnEditProject">編輯計畫資訊</button>
        <button class="btn btn-danger btn-sm" id="btnDeleteProject">刪除計畫</button>
        <button class="btn btn-ghost btn-sm" id="btnBatchImport">📦 批次匯入監測報告</button>
        <button class="btn btn-primary btn-sm" id="btnExportAll">匯出全部類別</button>
      </div>
    </div>
    <div class="tabs">${tabsHtml}</div>
    <div id="tabBody"></div>
  `;

  content.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => { state.currentTab = btn.dataset.tab; renderContent(); });
  });
  document.getElementById('btnEditProject').addEventListener('click', () => openProjectModal(project));
  document.getElementById('btnDeleteProject').addEventListener('click', () => deleteProjectFlow(project));
  document.getElementById('btnBatchImport').addEventListener('click', openBatchImportModal);
  document.getElementById('btnExportAll').addEventListener('click', () => {
    const anyData = CATEGORY_ORDER.some(c => DataStore.getData(project.id, c).length > 0);
    if (!anyData) { alert('目前尚無任何監測資料可匯出，請先匯入或新增資料。'); return; }
    openExportSelectModal(project);
  });

  if (state.currentTab === 'basic') renderBasicTab(project);
  else renderCategoryTab(project, state.currentTab);
}

// ---------- basic info tab ----------
function renderBasicTab(project) {
  const body = document.getElementById('tabBody');
  const info = DataStore.getBasicInfo(project.id);

  const fieldsHtml = BASIC_INFO_FIELDS.map(f => {
    const val = info[f.key] ?? '';
    let control;
    if (f.type === 'select') {
      control = `<select data-field="${f.key}">${f.options.map(o => `<option value="${escapeAttr(o)}" ${o === val ? 'selected' : ''}>${o || '（未選擇）'}</option>`).join('')}</select>`;
    } else if (f.type === 'date') {
      control = `<input type="text" data-field="${f.key}" value="${escapeAttr(toDateDisplayValue(val))}" class="date-input" placeholder="YYYY/MM/DD" inputmode="numeric" maxlength="10">`;
    } else if (f.type === 'textarea') {
      control = `<textarea data-field="${f.key}" rows="2">${escapeHtml(val)}</textarea>`;
    } else {
      control = `<input type="text" data-field="${f.key}" value="${escapeAttr(val)}">`;
    }
    const wide = (f.type === 'textarea' || f.key === '書件名稱') ? 'full' : '';
    return `<label class="${wide}">${f.label}${f.required ? '<span class="req">＊</span>' : ''}
      ${control}
      ${f.help ? `<span class="field-help">${escapeHtml(f.help)}</span>` : ''}
    </label>`;
  }).join('');

  body.innerHTML = `
    <div class="panel">
      <p class="hint" style="margin-top:0">此頁資料會套用到「全部監測類別」匯出檔案中的「監測點基本資料」工作表，可隨時補充修改。</p>
      <div class="basic-form">${fieldsHtml}</div>
    </div>
  `;

  body.querySelectorAll('[data-field]').forEach(el => {
    el.addEventListener('input', () => {
      const info = DataStore.getBasicInfo(project.id);
      info[el.dataset.field] = el.value;
      DataStore.saveBasicInfo(project.id, info);
    });
    el.addEventListener('change', () => {
      const info = DataStore.getBasicInfo(project.id);
      info[el.dataset.field] = el.value;
      DataStore.saveBasicInfo(project.id, info);
    });
    if (el.classList.contains('date-input')) {
      el.addEventListener('focusout', () => {
        const normalized = normalizeDateString(el.value);
        el.value = toDateDisplayValue(normalized) || normalized;
        const info = DataStore.getBasicInfo(project.id);
        info[el.dataset.field] = normalized;
        DataStore.saveBasicInfo(project.id, info);
      });
      el.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); el.blur(); } });
    }
  });
}

// ---------- category data tab ----------
function renderCategoryTab(project, catKey) {
  const cat = CATEGORIES[catKey];
  const allRows = DataStore.getData(project.id, catKey);
  const batches = DataStore.getImportBatches(project.id, catKey);
  const body = document.getElementById('tabBody');

  const knownPeriods = getKnownPeriods(project, catKey);
  const hasUnlabeled = allRows.some(r => !r._period);
  if (state.periodFilter[catKey] === undefined) state.periodFilter[catKey] = '';
  const activePeriod = state.periodFilter[catKey];
  const showPeriodUI = knownPeriods.length > 0 || hasUnlabeled;

  // keep each row's ORIGINAL array index (not display order) so edits/deletes still
  // target the right record in DataStore after filtering by period for display.
  const displayEntries = allRows
    .map((row, idx) => ({ row, idx }))
    .filter(({ row }) => {
      if (!activePeriod) return true;
      if (activePeriod === '__none__') return !row._period;
      return row._period === activePeriod;
    });

  // Excel-style column sort — reorders the array itself (not a DOM show/hide trick
  // like search/column-filter use) since every row keeps carrying its ORIGINAL idx
  // regardless of display order, so edits/deletes still land on the right record.
  const sortState = state.columnSort[catKey];
  if (sortState) {
    const dir = sortState.direction === 'desc' ? -1 : 1;
    displayEntries.sort((a, b) => dir * compareForSort(a.row[sortState.fieldKey], b.row[sortState.fieldKey]));
  }

  const displayRows = displayEntries.map(e => e.row);

  body.innerHTML = `
    ${catKey === 'eco' ? `<div class="warning warning-strong" style="margin-bottom:10px">🌿 ${ECO_IMPORT_ONLY_NOTE}</div>` : ''}
    <div class="toolbar">
      <div class="toolbar-left">
        <button class="btn btn-primary btn-sm" id="btnImport">📥 匯入資料（Excel/PDF）</button>
        <button class="btn btn-ghost btn-sm" id="btnAddRow">＋ 新增一筆</button>
        <button class="btn btn-ghost btn-sm" id="btnAddCommonItems">☑️ 常用測項新增</button>
        <button class="btn btn-ghost btn-sm" id="btnCoordManager">📍 測站座標管理</button>
        ${cat.methodField ? `<button class="btn btn-ghost btn-sm" id="btnMethodManager">🧪 檢測方法管理</button>` : ''}
        <button class="btn btn-ghost btn-sm" id="btnBatchHistory">📜 匯入紀錄${batches.length ? ` (${batches.length})` : ''}</button>
        ${showPeriodUI ? `
        <select id="periodFilterSelect" title="篩選要查看／編輯／匯出哪一期的資料；「匯出此類別」會依此篩選範圍匯出">
          <option value="" ${activePeriod === '' ? 'selected' : ''}>全部期別</option>
          ${knownPeriods.map(p => `<option value="${escapeAttr(p)}" ${activePeriod === p ? 'selected' : ''}>${escapeHtml(p)}</option>`).join('')}
          ${hasUnlabeled ? `<option value="__none__" ${activePeriod === '__none__' ? 'selected' : ''}>未標示期別</option>` : ''}
        </select>` : ''}
        <input type="text" id="rowSearchInput" placeholder="🔍 搜尋任何欄位內容（測站、測項、日期…）" title="輸入關鍵字即時篩選畫面顯示的資料列，方便檢查或修正特定資料；不影響匯出範圍（匯出仍依上方期別篩選）">
        <button class="btn btn-ghost btn-sm hidden" id="btnClearSearch">✕ 清除篩選</button>
        <button class="btn btn-ghost btn-sm" id="btnExportCat">匯出此類別（${cat.sourceFile}）${activePeriod ? '（僅目前篩選期別）' : ''}</button>
        <button class="btn btn-danger btn-sm" id="btnClearCat" ${allRows.length === 0 ? 'disabled' : ''}>🗑 清空此類別${activePeriod ? '（僅目前篩選期別）' : ''}</button>
        ${(() => { const u = peekUndo(project.id, catKey); return u ? `<button class="btn btn-ghost btn-sm" id="btnUndo" title="復原：${escapeAttr(u.description)}">↩️ 復原上一步</button>` : ''; })()}
        ${(() => { const r = peekRedo(project.id, catKey); return r ? `<button class="btn btn-ghost btn-sm" id="btnRedo" title="重做：${escapeAttr(r.description)}">↪️ 返回下一步</button>` : ''; })()}
      </div>
      <div class="row-count" id="rowCountDisplay">共 ${displayRows.length} 筆資料${activePeriod ? `（篩選中，全部共 ${allRows.length} 筆）` : ''}</div>
    </div>
    <div class="toolbar bulk-toolbar hidden" id="bulkToolbar">
      <span id="bulkSelCount">已選取 0 筆</span>
      <button class="btn btn-primary btn-sm" id="btnBulkEdit">✏️ 批次修改欄位</button>
      <button class="btn btn-danger btn-sm" id="btnBulkDelete">🗑 刪除已選取</button>
      <button class="btn btn-ghost btn-sm" id="btnBulkClear">取消選取</button>
    </div>
    <div class="table-wrap">
      <table class="data-grid">
        <thead><tr>
          <th class="col-check">${checkboxHTML(` id="checkAllRows" ${displayRows.length === 0 ? 'disabled' : ''}`)}</th>
          ${(() => {
            const renderHeaderField = (f) => {
              const activeFilter = state.columnFilters[catKey]?.[f.key];
              const filterActive = activeFilter && activeFilter.size > 0;
              const sortState = state.columnSort[catKey];
              const isSorted = sortState && sortState.fieldKey === f.key;
              // Always show SOME sort affordance next to the label — not just on
              // hover (which is invisible on touch devices, and easy to miss even
              // with a mouse if the person doesn't happen to hover before
              // clicking). A neutral "⇅" hints "click to sort" for unsorted
              // columns; an active sort shows its direction instead.
              const sortIcon = isSorted
                ? `<span class="sort-indicator sort-active">${sortState.direction === 'asc' ? '▲' : '▼'}</span>`
                : `<span class="sort-indicator">⇅</span>`;
              return `<th${f.key === cat.itemField ? ' class="col-item"' : f.key === cat.locationField ? ' class="col-loc"' : ''}${f.help ? ` title="${escapeAttr(f.help)}"` : ''}>
                <span class="th-label th-sortable" data-sort-field="${escapeAttr(f.key)}" title="點擊依此欄位排序">${escapeHtml(f.label)}${f.required ? '<span class="req">＊</span>' : ''}${f.help ? ' ℹ️' : ''}${sortIcon}</span>
                <button class="col-filter-btn${filterActive ? ' col-filter-active' : ''}" data-field-key="${escapeAttr(f.key)}" title="篩選「${escapeAttr(f.label)}」">▾</button>
              </th>`;
            };
            const order = displayFieldOrder(cat);
            const pinnedHeaders = order.slice(0, 2).map(renderHeaderField).join('');
            const restHeaders = order.slice(2).map(renderHeaderField).join('');
            // Same reordering as rowHtml: 操作/# sit right after the pinned 地點/測項
            // headers, before the rest — see the comment in rowHtml for why.
            return `${pinnedHeaders}<th class="col-actions">操作</th><th class="col-num">#</th>${restHeaders}`;
          })()}
        </tr></thead>
        <tbody id="gridBody">${displayEntries.map(({ row, idx }) => rowHtml(cat, row, idx)).join('')}</tbody>
      </table>
    </div>
    ${allRows.length === 0 ? '<p class="hint" style="margin-top:10px">尚無資料。可點「匯入資料」上傳該類別的檢測結果檔案，或「新增一筆」手動輸入。</p>' : ''}
    ${allRows.length > 0 && displayRows.length === 0 ? '<p class="hint" style="margin-top:10px">此期別目前沒有資料。</p>' : ''}
  `;

  document.getElementById('btnImport').addEventListener('click', () => openImportModal(catKey));
  document.getElementById('btnAddRow').addEventListener('click', () => addEmptyRow(project, catKey));
  document.getElementById('btnAddCommonItems').addEventListener('click', () => openCommonItemsModal(project, catKey));
  document.getElementById('btnCoordManager').addEventListener('click', () => openCoordModal(project, catKey));
  if (cat.methodField) {
    document.getElementById('btnMethodManager').addEventListener('click', () => openMethodModal(project, catKey));
  }
  document.getElementById('btnBatchHistory').addEventListener('click', () => openBatchHistoryModal(project, catKey));
  if (showPeriodUI) {
    document.getElementById('periodFilterSelect').addEventListener('change', (e) => {
      state.periodFilter[catKey] = e.target.value;
      renderContent();
    });
  }
  document.getElementById('btnExportCat').addEventListener('click', () => {
    if (displayRows.length === 0) { alert('目前篩選範圍內沒有資料可匯出。'); return; }
    ExportEngine.downloadCategory(project, DataStore.getBasicInfo(project.id), catKey, displayRows);
    showToast('📥 已匯出。<strong>若之後修改這份檔案</strong>（手動補值、新增測項等），完成後可到本頁上方點原本的「📥 匯入資料」按鈕、選擇這份修改過的檔案<strong>重新匯入即可</strong>——系統會自動比對，內容相同的資料會忽略，內容不同的會列出讓您確認要不要用新版本取代，不會憑空覆蓋或造成重複。', 9000);
  });
  document.getElementById('btnClearCat').addEventListener('click', () => {
    if (allRows.length === 0) return;
    if (activePeriod) {
      const periodLabel = activePeriod === '__none__' ? '未標示期別' : activePeriod;
      const rowsToKeep = allRows.filter(r => activePeriod === '__none__' ? !!r._period : r._period !== activePeriod);
      const removedCount = allRows.length - rowsToKeep.length;
      if (!confirm(`確定要清空「${cat.label}」目前篩選的「${periodLabel}」共 ${removedCount} 筆資料嗎？其他期別的資料不會受影響。（可用「↩️ 復原上一步」救回）`)) return;
      pushUndoSnapshot(project.id, catKey, `清空「${periodLabel}」（${removedCount}筆）`);
      DataStore.saveData(project.id, catKey, rowsToKeep);
      state.periodFilter[catKey] = '';
    } else {
      if (!confirm(`確定要清空「${cat.label}」的全部 ${allRows.length} 筆資料嗎？（含所有期別，可用「↩️ 復原上一步」救回）`)) return;
      pushUndoSnapshot(project.id, catKey, `清空全部（${allRows.length}筆）`);
      DataStore.clearData(project.id, catKey);
    }
    renderContent();
  });
  const btnUndo = document.getElementById('btnUndo');
  if (btnUndo) {
    btnUndo.addEventListener('click', () => {
      const snapshot = peekUndo(project.id, catKey);
      if (!snapshot) return;
      if (!confirm(`確定要復原「${snapshot.description}」嗎？`)) return;
      popUndoAndRestore(project.id, catKey);
      renderContent();
    });
  }
  const btnRedo = document.getElementById('btnRedo');
  if (btnRedo) {
    btnRedo.addEventListener('click', () => {
      const snapshot = peekRedo(project.id, catKey);
      if (!snapshot) return;
      if (!confirm(`確定要重做「${snapshot.description}」嗎？`)) return;
      popRedoAndRestore(project.id, catKey);
      renderContent();
    });
  }

  wireGridEvents(project, catKey, cat);
  wireBulkSelection(project, catKey, cat);
  wireRowSearch(catKey, cat, displayRows.length, allRows.length, activePeriod);

  document.querySelectorAll('.th-sortable').forEach(el => {
    el.addEventListener('click', () => {
      const fieldKey = el.dataset.sortField;
      const current = state.columnSort[catKey];
      if (!current || current.fieldKey !== fieldKey) {
        state.columnSort[catKey] = { fieldKey, direction: 'asc' };
      } else if (current.direction === 'asc') {
        state.columnSort[catKey] = { fieldKey, direction: 'desc' };
      } else {
        delete state.columnSort[catKey];
      }
      renderContentPreservingScroll();
    });
  });
}

// ---------- 表格篩選：文字搜尋 + Excel風格欄位篩選（皆為畫面篩選，不影響匯出範圍——匯出範圍由上方期別篩選決定） ----------
function getRowFieldDisplayValue(tr, fieldKey) {
  const el = tr.querySelector(`[data-field="${CSS.escape(fieldKey)}"]`);
  if (!el) return '';
  if (el.tagName === 'SELECT') {
    const opt = el.options[el.selectedIndex];
    return opt ? opt.text : (el.value || '');
  }
  return el.value ?? '';
}

function wireRowSearch(catKey, cat, displayCount, totalCount, activePeriod) {
  const searchInput = document.getElementById('rowSearchInput');
  const clearBtn = document.getElementById('btnClearSearch');
  const tbody = document.getElementById('gridBody');
  const countEl = document.getElementById('rowCountDisplay');
  if (!searchInput || !tbody) return;

  const baseCountText = `共 ${displayCount} 筆資料${activePeriod ? `（篩選中，全部共 ${totalCount} 筆）` : ''}`;

  // Search text is built from each cell's ACTUAL current value — not raw DOM
  // textContent, which for a <select> silently includes every hidden <option> label
  // (not just the selected one), causing a search term to match rows regardless of
  // what's actually selected there.
  const rowSearchText = (tr) => {
    const parts = [];
    tr.querySelectorAll('td').forEach(td => {
      const select = td.querySelector('select');
      const input = td.querySelector('input:not([type="checkbox"])');
      if (select) {
        const opt = select.options[select.selectedIndex];
        parts.push(opt ? opt.text : select.value);
      } else if (input) {
        parts.push(input.value || '');
      } else {
        parts.push(td.textContent || '');
      }
    });
    return parts.join(' ').toLowerCase();
  };

  const colFiltersForCat = () => state.columnFilters[catKey] || {};

  const applyAllFilters = () => {
    const q = searchInput.value.trim().toLowerCase();
    const colFilters = colFiltersForCat();
    const anyColFilter = Object.values(colFilters).some(s => s && s.size > 0);
    clearBtn.classList.toggle('hidden', q === '' && !anyColFilter);
    let visible = 0;
    const rows = tbody.querySelectorAll('tr');
    rows.forEach(tr => {
      let matches = !q || rowSearchText(tr).includes(q);
      if (matches && anyColFilter) {
        for (const [fieldKey, allowedSet] of Object.entries(colFilters)) {
          if (!allowedSet || allowedSet.size === 0) continue;
          if (!allowedSet.has(getRowFieldDisplayValue(tr, fieldKey))) { matches = false; break; }
        }
      }
      tr.classList.toggle('row-hidden', !matches);
      if (matches) visible++;
    });
    countEl.textContent = (q || anyColFilter) ? `符合篩選：${visible} / ${rows.length} 筆` : baseCountText;
  };

  searchInput.addEventListener('input', applyAllFilters);
  clearBtn.addEventListener('click', () => {
    searchInput.value = '';
    state.columnFilters[catKey] = {};
    renderContent();
  });

  // Scoped to #gridBody's own <thead> siblings (the main data table itself), not a
  // bare document-wide selector — a global `.col-filter-btn` query here would also
  // pick up the SAME class used by the 常用測項新增 table's column-filter buttons
  // (shared styling, deliberately reused), double-binding a click handler onto
  // those buttons that calls this function with the wrong arguments and throws.
  // Scoping to the table this function actually belongs to avoids that collision
  // entirely, regardless of what else on the page happens to reuse the class name.
  document.querySelectorAll('table.data-grid .col-filter-btn').forEach(btn => {
    btn.addEventListener('click', (e) => {
      e.stopPropagation();
      openColumnFilterPopup(catKey, fieldKeyOf(btn), btn);
    });
  });

  applyAllFilters();
}

function fieldKeyOf(btn) { return btn.dataset.fieldKey; }

/** Excel-style AutoFilter dropdown: lists every distinct value currently shown in a
 *  column (as actually displayed — select labels, not codes) with checkboxes. */
function openColumnFilterPopup(catKey, fieldKey, btnEl) {
  const tbody = document.getElementById('gridBody');
  const popup = document.getElementById('colFilterPopup');
  const allTrs = [...tbody.querySelectorAll('tr')];
  const uniqueValues = [...new Set(allTrs.map(tr => getRowFieldDisplayValue(tr, fieldKey)))]
    .sort((a, b) => a.localeCompare(b, 'zh-Hant'));
  const currentFilter = state.columnFilters[catKey]?.[fieldKey]; // Set or undefined (undefined = show all)

  popup.innerHTML = `
    <div class="col-filter-search"><input type="text" id="colFilterSearchInput" placeholder="搜尋選項..."></div>
    <div class="col-filter-actions">
      <button type="button" id="colFilterSelectAll" class="btn btn-ghost btn-sm">全選</button>
      <button type="button" id="colFilterClearAll" class="btn btn-ghost btn-sm">清除</button>
    </div>
    <div class="col-filter-list" id="colFilterList">
      ${uniqueValues.map(v => `
        <label class="col-filter-item">
          <input type="checkbox" value="${escapeAttr(v)}" ${!currentFilter || currentFilter.has(v) ? 'checked' : ''}>
          <span>${escapeHtml(v === '' ? '（空白）' : v)}</span>
        </label>
      `).join('')}
    </div>
    <div class="col-filter-buttons">
      <button type="button" id="colFilterCancel" class="btn btn-ghost btn-sm">取消</button>
      <button type="button" id="colFilterApply" class="btn btn-primary btn-sm">確定</button>
    </div>
  `;

  const rect = btnEl.getBoundingClientRect();
  const popupWidth = 240;
  popup.style.left = Math.max(4, Math.min(rect.left, window.innerWidth - popupWidth - 4)) + 'px';
  popup.style.top = (rect.bottom + window.scrollY + 4) + 'px';
  popup.classList.remove('hidden');

  document.getElementById('colFilterSearchInput').addEventListener('input', (e) => {
    const q = e.target.value.trim().toLowerCase();
    document.querySelectorAll('#colFilterList .col-filter-item').forEach(label => {
      label.style.display = (!q || label.textContent.toLowerCase().includes(q)) ? '' : 'none';
    });
  });
  document.getElementById('colFilterSelectAll').addEventListener('click', () => {
    document.querySelectorAll('#colFilterList input[type="checkbox"]').forEach(cb => { cb.checked = true; });
  });
  document.getElementById('colFilterClearAll').addEventListener('click', () => {
    document.querySelectorAll('#colFilterList input[type="checkbox"]').forEach(cb => { cb.checked = false; });
  });
  document.getElementById('colFilterCancel').addEventListener('click', () => {
    popup.classList.add('hidden');
  });
  document.getElementById('colFilterApply').addEventListener('click', () => {
    const checked = [...document.querySelectorAll('#colFilterList input[type="checkbox"]:checked')].map(cb => cb.value);
    if (!state.columnFilters[catKey]) state.columnFilters[catKey] = {};
    if (checked.length === uniqueValues.length) {
      delete state.columnFilters[catKey][fieldKey]; // everything checked = no filter
    } else {
      state.columnFilters[catKey][fieldKey] = new Set(checked);
    }
    popup.classList.add('hidden');
    renderContentPreservingScroll();
  });

  setTimeout(() => {
    const closeOnOutsideClick = (e) => {
      if (!popup.contains(e.target) && e.target !== btnEl) {
        popup.classList.add('hidden');
        document.removeEventListener('click', closeOnOutsideClick);
      }
    };
    document.addEventListener('click', closeOnOutsideClick);
  }, 0);
}

function wireBulkSelection(project, catKey, cat) {
  const tbody = document.getElementById('gridBody');
  const checkAll = document.getElementById('checkAllRows');
  const bulkToolbar = document.getElementById('bulkToolbar');
  const bulkCount = document.getElementById('bulkSelCount');

  // "全選"／批次刪除只作用於目前畫面上看得到的列（考慮搜尋篩選後被隱藏的列），
  // 避免使用者在打了關鍵字、只看到部分資料時，誤選/誤刪看不到的其他資料。
  const isVisible = (cb) => !cb.closest('tr').classList.contains('row-hidden');
  const getVisibleCheckboxes = () => [...tbody.querySelectorAll('.row-check')].filter(isVisible);
  const getChecked = () => getVisibleCheckboxes().filter(cb => cb.checked).map(cb => Number(cb.dataset.row));
  const updateBulkUI = () => {
    const n = getChecked().length;
    bulkToolbar.classList.toggle('hidden', n === 0);
    bulkCount.textContent = `已選取 ${n} 筆`;
    if (checkAll) checkAll.checked = n > 0 && n === getVisibleCheckboxes().length;
  };

  tbody.addEventListener('change', (e) => {
    if (e.target.classList.contains('row-check')) updateBulkUI();
  });
  if (checkAll) {
    checkAll.addEventListener('change', () => {
      getVisibleCheckboxes().forEach(cb => { cb.checked = checkAll.checked; });
      updateBulkUI();
    });
  }
  const bulkDeleteBtn = document.getElementById('btnBulkDelete');
  if (bulkDeleteBtn) {
    bulkDeleteBtn.addEventListener('click', () => {
      const indices = new Set(getChecked());
      if (indices.size === 0) return;
      if (!confirm(`確定要刪除已選取的 ${indices.size} 筆資料嗎？（可用「↩️ 復原上一步」救回）`)) return;
      pushUndoSnapshot(project.id, catKey, `刪除已選取（${indices.size}筆）`);
      const rows = DataStore.getData(project.id, catKey).filter((_, i) => !indices.has(i));
      DataStore.saveData(project.id, catKey, rows);
      renderContent();
    });
  }
  const bulkClearBtn = document.getElementById('btnBulkClear');
  if (bulkClearBtn) {
    bulkClearBtn.addEventListener('click', () => {
      tbody.querySelectorAll('.row-check').forEach(cb => { cb.checked = false; });
      updateBulkUI();
    });
  }
  const bulkEditBtn = document.getElementById('btnBulkEdit');
  if (bulkEditBtn) {
    bulkEditBtn.addEventListener('click', () => {
      const indices = getChecked();
      if (indices.length === 0) return;
      openBatchEditModal(project, catKey, cat, indices);
    });
  }
}

/**
 * Lets the person set ONE field to the same value across every currently-selected
 * (checked + visible under the current search/filter) row at once — e.g. fixing a
 * misspelled site name across many rows, or setting the same 檢測方法 for a batch
 * of rows in one go, instead of editing each cell individually. Excludes the actual
 * measurement value fields (檢測數值/監測數值) — those should always be typed per
 * row, batch-setting them to one shared value is far more likely to be a mistake
 * than something genuinely intended.
 */
function openBatchEditModal(project, catKey, cat, indices) {
  const excludedFields = new Set(['檢測數值', '監測數值']);
  const editableFields = cat.fields.filter(f => !excludedFields.has(f.key));

  document.getElementById('batchEditSummary').textContent =
    `即將修改已選取的 ${indices.length} 筆資料，其餘未選取的資料不受影響。`;

  const fieldSelect = document.getElementById('batchEditFieldSelect');
  fieldSelect.innerHTML = editableFields.map(f => `<option value="${escapeAttr(f.key)}">${escapeHtml(f.label)}</option>`).join('');

  const renderValueInput = () => {
    const field = editableFields.find(f => f.key === fieldSelect.value);
    document.getElementById('batchEditValueWrap').innerHTML = fieldControlHTML(field, '', '');
  };
  renderValueInput();
  fieldSelect.onchange = renderValueInput;

  const modal = document.getElementById('batchEditModal');
  modal.classList.remove('hidden');

  document.getElementById('btnBatchEditCancel').onclick = () => modal.classList.add('hidden');
  document.getElementById('btnBatchEditApply').onclick = () => {
    const field = editableFields.find(f => f.key === fieldSelect.value);
    const inputEl = document.querySelector('#batchEditValueWrap [data-field]');
    let newValue = inputEl.value;
    if (field.type === 'date') newValue = normalizeDateString(newValue);
    if (field.type === 'time') newValue = normalizeTimeString(newValue);

    if (!confirm(`確定要把已選取的 ${indices.length} 筆資料的「${field.label}」統一改成「${newValue || '（空白）'}」嗎？（可用「↩️ 復原上一步」救回）`)) return;

    pushUndoSnapshot(project.id, catKey, `批次修改「${field.label}」（${indices.length}筆）`);
    const rows = DataStore.getData(project.id, catKey);
    const touchedRows = [];
    indices.forEach(idx => {
      if (rows[idx]) { rows[idx][field.key] = newValue; touchedRows.push(rows[idx]); }
    });
    DataStore.saveData(project.id, catKey, rows);
    if (touchedRows.length > 0) learnSiteItemHistory(project.id, catKey, cat, touchedRows);
    modal.classList.add('hidden');
    renderContent();
  };
}

// ---------- import batch history ----------
function openBatchHistoryModal(project, catKey) {
  const cat = CATEGORIES[catKey];
  const batches = DataStore.getImportBatches(project.id, catKey);
  const wrap = document.getElementById('batchHistoryWrap');

  if (batches.length === 0) {
    wrap.innerHTML = '<p class="hint" style="padding:14px">此類別目前沒有透過「匯入」建立的紀錄（手動新增的資料不會列在這裡）。</p>';
  } else {
    wrap.innerHTML = `<table class="mapping-table">
      <thead><tr><th>期別</th><th>匯入時間</th><th>來源檔案</th><th>方式</th><th>筆數</th><th>操作</th></tr></thead>
      <tbody>
        ${batches.slice().reverse().map(b => `
          <tr data-batch-id="${escapeAttr(b.id)}">
            <td>${escapeHtml(b.period || '（未標示）')}</td>
            <td>${new Date(b.timestamp).toLocaleString('zh-TW')}</td>
            <td>${escapeHtml(b.sourceLabel)}</td>
            <td>${b.mode === 'smart' ? '智慧解析' : '一般欄位比對'}</td>
            <td>${b.rowCount}</td>
            <td><button class="btn btn-danger btn-sm btn-delete-batch" data-batch-id="${escapeAttr(b.id)}">🗑 刪除此批次</button></td>
          </tr>
        `).join('')}
      </tbody>
    </table>`;
    wrap.querySelectorAll('.btn-delete-batch').forEach(btn => {
      btn.addEventListener('click', () => {
        const batchId = btn.dataset.batchId;
        const row = batches.find(b => b.id === batchId);
        if (!confirm(`確定要刪除這批匯入資料嗎？（來源：${row.sourceLabel}，共 ${row.rowCount} 筆，可到分類頁面用「↩️ 復原上一步」救回）`)) return;
        pushUndoSnapshot(project.id, catKey, `刪除匯入批次「${row.sourceLabel}」（${row.rowCount}筆）`);
        DataStore.deleteImportBatch(project.id, catKey, batchId);
        openBatchHistoryModal(project, catKey); // refresh list in place
        renderContent();
      });
    });
  }
  document.getElementById('batchHistoryModal').classList.remove('hidden');
}

/** Reorders fields for on-screen DISPLAY ONLY — never touches cat.fields itself,
 *  which must stay in the official schema order everywhere else (export, data
 *  storage, mapping tables) — so the pinned/sticky columns (地點, 測項) are truly
 *  DOM-adjacent to the other pinned columns (勾選, 操作). This matters because
 *  position:sticky's `left` offset only works correctly when the pinned columns
 *  are contiguous in actual document order. Confirmed as a real bug: 採樣地點 sits
 *  5th in the schema (after 日期/時間 columns), so its CSS `left:78px` — written
 *  assuming it directly follows 操作 — didn't match where it actually sat in the
 *  DOM, breaking sticky positioning enough that the checkbox column effectively
 *  disappeared while scrolling right. */
/**
 * Renders a checkbox as a hidden native <input> (fully functional — .checked,
 * change events, :disabled all work exactly as before) paired with a plain CSS
 * box drawn as a sibling <span>. This exists specifically because the native
 * checkbox's own rendering, inside a position:sticky cell, was confirmed to
 * visually vanish while scrolled (while remaining fully clickable underneath) in
 * real testing — a browser compositing quirk that persisted even after reducing
 * the sticky column count. Text content in the OTHER sticky columns (地點/測項)
 * never had this problem, which is the tell: plain box/text rendering is stable
 * under sticky positioning, but the native checkbox WIDGET specifically was not.
 * Drawing the visible box ourselves with a `<span>` and `border`/`background`
 * sidesteps whatever internal compositing path was failing, since a styled span
 * renders exactly the same simple way as any of the text that was never a
 * problem. The real `<input>` stays functionally present (opacity:0, same
 * position/size) so every existing .checked/change-event/:checked selector, and
 * every test, keeps working unchanged — only the visual box is new.
 */
function checkboxHTML(extraAttrs, checked) {
  return `<label class="cb-wrap"><input type="checkbox"${extraAttrs}${checked ? ' checked' : ''}><span class="cb-box"></span></label>`;
}

function displayFieldOrder(cat) {
  const locField = cat.locationField;
  const itemField = cat.itemField;
  return [
    ...cat.fields.filter(f => f.key === locField),
    ...cat.fields.filter(f => f.key === itemField && f.key !== locField),
    ...cat.fields.filter(f => f.key !== locField && f.key !== itemField),
  ];
}
function rowHtml(cat, row, idx) {
  const pinnedCells = displayFieldOrder(cat).slice(0, 2).map(f => `<td${f.key === cat.itemField ? ' class="col-item"' : f.key === cat.locationField ? ' class="col-loc"' : ''}>${fieldControlHTML(f, row[f.key], `data-row="${idx}"`)}</td>`).join('');
  const restCells = displayFieldOrder(cat).slice(2).map(f => `<td>${fieldControlHTML(f, row[f.key], `data-row="${idx}"`)}</td>`).join('');
  // 操作 (delete button) and # (row number) sit right after 地點/測項 (the pinned
  // columns), before the rest of the normally-scrolling fields — reducing the
  // pinned/sticky column chain from 5 down to 3 (勾選, 地點, 測項). Fewer adjacent
  // sticky columns means less surface area for the browser's sticky-repaint
  // compositing quirk that was making the checkbox visually vanish while scrolled
  // (confirmed real, not a positioning bug — the checkbox stayed fully clickable
  // underneath). 操作/# were never reported as having this problem themselves, so
  // they're the columns it's safest to drop from the pinned set — still visible
  // right after the pinned ones without needing to scroll all the way right.
  return `<tr data-row="${idx}"><td class="col-check">${checkboxHTML(` class="row-check" data-row="${idx}"`)}</td>${pinnedCells}<td class="col-actions"><button class="row-del-btn" data-row="${idx}" title="刪除此列">🗑</button></td><td class="col-num">${idx + 1}</td>${restCells}</tr>`;
}

function fieldControlHTML(field, value, rowAttr) {
  value = value ?? '';
  const base = `${rowAttr} data-field="${field.key}"`;
  switch (field.type) {
    case 'select': {
      const opts = field.options.map(o => {
        const label = (field.optionLabels && field.optionLabels[o]) || o || '（未選擇）';
        return `<option value="${escapeAttr(o)}" ${o === value ? 'selected' : ''}>${escapeHtml(label)}</option>`;
      }).join('');
      const titleAttr = field.help ? ` title="${escapeAttr(field.help)}"` : '';
      return `<select ${base}${titleAttr}>${opts}</select>`;
    }
    case 'date':
      // Plain text rather than a native <input type=date>: native date pickers render
      // according to browser/OS locale and can't be forced to show "YYYY/MM/DD" —
      // typing "2026/5/12" or "2026-5-12" both work, normalized on blur.
      return `<input type="text" ${base} value="${escapeAttr(toDateDisplayValue(value))}" class="date-input" placeholder="YYYY/MM/DD" inputmode="numeric" maxlength="10">`;
    case 'time':
      // Plain text rather than a native <input type=time>: native time pickers on many
      // devices show a scroll-wheel that's fiddly to land on an exact second, and on
      // some mobile browsers don't reliably fire change events at all. Typing "1430",
      // "14:30", or "14:30:00" all work — normalized to HH:MM on blur (the official
      // template's own time format has no seconds either).
      return `<input type="text" ${base} value="${escapeAttr(toTimeDisplayValue(value))}" class="time-input" placeholder="HH:MM" inputmode="numeric" maxlength="8">`;
    case 'suggest': {
      const listId = `suggest-${field.key.replace(/[^a-zA-Z0-9]/g, '')}`;
      if (!document.getElementById(listId)) {
        const dl = document.createElement('datalist');
        dl.id = listId;
        dl.innerHTML = field.options.map(o => `<option value="${escapeAttr(o)}">`).join('');
        document.body.appendChild(dl);
      }
      return `<input type="text" ${base} value="${escapeAttr(value)}" list="${listId}">`;
    }
    case 'unitcode':
      return `<input type="text" ${base} value="${escapeAttr(value)}" class="code-input" data-codetype="unit" title="${escapeAttr(lookupUnit(value))}" placeholder="代碼">`;
    case 'agencycode':
      return `<input type="text" ${base} value="${escapeAttr(value)}" class="code-input" data-codetype="agency" title="${escapeAttr(lookupAgency(value))}" placeholder="代碼">`;
    default:
      return `<input type="text" ${base} value="${escapeAttr(value)}">`;
  }
}

function wireGridEvents(project, catKey, cat) {
  const tbody = document.getElementById('gridBody');
  const COORD_FIELDS = ['座標系統', '採樣座標-經度 X', '採樣座標-緯度 Y'];

  const commit = (rowIdx, fieldKey, value) => {
    const rows = DataStore.getData(project.id, catKey);
    if (!rows[rowIdx]) return;
    rows[rowIdx][fieldKey] = value;
    // 比較關係 is derived from whatever's actually in 檢測數值/監測數值, not an
    // independently-set attribute — keep it in sync automatically whenever the
    // value field itself changes, so editing the value never leaves a stale
    // comparison symbol behind (e.g. a row inherited "ND" from history/import but
    // now has a real number typed in). Silent, not a sync-prompt: this is
    // maintaining internal consistency WITHIN one row, not propagating a change to
    // OTHER rows, so there's nothing to ask permission for.
    if ((fieldKey === '檢測數值' || fieldKey === '監測數值') && '比較關係' in rows[rowIdx]) {
      const derived = deriveComparisonRelation(value);
      if (derived) {
        rows[rowIdx]['比較關係'] = derived.cmp;
        if (derived.val !== value) rows[rowIdx][fieldKey] = derived.val;
        // Typing "NA"/"未檢測" into the value field means "未檢測" belongs in 備註,
        // not in 比較關係/檢測數值 — same rule the import parsers use. Only touch
        // 備註 when there's actually a note to record or a stale one to clear (so
        // this never clobbers unrelated text the person already wrote in 備註).
        if ('備註' in rows[rowIdx]) {
          const existingNote = rows[rowIdx]['備註'] || '';
          if (derived.note && !existingNote.includes(derived.note)) {
            rows[rowIdx]['備註'] = existingNote ? `${existingNote}；${derived.note}` : derived.note;
          } else if (!derived.note && existingNote === '未檢測') {
            // the value was corrected away from NA — drop the now-stale auto-note,
            // but leave any note the person wrote themselves untouched.
            rows[rowIdx]['備註'] = '';
          }
        }
      }
    }
    DataStore.saveData(project.id, catKey, rows);
    // Keep the site-item history snapshot current with manual corrections too — not
    // just at import time — so a value the person fixes by hand (e.g. filling in
    // coordinates a water report never provides) is what gets carried forward next
    // season, not whatever was frozen in at the moment of import.
    learnSiteItemHistory(project.id, catKey, cat, [rows[rowIdx]]);
  };

  // Whether row `r` belongs to the same "sampling event" as `source` for sync
  // purposes: same import batch, same location, and — unless the field actually
  // being synced IS the category itself — the same 檢測類別. A physical site can
  // carry BOTH a 環境噪音/道路交通噪音 sub-report and a separate 振動 sub-report
  // on the very same day; without the category check, correcting one's coordinate
  // could silently blank out the other's Y value the moment only X had been typed
  // in so far — this actually happened with real data, not just a theoretical risk.
  const matchesSyncGroup = (source, r, { requireCategory = true, requireDate = true } = {}) => {
    const locField = cat.locationField;
    if (r._batchId !== source._batchId) return false;
    if (r[locField] !== source[locField]) return false;
    if (requireCategory && r['檢測類別'] !== source['檢測類別']) return false;
    // Both ends of the sampling window have to match, not just 日期(起) — a site
    // sampled on the same start date but over a different span (a one-off grab
    // sample vs a 24-hour composite) is a different sampling event and must not be
    // swept up by a sync.
    if (requireDate && r['日期(起)'] !== source['日期(起)']) return false;
    if (requireDate && r['日期(迄)'] !== source['日期(迄)']) return false;
    return true;
  };

  // If the person corrects a coordinate on one row, offer to sync all three
  // coordinate fields to every other row that (a) came from the same import batch,
  // (b) shares the same sampling location, (c) shares the same 檢測類別, AND
  // (d) shares the same sampling date — e.g. a site sampled in both April and May
  // can genuinely have slightly different coordinates between visits, so syncing
  // must never cross dates. Always asks first, same reasoning as the date/time sync
  // below.
  const offerCoordSync = (rowIdx) => {
    const rows = DataStore.getData(project.id, catKey);
    const source = rows[rowIdx];
    const locField = cat.locationField;
    if (!source || !source._batchId || !source[locField] || !source['日期(起)']) return false;
    const matches = rows
      .map((r, idx) => ({ r, idx }))
      .filter(({ r, idx }) => idx !== rowIdx && matchesSyncGroup(source, r));
    if (matches.length === 0) return false;
    const anyDiff = matches.some(({ r }) => COORD_FIELDS.some(f => r[f] !== source[f]));
    if (!anyDiff) return false;
    const ok = confirm(
      `偵測到同一份檔案、同一天（${toDateDisplayValue(source['日期(起)'])}）、同一個測站「${source[locField]}」${source['檢測類別'] ? `、同為「${source['檢測類別']}」` : ''}還有 ${matches.length} 筆其他資料。\n` +
      `是否要將這些資料的座標一併同步更新為與這一筆相同？\n\n` +
      `（選擇「取消」則只修改目前這一筆，其他資料維持原狀。不同採樣日期或不同檢測類別的資料不會被同步。）`
    );
    if (!ok) return false;
    matches.forEach(({ r }) => { COORD_FIELDS.forEach(f => { r[f] = source[f]; }); });
    DataStore.saveData(project.id, catKey, rows);
    learnSiteItemHistory(project.id, catKey, cat, matches.map(m => m.r).concat([source]));
    return true;
  };

  // If the person corrects a date/time on one row, offer to sync all four date/time
  // fields to every other row that (a) came from the same import batch, (b) shares
  // the same sampling location, and (c) shares the same 檢測類別 — e.g. correcting
  // one test item's date in a multi-item water report should usually update the
  // whole site's rows from that report, since they're really one sampling event.
  // This always asks first rather than silently overwriting, both to avoid
  // surprising edits and because some browsers/devices don't reliably fire events
  // for native time pickers, which made a silent version of this hard to trust.
  const DATE_TIME_FIELDS = ['日期(起)', '時間(起)', '日期(迄)', '時間(迄)'];
  const offerDateTimeSync = (rowIdx) => {
    const rows = DataStore.getData(project.id, catKey);
    const source = rows[rowIdx];
    const locField = cat.locationField;
    if (!source || !source._batchId || !source[locField]) return false;
    const matches = rows
      .map((r, idx) => ({ r, idx }))
      .filter(({ r, idx }) => idx !== rowIdx && matchesSyncGroup(source, r, { requireDate: false }));
    if (matches.length === 0) return false;
    const anyDiff = matches.some(({ r }) => DATE_TIME_FIELDS.some(f => r[f] !== source[f]));
    if (!anyDiff) return false;
    const ok = confirm(
      `偵測到同一份檔案、同一個測站「${source[locField]}」${source['檢測類別'] ? `、同為「${source['檢測類別']}」` : ''}還有 ${matches.length} 筆其他資料。\n` +
      `是否要將這些資料的採樣日期／時間一併同步更新為與這一筆相同？\n\n` +
      `（選擇「取消」則只修改目前這一筆，其他資料維持原狀。不同檢測類別的資料不會被同步。）`
    );
    if (!ok) return false;
    matches.forEach(({ r }) => { DATE_TIME_FIELDS.forEach(f => { r[f] = source[f]; }); });
    DataStore.saveData(project.id, catKey, rows);
    learnSiteItemHistory(project.id, catKey, cat, matches.map(m => m.r).concat([source]));
    return true;
  };

  // 檢測類別 sync follows the same rule as coordinates (per the person's own
  // clarification): same batch (file) + same site + same sampling date. Unlike
  // date/time sync, this must NOT cross dates — a site's category classification for
  // one visit shouldn't silently overwrite a different visit's classification. This
  // is the one sync that can't require "same category" as a matching criterion,
  // since 檢測類別 is the very field being synced.
  const offerCategorySync = (rowIdx) => {
    const rows = DataStore.getData(project.id, catKey);
    const source = rows[rowIdx];
    const locField = cat.locationField;
    if (!source || !source._batchId || !source[locField] || !source['日期(起)']) return false;
    const matches = rows
      .map((r, idx) => ({ r, idx }))
      .filter(({ r, idx }) => idx !== rowIdx && matchesSyncGroup(source, r, { requireCategory: false }));
    if (matches.length === 0) return false;
    const anyDiff = matches.some(({ r }) => r['檢測類別'] !== source['檢測類別']);
    if (!anyDiff) return false;
    const ok = confirm(
      `偵測到同一份檔案、同一天（${toDateDisplayValue(source['日期(起)'])}）、同一個測站「${source[locField]}」還有 ${matches.length} 筆其他資料。\n` +
      `是否要將這些資料的檢測類別一併同步更新為「${source['檢測類別']}」？\n\n` +
      `（選擇「取消」則只修改目前這一筆，其他資料維持原狀。）`
    );
    if (!ok) return false;
    matches.forEach(({ r }) => { r['檢測類別'] = source['檢測類別']; });
    DataStore.saveData(project.id, catKey, rows);
    learnSiteItemHistory(project.id, catKey, cat, matches.map(m => m.r).concat([source]));
    return true;
  };

  // Generic sync for every OTHER field (管制標準、管制區、檢測方法、單位代碼、檢測
  // 機構、備註等等) — same batch/site/category/date rule as the coordinate sync
  // above. Deliberately excludes: the item field itself (each row IS a different
  // item, syncing it would be nonsensical), the location field (editing it would
  // break the very "same location" matching this relies on — see openCoordModal's
  // separate, date-aware handling for correcting locations), the measurement value
  // fields (always unique per row by definition), and anything already covered by
  // a dedicated sync above (coordinates / date-time / 檢測類別) to avoid asking twice.
  const SYNC_EXCLUDED_FIELDS = new Set([
    cat.itemField, cat.locationField, '檢測數值', '監測數值',
    ...COORD_FIELDS, ...DATE_TIME_FIELDS, '檢測類別',
  ]);
  const offerGenericFieldSync = (rowIdx, fieldKey) => {
    const rows = DataStore.getData(project.id, catKey);
    const source = rows[rowIdx];
    const locField = cat.locationField;
    if (!source || !source._batchId || !source[locField] || !source['日期(起)']) return false;
    const matches = rows
      .map((r, idx) => ({ r, idx }))
      .filter(({ r, idx }) => idx !== rowIdx && matchesSyncGroup(source, r));
    if (matches.length === 0) return false;
    const anyDiff = matches.some(({ r }) => r[fieldKey] !== source[fieldKey]);
    if (!anyDiff) return false;
    const fieldLabel = (cat.fields.find(f => f.key === fieldKey) || {}).label || fieldKey;
    const ok = confirm(
      `偵測到同一份檔案、同一天（${toDateDisplayValue(source['日期(起)'])}）、同一個測站「${source[locField]}」${source['檢測類別'] ? `、同為「${source['檢測類別']}」` : ''}還有 ${matches.length} 筆其他資料。\n` +
      `是否要將這些資料的「${fieldLabel}」一併同步更新為「${source[fieldKey] || '（空白）'}」？\n\n` +
      `（選擇「取消」則只修改目前這一筆，其他資料維持原狀。）`
    );
    if (!ok) return false;
    matches.forEach(({ r }) => { r[fieldKey] = source[fieldKey]; });
    DataStore.saveData(project.id, catKey, rows);
    learnSiteItemHistory(project.id, catKey, cat, matches.map(m => m.r).concat([source]));
    return true;
  };

  tbody.addEventListener('input', (e) => {
    const t = e.target;
    if (!t.dataset.field) return;
    if (t.classList.contains('code-input')) {
      t.title = t.dataset.codetype === 'unit' ? lookupUnit(t.value) : lookupAgency(t.value);
    }
    if (t.tagName === 'SELECT') return; // handled by change
    commit(Number(t.dataset.row), t.dataset.field, t.value);
  });
  tbody.addEventListener('change', (e) => {
    const t = e.target;
    if (!t.dataset.field) return;
    if (t.tagName === 'SELECT') {
      commit(Number(t.dataset.row), t.dataset.field, t.value);
      const fieldKey = t.dataset.field;
      const rowIdx = Number(t.dataset.row);
      if (COORD_FIELDS.includes(fieldKey) && offerCoordSync(rowIdx)) { renderContentPreservingScroll(); return; }
      if (fieldKey === '檢測類別' && offerCategorySync(rowIdx)) { renderContentPreservingScroll(); return; }
      if (!SYNC_EXCLUDED_FIELDS.has(fieldKey) && offerGenericFieldSync(rowIdx, fieldKey)) renderContentPreservingScroll();
    }
  });
  // use focusout (bubbles) rather than blur to catch this via delegation; only
  // re-render on blur (not every keystroke) so typing isn't interrupted.
  tbody.addEventListener('focusout', (e) => {
    const t = e.target;
    if (!t.dataset.field || t.tagName === 'SELECT') return;
    const rowIdx = Number(t.dataset.row);
    const fieldKey = t.dataset.field;

    if (t.classList.contains('time-input')) {
      const normalized = normalizeTimeString(t.value); // full HH:MM:SS for storage
      t.value = normalized.slice(0, 5); // HH:MM for display
      commit(rowIdx, fieldKey, normalized);
    }
    if (t.classList.contains('date-input')) {
      const normalized = normalizeDateString(t.value); // ISO YYYY-MM-DD for storage
      t.value = toDateDisplayValue(normalized) || normalized; // YYYY/MM/DD for display
      commit(rowIdx, fieldKey, normalized);
    }

    if (COORD_FIELDS.includes(fieldKey) && offerCoordSync(rowIdx)) { renderContentPreservingScroll(); return; }
    if (DATE_TIME_FIELDS.includes(fieldKey) && offerDateTimeSync(rowIdx)) { renderContentPreservingScroll(); return; }
    if (!SYNC_EXCLUDED_FIELDS.has(fieldKey) && offerGenericFieldSync(rowIdx, fieldKey)) renderContentPreservingScroll();
  });
  tbody.addEventListener('click', (e) => {
    const btn = e.target.closest('.row-del-btn');
    if (!btn) return;
    if (!confirm('確定要刪除這一列資料嗎？（可用「↩️ 復原上一步」救回）')) return;
    pushUndoSnapshot(project.id, catKey, '刪除 1 筆資料');
    const rows = DataStore.getData(project.id, catKey);
    rows.splice(Number(btn.dataset.row), 1);
    DataStore.saveData(project.id, catKey, rows);
    renderContent();
  });
  // Pressing Enter should confirm the edit immediately (normalize, commit, and offer
  // any sync) rather than requiring the person to click elsewhere first — blur()
  // reuses the exact same focusout logic above rather than duplicating it.
  tbody.addEventListener('keydown', (e) => {
    if (e.key !== 'Enter') return;
    const t = e.target;
    if (!t.dataset.field || t.tagName === 'SELECT') return;
    e.preventDefault();
    t.blur();
  });
}

// ---------- item-selection presets ("記憶標籤") for 常用測項新增 ----------
// Stored per-category, shared across ALL projects in this browser (not tied to one
// project) — the person's "typical combo of 10 items for 3 stations" recurs across
// different projects, so a global-per-browser scope is more useful here than a
// per-project one. Stores { itemName, method } pairs (not variant array indices),
// so a preset still applies correctly even if the shared catalog's variant order
// later changes — matching is done by method text at apply time, skipping any item
// that's no longer findable rather than silently misapplying the wrong method.
const ITEM_PRESET_LIMIT = 8;
function presetStorageKey(catKey) { return `envapp_itemPresets_${catKey}`; }
function getItemPresets(catKey) {
  try { return JSON.parse(localStorage.getItem(presetStorageKey(catKey))) || []; }
  catch (e) { return []; }
}
function saveItemPresets(catKey, presets) {
  localStorage.setItem(presetStorageKey(catKey), JSON.stringify(presets.slice(0, ITEM_PRESET_LIMIT)));
}
function rememberItemPreset(catKey, checkedItems) {
  // checkedItems: [{ itemName, method }]
  if (checkedItems.length === 0) return;
  const presets = getItemPresets(catKey);
  const signature = checkedItems.map(i => `${i.itemName}::${i.method}`).sort().join('|');
  const withoutDupe = presets.filter(p => p.items.map(i => `${i.itemName}::${i.method}`).sort().join('|') !== signature);
  const names = checkedItems.map(i => i.itemName);
  const label = names.length <= 3 ? names.join('、') : `${names.slice(0, 3).join('、')}…等${names.length}項`;
  withoutDupe.unshift({ label, items: checkedItems });
  saveItemPresets(catKey, withoutDupe);
}
function renderItemPresets(catKey) {
  const wrap = document.getElementById('commonItemsPresets');
  const presets = getItemPresets(catKey);
  if (presets.length === 0) { wrap.innerHTML = ''; return; }
  wrap.innerHTML = `
    <div class="hint" style="width:100%;margin-bottom:4px">🏷️ 最近使用的組合（點擊可快速套用同樣的勾選）：</div>
    ${presets.map((p, i) => `
      <span class="item-preset-chip" data-idx="${i}">
        <span class="item-preset-label">${escapeHtml(p.label)}</span>
        <button type="button" class="item-preset-remove" data-idx="${i}" title="刪除這個組合">✕</button>
      </span>
    `).join('')}
  `;
  wrap.querySelectorAll('.item-preset-label').forEach(el => {
    el.addEventListener('click', () => applyItemPreset(catKey, Number(el.closest('.item-preset-chip').dataset.idx)));
  });
  wrap.querySelectorAll('.item-preset-remove').forEach(btn => {
    btn.addEventListener('click', (e) => {
      e.stopPropagation();
      const idx = Number(btn.dataset.idx);
      const presets2 = getItemPresets(catKey);
      presets2.splice(idx, 1);
      saveItemPresets(catKey, presets2);
      renderItemPresets(catKey);
    });
  });
}
function applyItemPreset(catKey, presetIdx) {
  const presets = getItemPresets(catKey);
  const preset = presets[presetIdx];
  if (!preset) return;
  // Clear any active search/checked-only filter first, so every row the preset is
  // about to check is actually visible afterward — otherwise a stale filter could
  // hide the very rows this just checked, making it look like nothing happened.
  document.getElementById('commonItemsSearch').value = '';
  document.getElementById('commonItemsShowCheckedOnly').checked = false;
  document.querySelectorAll('.common-item-row').forEach(row => row.classList.remove('row-hidden'));
  document.querySelectorAll('.common-item-check').forEach(cb => { cb.checked = false; });
  let skipped = 0;
  preset.items.forEach(({ itemName, method }) => {
    const entryIdx = state.commonItemsEntries.findIndex(e => e.itemName === itemName);
    if (entryIdx === -1) { skipped++; return; }
    const entry = state.commonItemsEntries[entryIdx];
    const cb = document.querySelector(`.common-item-check[data-idx="${entryIdx}"]`);
    if (!cb) { skipped++; return; }
    cb.checked = true;
    const variantIdx = entry.variants.findIndex(v => v.method === method);
    if (variantIdx !== -1) {
      const sel = document.querySelector(`.common-item-method-select[data-idx="${entryIdx}"]`);
      if (sel) sel.value = String(variantIdx);
    }
  });
  document.getElementById('commonItemsCheckedCount').textContent = `已勾選 ${preset.items.length - skipped} 項`;
  if (skipped > 0) alert(`有 ${skipped} 個項目已經不在目前的測項清單裡，已略過（可能是清單被更新過）。`);
}

/** Wires drag-and-drop onto a dropzone container that already contains a
 *  <input type="file">, so a person can drop files directly instead of clicking to
 *  open the OS file picker — same end result, just a faster path for the same
 *  three import entry points (regular import, batch import). `onFiles` receives
 *  the dropped FileList exactly as `<input>.files`/`.change` would. */
function wireDropzone(dropzoneId, onFiles) {
  const zone = document.getElementById(dropzoneId);
  if (!zone) return;
  let dragDepth = 0; // dragenter/dragleave fire for every child element too; only
                      // toggle the active style when depth returns to 0
  zone.addEventListener('dragenter', (e) => {
    e.preventDefault();
    dragDepth++;
    zone.classList.add('dropzone-active');
  });
  zone.addEventListener('dragover', (e) => { e.preventDefault(); });
  zone.addEventListener('dragleave', (e) => {
    e.preventDefault();
    dragDepth = Math.max(0, dragDepth - 1);
    if (dragDepth === 0) zone.classList.remove('dropzone-active');
  });
  zone.addEventListener('drop', (e) => {
    e.preventDefault();
    dragDepth = 0;
    zone.classList.remove('dropzone-active');
    const files = e.dataTransfer && e.dataTransfer.files;
    if (files && files.length > 0) onFiles(files);
  });
}

/**
 * Lets files ACCUMULATE across multiple separate drag-and-drop actions (or
 * multiple file-picker selections) before actually starting the import — rather
 * than each drop immediately kicking off processing. This matters because the OS
 * file picker/drag-and-drop can only select files from ONE folder view at a time;
 * someone with files scattered across several folders had no way to combine them
 * into a single import short of dragging one at a time and re-running the import
 * flow for each — confirmed as a real workflow pain point. Each new drop/selection
 * is MERGED into the running list (by name+size, so re-dropping the same file
 * twice by accident doesn't duplicate it), individually removable, with an
 * explicit confirm button that only appears once at least one file is staged.
 * Returns { reset() } so the caller can clear the staged list when its modal closes.
 */
function wireFileStaging({ dropzoneId, fileInputId, listId, confirmBtnId, onConfirm }) {
  const fileInput = document.getElementById(fileInputId);
  const listEl = document.getElementById(listId);
  const confirmBtn = document.getElementById(confirmBtnId);
  let staged = [];

  const renderList = () => {
    if (staged.length === 0) {
      listEl.innerHTML = '';
      confirmBtn.classList.add('hidden');
      return;
    }
    listEl.innerHTML = staged.map((f, i) => `
      <div class="staged-file-item">
        <span class="staged-file-name">📄 ${escapeHtml(f.name)}</span>
        <button type="button" class="staged-file-remove" data-idx="${i}" title="移除這個檔案">✕</button>
      </div>
    `).join('');
    listEl.querySelectorAll('.staged-file-remove').forEach(btn => {
      btn.addEventListener('click', () => {
        staged.splice(Number(btn.dataset.idx), 1);
        renderList();
      });
    });
    confirmBtn.classList.remove('hidden');
    confirmBtn.textContent = `✅ 開始判讀這 ${staged.length} 個檔案`;
  };

  const addFiles = (fileList) => {
    [...fileList].forEach(f => {
      if (!staged.some(s => s.name === f.name && s.size === f.size)) staged.push(f);
    });
    renderList();
  };

  fileInput.addEventListener('change', (e) => {
    addFiles(e.target.files);
    fileInput.value = ''; // reset so picking the same file again still fires 'change'
  });
  wireDropzone(dropzoneId, addFiles);
  confirmBtn.addEventListener('click', () => {
    const filesToImport = staged;
    staged = [];
    renderList();
    onConfirm(filesToImport);
  });

  return { reset: () => { staged = []; renderList(); } };
}

function addEmptyRow(project, catKey) {
  const cat = CATEGORIES[catKey];
  const rows = DataStore.getData(project.id, catKey);
  const blank = {};
  cat.fields.forEach(f => { blank[f.key] = ''; });
  rows.push(blank);
  DataStore.saveData(project.id, catKey, rows);
  renderContent();
  // scroll to bottom of table
  const wrap = document.querySelector('.table-wrap');
  if (wrap) wrap.scrollTop = wrap.scrollHeight;
}

/**
 * Lets the person pick from a checklist of commonly-used items for this category —
 * built-in items come from 測項對照表.txt (shared, site-wide — see loadSiteConfig),
 * merged with any items THIS project has already used before (from its own item
 * memory) that aren't already in the built-in list, so a project's own history
 * naturally grows the list over time without needing the built-in catalog touched.
 * An item with more than one known method (e.g. 溶氧 measured by either 電極法 or
 * 碘定量法, which can even have different units) shows a method dropdown so the
 * person picks which one — the unit that gets filled in follows whichever method
 * they chose, not a fixed default. A person can also type in a brand-new item this
 * session that isn't in either list yet; that gets remembered into this project's
 * own item memory (not the shared catalog) for next time.
 */
function openCommonItemsModal(project, catKey) {
  const cat = CATEGORIES[catKey];
  const builtIn = siteConfig.itemCatalog[catKey] || [];
  const itemMemory = DataStore.getItemMemory(project.id, catKey);
  const builtInNames = new Set(builtIn.map(e => e.itemName));
  const projectOwnItems = Object.entries(itemMemory)
    .filter(([name]) => name && !builtInNames.has(name))
    .map(([name, mem]) => ({ itemName: name, variants: [{ method: mem.method || '', unitCode: mem.unitCode || '' }] }));

  state.commonItemsCatKey = catKey;
  state.commonItemsEntries = [...builtIn, ...projectOwnItems];
  state.commonItemsColumnFilters = {};
  renderCommonItemsList(cat);
  renderItemPresets(catKey);
  document.getElementById('commonItemsQuantity').value = '1';
  document.getElementById('commonItemsFilterName').classList.remove('col-filter-active');
  document.getElementById('commonItemsFilterMethod').classList.remove('col-filter-active');
  document.getElementById('commonItemsFilterName').onclick = (e) => openCommonItemsColumnFilterPopup('name', e.currentTarget);
  document.getElementById('commonItemsFilterMethod').onclick = (e) => openCommonItemsColumnFilterPopup('method', e.currentTarget);
  document.getElementById('commonItemsModal').classList.remove('hidden');
}
function renderCommonItemsList(cat) {
  const entries = state.commonItemsEntries;
  state.commonItemsSort = null;
  state.commonItemsDisplayOrder = entries.map((_, i) => i);
  renderCommonItemsTableBody(cat);
  wireCommonItemsSortHeaders(cat);
}
/** Redraws just the table BODY in whatever order state.commonItemsDisplayOrder
 *  currently specifies — used both for the initial render and after re-sorting.
 *  `data-idx` always refers to the position in state.commonItemsEntries (the
 *  underlying data), never the on-screen row position, so re-sorting never
 *  scrambles which entry a checkbox/method-select belongs to. Preserves whatever
 *  was already checked/selected before this redraw (sorting shouldn't lose the
 *  person's in-progress selections). */
function renderCommonItemsTableBody(cat) {
  const tbody = document.getElementById('commonItemsTableBody');
  const entries = state.commonItemsEntries;
  const order = state.commonItemsDisplayOrder;
  if (entries.length === 0) {
    tbody.innerHTML = '<tr><td colspan="3" class="hint" style="padding:16px;text-align:center">目前沒有內建測項，也還沒有這個計畫自己用過的測項紀錄。可以用下方「新增自訂測項」開始建立。</td></tr>';
    return;
  }
  const checkedIdxSet = new Set([...document.querySelectorAll('.common-item-check:checked')].map(cb => Number(cb.dataset.idx)));
  const selectedVariantByIdx = {};
  document.querySelectorAll('.common-item-method-select').forEach(sel => { selectedVariantByIdx[Number(sel.dataset.idx)] = sel.value; });

  tbody.innerHTML = order.map(i => {
    const entry = entries[i];
    const hasMultiple = entry.variants.length > 1;
    const searchText = [entry.itemName, ...entry.variants.map(v => v.method)].filter(Boolean).join(' ').toLowerCase();
    const isChecked = checkedIdxSet.has(i);
    const selectedVariantIdx = selectedVariantByIdx[i] !== undefined ? selectedVariantByIdx[i] : '0';
    const methodCell = hasMultiple
      ? `<select class="common-item-method-select" data-idx="${i}">
          ${entry.variants.map((v, vi) => `<option value="${vi}" ${String(vi) === selectedVariantIdx ? 'selected' : ''}>${escapeHtml(v.method || '（無方法資訊）')}${v.unitCode ? `（${escapeHtml(UNIT_CODES[v.unitCode] || v.unitCode)}）` : ''}</option>`).join('')}
        </select>`
      : escapeHtml(entry.variants[0] ? [entry.variants[0].method, entry.variants[0].unitCode ? (UNIT_CODES[entry.variants[0].unitCode] || entry.variants[0].unitCode) : ''].filter(Boolean).join('，') : '');
    return `
      <tr class="common-item-row" data-idx="${i}" data-search-text="${escapeAttr(searchText)}" data-name-value="${escapeAttr(entry.itemName)}" data-method-value="${escapeAttr(entry.variants[0]?.method || '')}">
        <td class="ci-col-check"><input type="checkbox" class="common-item-check" data-idx="${i}" ${isChecked ? 'checked' : ''}></td>
        <td class="common-item-name">${escapeHtml(entry.itemName)}</td>
        <td class="common-item-method">${methodCell}</td>
      </tr>
    `;
  }).join('');
  wireCommonItemsListFilters();
}
/** Click-to-sort on the table headers, Excel-style: cycles asc → desc → original
 *  order. Bound once per modal-open (the <th> elements themselves are static HTML,
 *  not regenerated by renderCommonItemsTableBody), using onclick assignment so
 *  reopening the modal never double-binds a second handler. */
function wireCommonItemsSortHeaders(cat) {
  document.querySelectorAll('.ci-sortable').forEach(th => {
    th.onclick = () => {
      const sortKey = th.dataset.sort;
      const current = state.commonItemsSort;
      let direction = 'asc';
      if (current && current.key === sortKey) {
        direction = current.direction === 'asc' ? 'desc' : (current.direction === 'desc' ? null : 'asc');
      }
      const entries = state.commonItemsEntries;
      if (!direction) {
        state.commonItemsDisplayOrder = entries.map((_, i) => i);
        state.commonItemsSort = null;
      } else {
        const getKey = (i) => sortKey === 'name' ? entries[i].itemName : (entries[i].variants[0]?.method || '');
        state.commonItemsDisplayOrder = entries.map((_, i) => i).sort((a, b) => {
          const cmp = String(getKey(a)).localeCompare(String(getKey(b)), 'zh-Hant');
          return direction === 'asc' ? cmp : -cmp;
        });
        state.commonItemsSort = { key: sortKey, direction };
      }
      document.querySelectorAll('.ci-sortable .sort-indicator').forEach(el => { el.textContent = '⇅'; el.classList.remove('sort-active'); });
      if (direction) {
        const indicator = th.querySelector('.sort-indicator');
        indicator.textContent = direction === 'asc' ? '▲' : '▼';
        indicator.classList.add('sort-active');
      }
      renderCommonItemsTableBody(cat);
    };
  });
}
/** Excel-style column filter for the common-items table — a checkbox list of
 *  distinct values (not just sort) for 測項名稱 or 方法／單位, reusing the same
 *  #colFilterPopup element and interaction pattern as the main data grid's column
 *  filter and the import row-detail table's, for a consistent feel. Filter state
 *  lives in state.commonItemsColumnFilters and is combined with search/checked-only
 *  in applyFilters below — a row must pass every active filter simultaneously. */
function openCommonItemsColumnFilterPopup(filterKey, btnEl) {
  const popup = document.getElementById('colFilterPopup');
  const entries = state.commonItemsEntries;
  const uniqueValues = filterKey === 'name'
    ? [...new Set(entries.map(e => e.itemName))].sort((a, b) => a.localeCompare(b, 'zh-Hant'))
    : [...new Set(entries.map(e => e.variants[0]?.method || ''))].sort((a, b) => a.localeCompare(b, 'zh-Hant'));
  if (!state.commonItemsColumnFilters) state.commonItemsColumnFilters = {};
  const currentFilter = state.commonItemsColumnFilters[filterKey];

  popup.innerHTML = `
    <div class="col-filter-search"><input type="text" id="colFilterSearchInput" placeholder="搜尋選項..."></div>
    <div class="col-filter-actions">
      <button type="button" id="colFilterSelectAll" class="btn btn-ghost btn-sm">全選</button>
      <button type="button" id="colFilterClearAll" class="btn btn-ghost btn-sm">清除</button>
    </div>
    <div class="col-filter-list" id="colFilterList">
      ${uniqueValues.map(v => `
        <label class="col-filter-item">
          <input type="checkbox" value="${escapeAttr(v)}" ${!currentFilter || currentFilter.has(v) ? 'checked' : ''}>
          <span>${escapeHtml(v === '' ? '（空白）' : v)}</span>
        </label>
      `).join('')}
    </div>
    <div class="col-filter-buttons">
      <button type="button" id="colFilterCancel" class="btn btn-ghost btn-sm">取消</button>
      <button type="button" id="colFilterApply" class="btn btn-primary btn-sm">確定</button>
    </div>
  `;

  const rect = btnEl.getBoundingClientRect();
  const popupWidth = 240;
  popup.style.left = Math.max(4, Math.min(rect.left, window.innerWidth - popupWidth - 4)) + 'px';
  popup.style.top = (rect.bottom + window.scrollY + 4) + 'px';
  popup.classList.remove('hidden');

  document.getElementById('colFilterSearchInput').addEventListener('input', (e) => {
    const q = e.target.value.trim().toLowerCase();
    document.querySelectorAll('#colFilterList .col-filter-item').forEach(label => {
      label.style.display = (!q || label.textContent.toLowerCase().includes(q)) ? '' : 'none';
    });
  });
  document.getElementById('colFilterSelectAll').addEventListener('click', () => {
    document.querySelectorAll('#colFilterList input[type="checkbox"]').forEach(cb => { cb.checked = true; });
  });
  document.getElementById('colFilterClearAll').addEventListener('click', () => {
    document.querySelectorAll('#colFilterList input[type="checkbox"]').forEach(cb => { cb.checked = false; });
  });
  document.getElementById('colFilterCancel').addEventListener('click', () => { popup.classList.add('hidden'); });
  document.getElementById('colFilterApply').addEventListener('click', () => {
    const checked = [...document.querySelectorAll('#colFilterList input[type="checkbox"]:checked')].map(cb => cb.value);
    if (checked.length === uniqueValues.length) {
      delete state.commonItemsColumnFilters[filterKey];
    } else {
      state.commonItemsColumnFilters[filterKey] = new Set(checked);
    }
    popup.classList.add('hidden');
    btnEl.classList.toggle('col-filter-active', !!state.commonItemsColumnFilters[filterKey]);
    document.getElementById('commonItemsSearch').dispatchEvent(new Event('input', { bubbles: true }));
  });
}
/** Wires the search box, "只顯示已勾選" toggle, "全選目前顯示的" checkbox, and
 *  live selected-count indicator. Uses DOM show/hide (not re-rendering) so
 *  checkbox/method-select state is never lost while filtering. Shows an explicit
 *  message when the filtered result is empty — including the specific case where
 *  "只顯示已勾選" is on but nothing is checked yet, which otherwise looks
 *  indistinguishable from "the search box is broken" (confirmed real confusion:
 *  the list going blank with zero explanation looked exactly like a bug). Re-run
 *  after any re-render of the table body. */
function wireCommonItemsListFilters() {
  const searchInput = document.getElementById('commonItemsSearch');
  const showCheckedOnly = document.getElementById('commonItemsShowCheckedOnly');
  const checkAllVisible = document.getElementById('commonItemsCheckAllVisible');
  const countEl = document.getElementById('commonItemsCheckedCount');
  const tbody = document.getElementById('commonItemsTableBody');
  const rows = () => [...tbody.querySelectorAll('.common-item-row')];

  const updateCount = () => {
    const n = document.querySelectorAll('.common-item-check:checked').length;
    countEl.textContent = `已勾選 ${n} 項`;
  };
  const applyFilters = () => {
    const q = searchInput.value.trim().toLowerCase();
    const colFilters = state.commonItemsColumnFilters || {};
    let anyVisible = false;
    rows().forEach(row => {
      const matchesSearch = !q || row.dataset.searchText.includes(q);
      const matchesChecked = !showCheckedOnly.checked || row.querySelector('.common-item-check').checked;
      const matchesName = !colFilters.name || colFilters.name.has(row.dataset.nameValue);
      const matchesMethod = !colFilters.method || colFilters.method.has(row.dataset.methodValue);
      const visible = matchesSearch && matchesChecked && matchesName && matchesMethod;
      row.classList.toggle('row-hidden', !visible);
      if (visible) anyVisible = true;
    });
    let emptyRow = tbody.querySelector('.common-items-empty-row');
    if (!anyVisible && rows().length > 0) {
      if (!emptyRow) {
        emptyRow = document.createElement('tr');
        emptyRow.className = 'common-items-empty-row';
        tbody.appendChild(emptyRow);
      }
      const noneCheckedYet = showCheckedOnly.checked && document.querySelectorAll('.common-item-check:checked').length === 0;
      const hasColFilter = colFilters.name || colFilters.method;
      const msg = noneCheckedYet
        ? '目前還沒有勾選任何測項。取消「只顯示已勾選」即可看到完整清單。'
        : hasColFilter
          ? '找不到符合目前篩選條件的測項，請調整搜尋或欄位篩選再試試。'
          : '找不到符合搜尋條件的測項，請換個關鍵字試試。';
      emptyRow.innerHTML = `<td colspan="3" class="hint" style="padding:16px;text-align:center">${msg}</td>`;
    } else if (emptyRow) {
      emptyRow.remove();
    }
  };
  searchInput.value = '';
  showCheckedOnly.checked = false;
  searchInput.oninput = applyFilters;
  showCheckedOnly.onchange = applyFilters;
  if (checkAllVisible) {
    checkAllVisible.checked = false;
    checkAllVisible.onchange = () => {
      rows().forEach(row => {
        if (!row.classList.contains('row-hidden')) row.querySelector('.common-item-check').checked = checkAllVisible.checked;
      });
      updateCount();
    };
  }
  rows().forEach(row => {
    row.querySelector('.common-item-check').addEventListener('change', updateCount);
  });
  updateCount();
  applyFilters();
}

// ---------- coordinate manager (bulk-fill site coordinates, e.g. for handwritten/scanned reports) ----------
function openCoordModal(project, catKey) {
  const cat = CATEGORIES[catKey];
  const rows = DataStore.getData(project.id, catKey);
  const locField = cat.locationField;

  // Group by (location, sampling date) — NOT location alone. The same named site can
  // legitimately have slightly different coordinates between visits (e.g. April vs
  // May), so applying one set of coordinates across all dates for a site would
  // silently corrupt the other visits' data.
  const groups = {}; // "loc\u0001date" -> { loc, date, indices: [...], coordSystem, x, y }
  rows.forEach((row, idx) => {
    const loc = (row[locField] || '').trim() || '（未命名測站）';
    const date = row['日期(起)'] || '（未填日期）';
    const key = loc + '\u0001' + date;
    if (!groups[key]) {
      groups[key] = {
        loc, date, indices: [],
        coordSystem: row['座標系統'] || '',
        x: row['採樣座標-經度 X'] || '',
        y: row['採樣座標-緯度 Y'] || '',
      };
    }
    groups[key].indices.push(idx);
    // prefer an already-filled value if this group doesn't have one yet
    if (!groups[key].x && row['採樣座標-經度 X']) groups[key].x = row['採樣座標-經度 X'];
    if (!groups[key].y && row['採樣座標-緯度 Y']) groups[key].y = row['採樣座標-緯度 Y'];
    if (!groups[key].coordSystem && row['座標系統']) groups[key].coordSystem = row['座標系統'];
  });

  const entries = Object.entries(groups);
  const wrap = document.getElementById('coordSitesWrap');
  if (entries.length === 0) {
    wrap.innerHTML = '<p class="hint" style="padding:14px">目前此類別尚無資料，請先新增或匯入資料後再使用測站座標管理。</p>';
  } else {
    wrap.innerHTML = `<table class="mapping-table">
      <thead><tr><th>測站名稱</th><th>採樣日期</th><th title="2：WGS84（全球座標，例如經度 120.681，緯度 24.147）／3：TWD97-TM2（投影座標系，例如 X=193150, Y=2670900）">座標系統 ℹ️</th><th>經度 X</th><th>緯度 Y</th><th>筆數</th></tr></thead>
      <tbody id="coordSitesBody">
        ${entries.map(([key, g]) => `<tr data-group-key="${escapeAttr(key)}">
          <td>${escapeHtml(g.loc)}</td>
          <td>${escapeHtml(g.date)}</td>
          <td><select data-coord-field="座標系統" title="2：WGS84（全球座標，例如經度 120.681，緯度 24.147）／3：TWD97-TM2（投影座標系，例如 X=193150, Y=2670900）">
            <option value="" ${g.coordSystem === '' ? 'selected' : ''}>（未選擇）</option>
            <option value="2" ${g.coordSystem === '2' ? 'selected' : ''}>2：WGS84（全球座標）</option>
            <option value="3" ${g.coordSystem === '3' ? 'selected' : ''}>3：TWD97-TM2（投影座標系）</option>
          </select></td>
          <td><input type="text" data-coord-field="採樣座標-經度 X" value="${escapeAttr(g.x)}" placeholder="例：120.681 或 193150"></td>
          <td><input type="text" data-coord-field="採樣座標-緯度 Y" value="${escapeAttr(g.y)}" placeholder="例：24.147 或 2670900"></td>
          <td>${g.indices.length}</td>
        </tr>`).join('')}
      </tbody>
    </table>`;
  }

  document.getElementById('coordModal').dataset.projectId = project.id;
  document.getElementById('coordModal').dataset.catKey = catKey;
  document.getElementById('btnCoordSave').classList.toggle('hidden', entries.length === 0);
  document.getElementById('coordModal').classList.remove('hidden');
}

function closeCoordModal() {
  document.getElementById('coordModal').classList.add('hidden');
}

function saveCoordModal() {
  const modal = document.getElementById('coordModal');
  const projectId = modal.dataset.projectId;
  const catKey = modal.dataset.catKey;
  const cat = CATEGORIES[catKey];
  const rows = DataStore.getData(projectId, catKey);
  const locField = cat.locationField;
  const touchedRows = [];

  // Coordinate manager can touch many rows at once, same as bulk delete/bulk edit —
  // it should get the same undo protection as every other multi-row change, not be
  // the one exception a person can't recover from if they mistype a coordinate.
  pushUndoSnapshot(projectId, catKey, `套用測站座標`);

  document.querySelectorAll('#coordSitesBody tr').forEach(tr => {
    const [loc, date] = tr.dataset.groupKey.split('\u0001');
    const values = {};
    tr.querySelectorAll('[data-coord-field]').forEach(el => { values[el.dataset.coordField] = el.value; });
    rows.forEach(row => {
      const rowLoc = (row[locField] || '').trim() || '（未命名測站）';
      const rowDate = row['日期(起)'] || '（未填日期）';
      if (rowLoc === loc && rowDate === date) { Object.assign(row, values); touchedRows.push(row); }
    });
  });

  DataStore.saveData(projectId, catKey, rows);
  // Coordinates are frequently the one thing a raw water/air lab report never
  // includes at all — filled in here, after import, rather than at confirm time.
  // Without re-learning history now, the site-item snapshot stays frozen at
  // whatever it captured at import (blank coordinates), so next season's auto-fill
  // would never have anything to carry forward no matter how many times this gets
  // filled in — re-learning on every save keeps the remembered snapshot current.
  if (touchedRows.length > 0) learnSiteItemHistory(projectId, catKey, cat, touchedRows);
  closeCoordModal();
  renderContent();
  alert('已套用座標到符合的資料列（僅限相同測站＋相同採樣日期的資料）。');
}

// ---------- method manager (test method + unit code, remembered across seasons) ----------
function openMethodModal(project, catKey) {
  const cat = CATEGORIES[catKey];
  const rows = DataStore.getData(project.id, catKey);
  const memory = DataStore.getItemMemory(project.id, catKey);
  const itemField = cat.itemField;

  const groups = {}; // itemName -> { indices, method, unitCode }
  rows.forEach((row, idx) => {
    const item = (row[itemField] || '').trim() || '（未命名項目）';
    if (!groups[item]) {
      groups[item] = {
        indices: [],
        method: row[cat.methodField] || (memory[item] && memory[item].method) || '',
        unitCode: cat.unitField ? (row[cat.unitField] || (memory[item] && memory[item].unitCode) || '') : null,
      };
    }
    groups[item].indices.push(idx);
    if (!groups[item].method && row[cat.methodField]) groups[item].method = row[cat.methodField];
    if (cat.unitField && !groups[item].unitCode && row[cat.unitField]) groups[item].unitCode = row[cat.unitField];
  });

  const entries = Object.entries(groups);
  const wrap = document.getElementById('methodItemsWrap');
  if (entries.length === 0) {
    wrap.innerHTML = '<p class="hint" style="padding:14px">目前此類別尚無資料，請先新增或匯入資料後再使用檢測方法管理。</p>';
  } else {
    wrap.innerHTML = `<table class="mapping-table">
      <thead><tr>
        <th>測項</th><th>檢測方法</th>${cat.unitField ? '<th>單位代碼</th>' : ''}<th>筆數</th>
      </tr></thead>
      <tbody id="methodItemsBody">
        ${entries.map(([item, g]) => `<tr data-item="${escapeAttr(item)}">
          <td>${escapeHtml(item)}${!g.method ? ' <span class="req" title="尚未有檢測方法">＊</span>' : ''}</td>
          <td><input type="text" data-method-field="method" value="${escapeAttr(g.method)}" placeholder="例：NIEA W417"></td>
          ${cat.unitField ? `<td><input type="text" data-method-field="unitCode" value="${escapeAttr(g.unitCode || '')}" class="code-input" data-codetype="unit" title="${escapeAttr(lookupUnit(g.unitCode))}" placeholder="代碼"></td>` : ''}
          <td>${g.indices.length}</td>
        </tr>`).join('')}
      </tbody>
    </table>`;
    wrap.querySelectorAll('.code-input').forEach(inp => {
      inp.addEventListener('input', () => { inp.title = lookupUnit(inp.value); });
    });
  }

  document.getElementById('methodModal').dataset.projectId = project.id;
  document.getElementById('methodModal').dataset.catKey = catKey;
  document.getElementById('btnMethodSave').classList.toggle('hidden', entries.length === 0);
  document.getElementById('methodModal').classList.remove('hidden');
}

function closeMethodModal() {
  document.getElementById('methodModal').classList.add('hidden');
}

function saveMethodModal() {
  const modal = document.getElementById('methodModal');
  const projectId = modal.dataset.projectId;
  const catKey = modal.dataset.catKey;
  const cat = CATEGORIES[catKey];
  const rows = DataStore.getData(projectId, catKey);
  const itemField = cat.itemField;
  const memoryUpdates = {};
  const touchedRows = [];

  // Same undo protection as every other multi-row change (bulk delete/edit, coord
  // manager) — mistyping a method/unit here shouldn't be the one thing a person
  // can't recover from.
  pushUndoSnapshot(projectId, catKey, `套用檢測方法／單位代碼`);

  document.querySelectorAll('#methodItemsBody tr').forEach(tr => {
    const item = tr.dataset.item;
    const values = {};
    tr.querySelectorAll('[data-method-field]').forEach(el => {
      const targetField = el.dataset.methodField === 'method' ? cat.methodField : cat.unitField;
      if (targetField) values[targetField] = el.value;
    });
    rows.forEach(row => {
      const rowItem = (row[itemField] || '').trim() || '（未命名項目）';
      if (rowItem === item) { Object.assign(row, values); touchedRows.push(row); }
    });
    const memFields = {};
    if (values[cat.methodField]) memFields.method = values[cat.methodField];
    if (cat.unitField && values[cat.unitField]) memFields.unitCode = values[cat.unitField];
    if (Object.keys(memFields).length) memoryUpdates[item] = memFields;
  });

  DataStore.saveData(projectId, catKey, rows);
  if (Object.keys(memoryUpdates).length) DataStore.updateItemMemory(projectId, catKey, memoryUpdates);
  // Also refresh the FULL per-location snapshot memory, not just the flat item-name
  // memory above — otherwise a correction made here wouldn't carry through to next
  // season's "entirely absent location" reconstruction, which prefers the fuller
  // snapshot over the flat item memory whenever both exist, and would keep
  // resurrecting the pre-correction method/unit indefinitely.
  if (touchedRows.length > 0) learnSiteItemHistory(projectId, catKey, cat, touchedRows);
  closeMethodModal();
  renderContent();
  alert('已套用檢測方法／單位代碼到符合的資料列，並記住這些設定供下次（含下一季）匯入同一測項時自動帶入。');
}

// ---------- version / changelog ----------
function renderVersionBadge() {
  const badge = document.getElementById('versionBadge');
  badge.textContent = APP_VERSION;
}
function openChangelogModal() {
  document.getElementById('changelogCurrentVersion').textContent = APP_VERSION;
  const list = document.getElementById('changelogList');
  list.innerHTML = CHANGELOG.map(entry => `
    <div class="changelog-entry">
      <div class="changelog-header"><strong>${escapeHtml(entry.version)}</strong> <span class="hint">${escapeHtml(entry.date)}</span></div>
      <ul>${entry.notes.map(n => `<li>${escapeHtml(n)}</li>`).join('')}</ul>
    </div>
  `).join('');
  document.getElementById('changelogModal').classList.remove('hidden');
}

// ---------- unit code reference ----------
function renderUnitRefTable(filterText) {
  const body = document.getElementById('unitRefBody');
  const q = (filterText || '').trim().toLowerCase();
  const entries = Object.entries(UNIT_CODES).filter(([code, name]) =>
    !q || code.toLowerCase().includes(q) || String(name).toLowerCase().includes(q)
  );
  body.innerHTML = entries.length
    ? entries.map(([code, name]) => `<tr><td>${escapeHtml(code)}</td><td>${escapeHtml(name)}</td></tr>`).join('')
    : '<tr><td colspan="2" class="hint">找不到符合的單位</td></tr>';
}
function openUnitRefModal() {
  document.getElementById('unitRefSearch').value = '';
  renderUnitRefTable('');
  document.getElementById('unitRefModal').classList.remove('hidden');
}

function renderAgencyRefTable(filterText) {
  const body = document.getElementById('agencyRefBody');
  const q = (filterText || '').trim().toLowerCase();
  const entries = Object.entries(AGENCY_CODES).filter(([code, name]) =>
    !q || code.toLowerCase().includes(q) || String(name).toLowerCase().includes(q)
  );
  body.innerHTML = entries.length
    ? entries.map(([code, name]) => `<tr><td>${escapeHtml(code)}</td><td>${escapeHtml(name)}</td></tr>`).join('')
    : '<tr><td colspan="2" class="hint">找不到符合的檢測機構</td></tr>';
}
function openAgencyRefModal() {
  document.getElementById('agencyRefSearch').value = '';
  renderAgencyRefTable('');
  document.getElementById('agencyRefModal').classList.remove('hidden');
}

// ---------- site-wide shared config — plain TEXT files in clearly-named folders,
// no JSON/code syntax to get right, safe for a non-programmer to hand-edit in
// Notepad ----------
// A static site has no shared write-storage: nothing typed into this app itself can
// ever reach OTHER people's browsers, since each browser only has its own local
// storage. The only way an update can reach "everyone who opens this URL" is for it
// to be part of the deployed static files themselves — same mechanism as any other
// update to this app. These two folders exist specifically so that updating the
// shared EIAS link or the template list doesn't require touching app.js OR
// understanding any file-format syntax: 網址.txt is nothing but the URL itself on
// one line; 範本清單.txt is one "顯示名稱|檔案名稱" line per template. Replace the
// text, redeploy, and every visitor gets the update automatically the next time
// they load the page — with no in-app "anyone can upload something that spreads to
// other people" feature at all, since that upload path is exactly what would let a
// malicious file or link reach strangers. Updates only flow through whoever already
// has push access to the repo — the same trust boundary this app's own updates
// have always relied on.
const EIAS_URL_FOLDER = '更新網址的話丟到本資料夾';
const TEMPLATE_FOLDER = '更新範本的話請將檔案丟到本資料夾';
function eiasUrlFileUrl() { return `${encodeURIComponent(EIAS_URL_FOLDER)}/${encodeURIComponent('網址.txt')}`; }
function templateListFileUrl() { return `${encodeURIComponent(TEMPLATE_FOLDER)}/${encodeURIComponent('範本清單.txt')}`; }
function templateFileUrl(filename) { return `${encodeURIComponent(TEMPLATE_FOLDER)}/${encodeURIComponent(filename)}`; }

let siteConfig = { eiasUrl: null, templates: [], itemCatalog: { air: [], water: [], noise: [], geo: [], eco: [] } };

/** Parses "顯示名稱|檔案名稱" lines, one per template — far more forgiving to
 *  hand-edit in Notepad than JSON (no commas/brackets/quotes to get exactly right).
 *  Lines starting with # are treated as comments (used for the in-file instructions
 *  shipped in 範本清單.txt) and blank lines are ignored. Strips a UTF-8 BOM in case
 *  Notepad added one when saving. */
function parseTemplateListText(text) {
  return String(text || '')
    .replace(/^\uFEFF/, '')
    .split(/\r?\n/)
    .map(line => line.trim())
    .filter(line => line && !line.startsWith('#'))
    .map(line => {
      const idx = line.indexOf('|');
      if (idx === -1) return null;
      const label = line.slice(0, idx).trim();
      const file = line.slice(idx + 1).trim();
      return (label && file) ? { label, file } : null;
    })
    .filter(Boolean);
}

const ITEM_CATALOG_FOLDER = '更新測項單位方法對照表的話請編輯本資料夾';
function itemCatalogFileUrl() { return `${encodeURIComponent(ITEM_CATALOG_FOLDER)}/${encodeURIComponent('測項對照表.txt')}`; }

/** Parses "類別|測項名稱|檢測方法|單位代碼" lines into { air: [...], water: [...], ... }
 *  where each category's array holds { itemName, variants: [{method, unitCode}] } —
 *  multiple lines with the same category+itemName merge into one entry with
 *  multiple variants (e.g. 溶氧 measured by either 電極法/W455 or 碘定量法/W422).
 *  Same forgiving format as the other shared text files: # comments, blank lines
 *  ignored, BOM stripped. */
function parseItemCatalogText(text) {
  const catalog = { air: [], water: [], noise: [], geo: [], eco: [] };
  const byKey = {}; // `${cat}::${itemName}` -> entry, so repeated lines merge
  String(text || '')
    .replace(/^\uFEFF/, '')
    .split(/\r?\n/)
    .map(line => line.trim())
    .filter(line => line && !line.startsWith('#'))
    .forEach(line => {
      const parts = line.split('|');
      if (parts.length < 2) return;
      const cat = parts[0].trim();
      const itemName = (parts[1] || '').trim();
      const method = (parts[2] || '').trim();
      const unitCode = (parts[3] || '').trim();
      if (!catalog[cat] || !itemName) return;
      const key = `${cat}::${itemName}`;
      if (!byKey[key]) {
        const entry = { itemName, variants: [] };
        byKey[key] = entry;
        catalog[cat].push(entry);
      }
      byKey[key].variants.push({ method, unitCode });
    });
  return catalog;
}

async function loadSiteConfig() {
  try {
    const resp = await fetch(eiasUrlFileUrl());
    if (resp.ok) {
      const text = (await resp.text()).replace(/^\uFEFF/, '').trim();
      if (/^https?:\/\/.+/i.test(text)) siteConfig.eiasUrl = text;
    }
  } catch (e) {
    // Opened via file:// (double-clicked instead of served over http), offline, or
    // the file is missing — fall back to the hardcoded default rather than
    // breaking the page.
    console.warn('無法載入共用的申報網站網址設定，改用內建預設值。', e);
  }
  try {
    const resp = await fetch(templateListFileUrl());
    if (resp.ok) siteConfig.templates = parseTemplateListText(await resp.text());
  } catch (e) {
    console.warn('無法載入範本清單，範本下載功能將暫不顯示。', e);
  }
  try {
    const resp = await fetch(itemCatalogFileUrl());
    if (resp.ok) siteConfig.itemCatalog = parseItemCatalogText(await resp.text());
  } catch (e) {
    console.warn('無法載入常用測項對照表，「常用測項新增」將暫不提供內建測項（仍可手動新增自訂測項）。', e);
  }
  renderEiasLink();
  renderTemplateDownloadButton();
}

// ---------- EIAS (環評線上申報) website link ----------
// Two layers: the SHARED default comes from 網址.txt (see above) — the same for
// everyone who opens this site. A person can additionally set a PERSONAL override
// in their own browser (localStorage) if, for whatever reason, they want their own
// browser to point somewhere different without needing repo access — but that
// override only ever affects their own browser, never anyone else's.
const EIAS_DEFAULT_URL = 'https://eias.moenv.gov.tw/';
function getSharedEiasUrl() { return siteConfig.eiasUrl || EIAS_DEFAULT_URL; }
function getEiasUrl() {
  return localStorage.getItem('envapp_eiasUrl') || getSharedEiasUrl();
}
function renderEiasLink() {
  const link = document.getElementById('btnEiasLink');
  if (link) link.href = getEiasUrl();
}
function openEiasLinkEditModal() {
  document.getElementById('eiasLinkInput').value = getEiasUrl();
  document.getElementById('eiasLinkEditModal').classList.remove('hidden');
}
function saveEiasLinkEdit() {
  const raw = document.getElementById('eiasLinkInput').value.trim();
  if (!/^https?:\/\/.+/i.test(raw)) {
    alert('請輸入完整的網址，需以 http:// 或 https:// 開頭。');
    return;
  }
  localStorage.setItem('envapp_eiasUrl', raw);
  renderEiasLink();
  document.getElementById('eiasLinkEditModal').classList.add('hidden');
}

// ---------- template downloads (blank filing templates, listed in 範本清單.txt) ----------
// The toolbar always has exactly ONE "📥 範本下載" button regardless of how many
// templates exist — adding, removing, or renaming templates only ever changes what
// shows up INSIDE this one modal (via 範本清單.txt), never adds more toolbar
// clutter. Each template gets its own individual download link so picking a single
// category never means downloading all of them.
function renderTemplateDownloadButton() {
  const btn = document.getElementById('btnTemplateDownload');
  if (btn) btn.style.display = siteConfig.templates.length > 0 ? '' : 'none';
}
function markTemplateFileMissing(idx, filename) {
  const li = document.querySelector(`.template-download-item[data-idx="${idx}"]`);
  if (!li) return;
  li.classList.add('template-missing');
  const note = document.createElement('div');
  note.className = 'hint template-missing-note';
  note.textContent = `⚠️ 在資料夾裡找不到「${filename}」這個檔案，請確認檔案已放進「${TEMPLATE_FOLDER}」資料夾，且檔名跟範本清單.txt裡寫的完全一致（含副檔名）。`;
  li.appendChild(note);
}
function openTemplateDownloadModal() {
  const body = document.getElementById('templateDownloadBody');
  if (siteConfig.templates.length === 0) {
    body.innerHTML = '<li class="hint">目前沒有可下載的範本。</li>';
    document.getElementById('templateDownloadModal').classList.remove('hidden');
    return;
  }
  body.innerHTML = siteConfig.templates.map((t, i) => `
    <li class="template-download-item" data-idx="${i}">
      <span>${escapeHtml(t.label)}</span>
      <a class="btn btn-primary btn-sm" href="${templateFileUrl(t.file)}" download="${escapeAttr(t.file)}">⬇️ 下載</a>
    </li>
  `).join('');
  document.getElementById('templateDownloadModal').classList.remove('hidden');

  // Quietly verify each file actually exists (a HEAD request, not a real download)
  // so a typo in 範本清單.txt or a forgotten file shows up as a visible warning
  // right where the person would notice it, instead of a silently broken link.
  siteConfig.templates.forEach(async (t, i) => {
    try {
      const resp = await fetch(templateFileUrl(t.file), { method: 'HEAD' });
      if (!resp.ok) markTemplateFileMissing(i, t.file);
    } catch (e) {
      markTemplateFileMissing(i, t.file);
    }
  });
}

// ---------- export selection (choose which categories to include) ----------
function openExportSelectModal(project) {
  const list = document.getElementById('exportSelectList');
  const withData = CATEGORY_ORDER.filter(c => DataStore.getData(project.id, c).length > 0);
  list.innerHTML = withData.map(c => `
    <label class="item-check">
      <input type="checkbox" data-export-cat value="${c}" checked>
      ${CATEGORIES[c].label}（${DataStore.getData(project.id, c).length} 筆）
    </label>
  `).join('');

  // union of every period label seen across the categories being offered, so one
  // choice here applies to all of them — matches how a submission is normally
  // compiled (this whole quarter's data across every category, all at once).
  const allPeriods = new Set();
  let anyUnlabeled = false;
  withData.forEach(c => {
    DataStore.getData(project.id, c).forEach(r => { r._period ? allPeriods.add(r._period) : (anyUnlabeled = true); });
  });
  const periodWrap = document.getElementById('exportSelectPeriodWrap');
  if (allPeriods.size > 0 || anyUnlabeled) {
    periodWrap.innerHTML = `
      <label style="display:flex;flex-direction:column;gap:5px;margin:10px 0;font-size:13px;font-weight:600;">
        只匯出哪一期的資料？
        <select id="exportSelectPeriod">
          <option value="">全部期別</option>
          ${[...allPeriods].sort().map(p => `<option value="${escapeAttr(p)}">${escapeHtml(p)}</option>`).join('')}
          ${anyUnlabeled ? `<option value="__none__">未標示期別</option>` : ''}
        </select>
      </label>`;
  } else {
    periodWrap.innerHTML = '';
  }

  document.getElementById('exportSelectModal').dataset.projectId = project.id;
  document.getElementById('exportSelectModal').classList.remove('hidden');
}
function closeExportSelectModal() {
  document.getElementById('exportSelectModal').classList.add('hidden');
}
function confirmExportSelect() {
  const modal = document.getElementById('exportSelectModal');
  const project = getCurrentProject();
  const checked = [...document.querySelectorAll('#exportSelectList [data-export-cat]:checked')].map(cb => cb.value);
  if (checked.length === 0) { alert('請至少勾選一個類別。'); return; }
  const periodSel = document.getElementById('exportSelectPeriod');
  const period = periodSel ? periodSel.value : '';
  const basicInfo = DataStore.getBasicInfo(project.id);
  let exportedCount = 0;
  checked.forEach(catKey => {
    let rows = DataStore.getData(project.id, catKey);
    if (period === '__none__') rows = rows.filter(r => !r._period);
    else if (period) rows = rows.filter(r => r._period === period);
    if (rows.length === 0) return; // nothing for this category in the chosen period
    ExportEngine.downloadCategory(project, basicInfo, catKey, rows);
    exportedCount++;
  });
  closeExportSelectModal();
  if (exportedCount > 0) {
    showToast('📥 已匯出。<strong>若之後修改這些檔案</strong>（手動補值、新增測項等），完成後可到各分類頁面上方點「📥 匯入資料」按鈕、選擇修改過的檔案<strong>重新匯入即可</strong>——系統會自動比對，內容相同的資料會忽略，內容不同的會列出讓您確認要不要用新版本取代，不會憑空覆蓋或造成重複。', 9000);
  }
}

// ---------- project modal (create/edit) ----------
function openProjectModal(project) {
  state.editingProjectId = project ? project.id : null;
  document.getElementById('projectModalTitle').textContent = project ? '編輯計畫資訊' : '新增計畫';
  document.getElementById('projectCodeInput').value = project ? project.code : '';
  document.getElementById('projectNameInput').value = project ? project.name : '';
  document.getElementById('projectModal').classList.remove('hidden');
  document.getElementById('projectCodeInput').focus();
}
function closeProjectModal() {
  document.getElementById('projectModal').classList.add('hidden');
  state.editingProjectId = null;
}
function saveProjectModal() {
  const code = document.getElementById('projectCodeInput').value.trim();
  const name = document.getElementById('projectNameInput').value.trim();
  if (!code || !name) { alert('請填寫計畫代碼與計畫名稱。'); return; }
  if (state.editingProjectId) {
    DataStore.updateProject(state.editingProjectId, { code, name });
    // keep basic info's 計畫代碼 in sync if user wants — do not force overwrite, just leave as-is
  } else {
    const p = DataStore.createProject(code, name);
    // The new project itself is always created (harmless), but switching the
    // screen to it goes through the same guard as selectProject — if an import is
    // open and the person chooses to keep it, the new project still exists (they
    // can switch to it later from the sidebar) but the screen stays put.
    if (confirmProjectSwitchIfImporting(p.id)) state.currentProjectId = p.id;
  }
  closeProjectModal();
  renderProjectList();
  renderContent();
}
function deleteProjectFlow(project) {
  if (!confirm(`確定要刪除計畫「${project.code} ${project.name}」嗎？此操作將刪除該計畫所有已輸入的監測資料，且無法復原。`)) return;
  DataStore.deleteProject(project.id);
  // Clear this project's undo/redo history too — it's keyed by projectId::catKey so
  // stale entries here could never be mis-restored into a different project, but
  // there's no reason to keep holding onto snapshots for data that no longer exists.
  Object.keys(state.undoStack).forEach(k => { if (k.startsWith(`${project.id}::`)) delete state.undoStack[k]; });
  Object.keys(state.redoStack).forEach(k => { if (k.startsWith(`${project.id}::`)) delete state.redoStack[k]; });
  if (state.currentProjectId === project.id) state.currentProjectId = null;
  renderProjectList();
  renderContent();
}

// ---------- item-selection checklist (shared by smart and generic import) ----------
// Different projects report different sets of monitoring items (e.g. an air-quality
// report might contain SO2/NO2/NOx/NO/CO/O3/PM10/TSP/PM2.5, but a given filing may only
// need a subset). Rather than hard-coding which items "belong" in a filing, the parser
// extracts everything the report actually contains and this checklist lets the person
// choose which ones to bring in.
function renderItemChecklist(containerEl, rows, itemField, onChange) {
  const counts = {};
  // An item is PRIMARY unless every single row for it is marked secondary. Parsers
  // set `_secondaryItem` on readings a report states but a filing doesn't normally
  // submit: the hour-by-hour block behind a 24-hour noise report's L日/L晚/L夜, or the
  // Lveq stated next to the Lvd(10) that actually gets filed. Those are read, listed
  // and available — just folded away and unticked, so a 3-row import doesn't silently
  // become a 264-row one.
  const primary = new Set();
  rows.forEach(r => {
    const v = (r[itemField] || '').trim() || '（未標示）';
    counts[v] = (counts[v] || 0) + 1;
    if (!r._secondaryItem) primary.add(v);
  });
  const items = Object.keys(counts);
  if (items.length === 0) { containerEl.innerHTML = ''; return; }
  const primaryItems = items.filter(i => primary.has(i));
  const secondaryItems = items.filter(i => !primary.has(i));
  // Nothing is marked primary (every block was a fallback guess) — fall back to
  // showing everything normally rather than presenting an empty checklist.
  const showAllAsPrimary = primaryItems.length === 0;
  const mainItems = showAllAsPrimary ? items : primaryItems;
  const extraItems = showAllAsPrimary ? [] : secondaryItems;
  if (!state.itemSelection) state.itemSelection = new Set(mainItems);

  const checkboxFor = (item) => `
    <label class="item-check">
      <input type="checkbox" data-item-check value="${escapeAttr(item)}" ${state.itemSelection.has(item) ? 'checked' : ''}>
      ${escapeHtml(item)} <span class="hint">(${counts[item]})</span>
    </label>`;
  const extraCount = extraItems.reduce((n, i) => n + counts[i], 0);
  const wasOpen = containerEl.querySelector('.extra-items-toggle')?.open;

  containerEl.innerHTML = `
    <p class="hint">此份報告偵測到以下監測項目，請勾選要匯入的項目（可依實際需求增減）：</p>
    <div class="item-checklist">
      ${mainItems.map(checkboxFor).join('')}
      <button type="button" class="btn btn-ghost btn-sm" id="btnItemSelectAll">全選</button>
      <button type="button" class="btn btn-ghost btn-sm" id="btnItemSelectNone">全不選</button>
    </div>
    ${extraItems.length === 0 ? '' : `
    <details class="row-detail-toggle extra-items-toggle" ${wasOpen ? 'open' : ''}>
      <summary>➕ 其他偵測到的測項（${extraItems.length} 種、共 ${extraCount} 筆，預設不匯入）</summary>
      <p class="hint" style="margin:6px 0">
        報告裡還有這些數值，但一般申報不會送出（例如 24 小時報告的逐時測值，只有 L日／L晚／L夜 需要申報）。
        確實需要的話在這裡勾選即可，勾了就會一起匯入。
      </p>
      <div class="item-checklist">
        ${extraItems.map(checkboxFor).join('')}
        <button type="button" class="btn btn-ghost btn-sm" id="btnExtraSelectAll">全選其他測項</button>
        <button type="button" class="btn btn-ghost btn-sm" id="btnExtraSelectNone">全不選其他測項</button>
      </div>
    </details>`}
  `;
  containerEl.querySelectorAll('[data-item-check]').forEach(cb => {
    cb.addEventListener('change', () => {
      if (cb.checked) state.itemSelection.add(cb.value); else state.itemSelection.delete(cb.value);
      if (onChange) onChange();
    });
  });
  const bulk = (btnId, targets, add) => {
    const btn = containerEl.querySelector(btnId);
    if (!btn) return;
    btn.addEventListener('click', () => {
      targets.forEach(i => { if (add) state.itemSelection.add(i); else state.itemSelection.delete(i); });
      renderItemChecklist(containerEl, rows, itemField, onChange);
      if (onChange) onChange();
    });
  };
  // "全選"/"全不選" act on the main list only — the folded-away extras have their
  // own pair, so clicking 全選 can never pull in hundreds of rows the person
  // hasn't looked at.
  bulk('#btnItemSelectAll', mainItems, true);
  bulk('#btnItemSelectNone', mainItems, false);
  bulk('#btnExtraSelectAll', extraItems, true);
  bulk('#btnExtraSelectNone', extraItems, false);
}

/**
 * Some reports state SEVERAL different statistics for the same measurement on the
 * same day — a 24-hour air-quality report gives每小時測值, 日平均值, 最大小時平均值,
 * 最小小時平均值 and 最大8小時平均值 for every pollutant. Which one belongs in a filing
 * is a decision only the person can make, so the parser reads them all, tags each row
 * with `_statKind`, and this checklist picks. 日平均值 is pre-selected because that is
 * what these filings normally report.
 *
 * Rows for anything other than the daily average carry the statistic in their 檢測項目
 * name ("CO最大8小時平均值") — see AIR_STAT_ROWS in smartparse.js for why.
 */
function renderStatChecklist(containerEl, rows, onChange) {
  if (!containerEl) return;
  const kinds = [];
  rows.forEach(r => {
    if (!r._statKind) return;
    let k = kinds.find(x => x.key === r._statKind);
    if (!k) { k = { key: r._statKind, label: r._statLabel || r._statKind, count: 0 }; kinds.push(k); }
    k.count++;
  });
  if (kinds.length <= 1) { containerEl.innerHTML = ''; state.statSelection = null; return; }
  if (!state.statSelection) {
    state.statSelection = new Set(kinds.some(k => k.key === 'avg') ? ['avg'] : [kinds[0].key]);
  }
  containerEl.innerHTML = `
    <div class="warning" style="background:#e8f0fe;border-color:#a8c7fa;">
      📊 這份報告同一個測項給了不只一種數值，請選擇這次要匯入哪一種（預設為「日平均值」，也就是一般申報用的數值）：
      <div class="item-checklist" style="margin-top:6px">
        ${kinds.map(k => `
          <label class="item-check">
            <input type="checkbox" data-stat-check value="${escapeAttr(k.key)}" ${state.statSelection.has(k.key) ? 'checked' : ''}>
            ${escapeHtml(k.label)} <span class="hint">(${k.count})</span>
          </label>
        `).join('')}
      </div>
      <p class="hint" style="margin:6px 0 0 0">
        「日平均值」的檢測項目維持原本的名稱（例如 <code>CO</code>）；其他統計值會把名稱寫清楚（例如 <code>CO最大8小時平均值</code>），
        這樣同一天同一測站同一測項才不會出現兩筆分不出來的資料。若同時勾選多種，兩種都會匯入。
      </p>
    </div>`;
  containerEl.querySelectorAll('[data-stat-check]').forEach(cb => {
    cb.addEventListener('change', () => {
      if (cb.checked) state.statSelection.add(cb.value); else state.statSelection.delete(cb.value);
      if (state.statSelection.size === 0) { cb.checked = true; state.statSelection.add(cb.value); return; }
      state.itemSelection = null;      // the item list changes with the statistic
      state.excludedRowIndices = new Set();
      if (onChange) onChange();
    });
  });
}

/** Applies ONLY the statistic filter (see renderStatChecklist). Rows with no
 *  statistic tag — every other report type — always pass through. */
function filterRowsByStat(rows) {
  if (!state.statSelection || state.statSelection.size === 0) return rows;
  return rows.filter(r => !r._statKind || state.statSelection.has(r._statKind));
}

function filterRowsBySelection(rows, itemField) {
  let filtered = filterRowsByStat(rows);
  if (state.itemSelection) {
    filtered = filtered.filter(r => state.itemSelection.has((r[itemField] || '').trim() || '（未標示）'));
  }
  if (state.excludedRowIndices && state.excludedRowIndices.size > 0) {
    filtered = filtered.filter(r => !state.excludedRowIndices.has(r._rowUid));
  }
  return filtered;
}

/**
 * A row-level checklist below the item-type checklist — lets the person exclude
 * specific individual records (e.g. "this one CO reading from this date, but not
 * the rest") rather than only being able to exclude an entire item type at once.
 * Shows whatever the item checklist currently leaves in (so it doesn't duplicate
 * rows already excluded that way), collapsed by default since it can be long.
 */
function renderRowDetailTable(containerEl, rows, cat) {
  if (!state.excludedRowIndices) state.excludedRowIndices = new Set();
  if (!state.rowDetailColumnFilters) state.rowDetailColumnFilters = {}; // { loc: Set, item: Set }
  if (rows.length === 0) { containerEl.innerHTML = ''; return; }
  const locField = cat.locationField;
  const itemField = cat.itemField;
  const valueField = cat.fields.find(f => ['檢測數值', '監測數值'].includes(f.key))?.key || '';
  const valueLabel = cat.fields.find(f => f.key === valueField)?.label || '數值';
  const locFilterActive = state.rowDetailColumnFilters.loc && state.rowDetailColumnFilters.loc.size > 0;
  const itemFilterActive = state.rowDetailColumnFilters.item && state.rowDetailColumnFilters.item.size > 0;

  const wasOpen = containerEl.querySelector('details')?.open;
  containerEl.innerHTML = `
    <details class="row-detail-toggle" ${wasOpen ? 'open' : ''}>
      <summary>📋 詳細資料列表（可個別排除不匯入的資料，共 ${rows.length} 筆）</summary>
      <div style="display:flex;gap:8px;align-items:center;margin:8px 0;flex-wrap:wrap">
        <input type="text" id="rowDetailSearchInput" placeholder="🔍 搜尋地點、測項等關鍵字，快速篩選" style="flex:1;min-width:200px;max-width:360px;padding:6px 9px;border:1px solid var(--border);border-radius:5px;font-family:inherit;font-size:13px">
        <button type="button" id="rowDetailSelectAllVisible" class="btn btn-ghost btn-sm">勾選目前顯示的</button>
        <button type="button" id="rowDetailClearAllVisible" class="btn btn-ghost btn-sm">取消勾選目前顯示的</button>
      </div>
      <div class="mapping-table-wrap" style="max-height:320px">
        <table class="mapping-table">
          <thead><tr>
            <th>${checkboxHTML(' id="rowDetailCheckAll"', true)}</th>
            <th>
              <span class="th-label">地點</span>
              <button type="button" class="col-filter-btn${locFilterActive ? ' col-filter-active' : ''}" data-row-detail-field="loc" title="篩選地點">▾</button>
            </th>
            <th>
              <span class="th-label">測項</span>
              <button type="button" class="col-filter-btn${itemFilterActive ? ' col-filter-active' : ''}" data-row-detail-field="item" title="篩選測項">▾</button>
            </th>
            <th>日期</th><th>時間</th><th>${escapeHtml(valueLabel)}</th>
          </tr></thead>
          <tbody id="rowDetailTbody">
            ${rows.map(r => `
              <tr data-search-text="${escapeAttr([r[locField], r[itemField], r['日期(起)'], r['時間(起)'], r[valueField]].filter(Boolean).join(' ').toLowerCase())}" data-loc-value="${escapeAttr(r[locField] || '')}" data-item-value="${escapeAttr(r[itemField] || '')}">
                <td>${checkboxHTML(` class="row-detail-check" data-row-uid="${r._rowUid}"`, !state.excludedRowIndices.has(r._rowUid))}</td>
                <td>${escapeHtml(r[locField] || '')}</td>
                <td>${escapeHtml(r[itemField] || '')}</td>
                <td>${escapeHtml(toDateDisplayValue(r['日期(起)']) || '')}</td>
                <td>${escapeHtml(toTimeDisplayValue(r['時間(起)']) || '')}</td>
                <td>${escapeHtml(r[valueField] || '')}</td>
              </tr>
            `).join('')}
          </tbody>
        </table>
      </div>
    </details>
  `;
  const checkAll = containerEl.querySelector('#rowDetailCheckAll');
  const getRowChecks = () => [...containerEl.querySelectorAll('.row-detail-check')];
  const getVisibleRowChecks = () => getRowChecks().filter(cb => !cb.closest('tr').classList.contains('row-hidden'));
  const syncCheckAllState = () => {
    const visible = getVisibleRowChecks();
    checkAll.checked = visible.length > 0 && visible.every(cb => cb.checked);
  };
  getRowChecks().forEach(cb => {
    cb.addEventListener('change', () => {
      const uid = Number(cb.dataset.rowUid);
      if (cb.checked) state.excludedRowIndices.delete(uid); else state.excludedRowIndices.add(uid);
      syncCheckAllState();
    });
  });
  // "Select all" only ever affects rows currently visible under the search/column
  // filter — same convention as the main data grid's search+select-all, so
  // filtering down to a handful of rows and clicking select-all doesn't silently
  // re-check hundreds of rows the person can't currently see.
  checkAll.addEventListener('change', () => {
    getVisibleRowChecks().forEach(cb => {
      cb.checked = checkAll.checked;
      const uid = Number(cb.dataset.rowUid);
      if (checkAll.checked) state.excludedRowIndices.delete(uid); else state.excludedRowIndices.add(uid);
    });
  });
  containerEl.querySelector('#rowDetailSelectAllVisible').addEventListener('click', () => {
    getVisibleRowChecks().forEach(cb => { cb.checked = true; state.excludedRowIndices.delete(Number(cb.dataset.rowUid)); });
    syncCheckAllState();
  });
  containerEl.querySelector('#rowDetailClearAllVisible').addEventListener('click', () => {
    getVisibleRowChecks().forEach(cb => { cb.checked = false; state.excludedRowIndices.add(Number(cb.dataset.rowUid)); });
    syncCheckAllState();
  });

  // Combines the free-text search with any active column filters (地點/測項) — a
  // row must satisfy all of them to stay visible, same AND logic as the main grid.
  const applyAllFilters = () => {
    const q = searchInput.value.trim().toLowerCase();
    const locFilter = state.rowDetailColumnFilters.loc;
    const itemFilter = state.rowDetailColumnFilters.item;
    containerEl.querySelectorAll('#rowDetailTbody tr').forEach(tr => {
      let matches = !q || tr.dataset.searchText.includes(q);
      if (matches && locFilter && locFilter.size > 0) matches = locFilter.has(tr.dataset.locValue);
      if (matches && itemFilter && itemFilter.size > 0) matches = itemFilter.has(tr.dataset.itemValue);
      tr.classList.toggle('row-hidden', !matches);
    });
    syncCheckAllState();
  };

  const searchInput = containerEl.querySelector('#rowDetailSearchInput');
  searchInput.addEventListener('input', applyAllFilters);

  containerEl.querySelectorAll('.col-filter-btn[data-row-detail-field]').forEach(btn => {
    btn.addEventListener('click', (e) => {
      e.stopPropagation();
      openRowDetailColumnFilterPopup(btn.dataset.rowDetailField, btn, rows, locField, itemField, applyAllFilters);
    });
  });

  applyAllFilters(); // re-apply any column filter left over from before this re-render
}

/** Excel-style AutoFilter popup for the 詳細資料列表's 地點/測項 columns — mirrors
 *  openColumnFilterPopup's UI, but reads from the in-memory candidate rows (this is
 *  an import preview, not saved DataStore data) and stores its filter state
 *  separately in state.rowDetailColumnFilters. */
function openRowDetailColumnFilterPopup(fieldType, btnEl, rows, locField, itemField, onApply) {
  const popup = document.getElementById('colFilterPopup');
  const getVal = fieldType === 'loc' ? (r => r[locField] || '') : (r => r[itemField] || '');
  const uniqueValues = [...new Set(rows.map(getVal))].sort((a, b) => a.localeCompare(b, 'zh-Hant'));
  const currentFilter = state.rowDetailColumnFilters[fieldType];

  popup.innerHTML = `
    <div class="col-filter-search"><input type="text" id="colFilterSearchInput" placeholder="搜尋選項..."></div>
    <div class="col-filter-actions">
      <button type="button" id="colFilterSelectAll" class="btn btn-ghost btn-sm">全選</button>
      <button type="button" id="colFilterClearAll" class="btn btn-ghost btn-sm">清除</button>
    </div>
    <div class="col-filter-list" id="colFilterList">
      ${uniqueValues.map(v => `
        <label class="col-filter-item">
          <input type="checkbox" value="${escapeAttr(v)}" ${!currentFilter || currentFilter.has(v) ? 'checked' : ''}>
          <span>${escapeHtml(v === '' ? '（空白）' : v)}</span>
        </label>
      `).join('')}
    </div>
    <div class="col-filter-buttons">
      <button type="button" id="colFilterCancel" class="btn btn-ghost btn-sm">取消</button>
      <button type="button" id="colFilterApply" class="btn btn-primary btn-sm">確定</button>
    </div>
  `;

  const rect = btnEl.getBoundingClientRect();
  const popupWidth = 240;
  popup.style.left = Math.max(4, Math.min(rect.left, window.innerWidth - popupWidth - 4)) + 'px';
  popup.style.top = (rect.bottom + window.scrollY + 4) + 'px';
  popup.classList.remove('hidden');

  document.getElementById('colFilterSearchInput').addEventListener('input', (e) => {
    const q = e.target.value.trim().toLowerCase();
    document.querySelectorAll('#colFilterList .col-filter-item').forEach(label => {
      label.style.display = (!q || label.textContent.toLowerCase().includes(q)) ? '' : 'none';
    });
  });
  document.getElementById('colFilterSelectAll').addEventListener('click', () => {
    document.querySelectorAll('#colFilterList input[type="checkbox"]').forEach(cb => { cb.checked = true; });
  });
  document.getElementById('colFilterClearAll').addEventListener('click', () => {
    document.querySelectorAll('#colFilterList input[type="checkbox"]').forEach(cb => { cb.checked = false; });
  });
  document.getElementById('colFilterCancel').addEventListener('click', () => {
    popup.classList.add('hidden');
  });
  document.getElementById('colFilterApply').addEventListener('click', () => {
    const checked = [...document.querySelectorAll('#colFilterList input[type="checkbox"]:checked')].map(cb => cb.value);
    if (checked.length === uniqueValues.length) {
      delete state.rowDetailColumnFilters[fieldType];
      btnEl.classList.remove('col-filter-active');
    } else {
      state.rowDetailColumnFilters[fieldType] = new Set(checked);
      btnEl.classList.add('col-filter-active');
    }
    popup.classList.add('hidden');
    onApply();
  });

  setTimeout(() => {
    const closeOnOutsideClick = (e) => {
      if (!popup.contains(e.target) && e.target !== btnEl) {
        popup.classList.add('hidden');
        document.removeEventListener('click', closeOnOutsideClick);
      }
    };
    document.addEventListener('click', closeOnOutsideClick);
  }, 0);
}

// ---------- batch import (multi-file, auto-detect category) ----------
// Categories with report parsers precise enough to identify their OWN report format,
// which is what batch import needs (it has to decide which category a file belongs to
// before it can import it). The layout-agnostic AutoDetect reader is deliberately NOT
// used here: it recognizes fields, not categories, so letting it vote would mean
// guessing the category from a guess.
const AUTO_DETECT_CATEGORIES = ['noise', 'water', 'air', 'geo'];

function openBatchImportModal() {
  document.getElementById('batchFileInput').value = '';
  document.getElementById('batchDetectStatus').textContent = '';
  if (state._batchStaging) state._batchStaging.reset();
  document.getElementById('batchImportModal').classList.remove('hidden');
}
function closeBatchImportModal() {
  document.getElementById('batchImportModal').classList.add('hidden');
}

async function handleBatchFiles(fileList) {
  const files = [...fileList];
  if (files.length === 0) return;
  const statusEl = document.getElementById('batchDetectStatus');
  statusEl.textContent = `正在判讀 ${files.length} 個檔案...`;

  // perCategory[catKey] = { rows: [], matchedSheets: [], skippedSheets: [], sourceFiles: Set }
  const perCategory = {};
  AUTO_DETECT_CATEGORIES.forEach(c => { perCategory[c] = { rows: [], matchedSheets: [], skippedSheets: [], sourceFiles: new Set() }; });
  const unrecognizedFiles = [];

  for (const file of files) {
    const isSpreadsheet = /\.(xlsx|xls)$/i.test(file.name);
    const isPdf = /\.pdf$/i.test(file.name);
    if (!isSpreadsheet && !isPdf) { unrecognizedFiles.push(`${file.name}（不支援的格式）`); continue; }
    let grids;
    try {
      grids = isSpreadsheet ? await ImportEngine.readWorkbookGrids(file) : await ImportEngine.readPdfAsGrids(file);
    } catch (err) {
      unrecognizedFiles.push(`${file.name}（讀取失敗：${err.message}）`);
      continue;
    }
    let anyMatchInFile = false;
    for (const [sheetName, grid] of Object.entries(grids)) {
      let matchedCat = null;
      for (const catKey of AUTO_DETECT_CATEGORIES) {
        const rows = SmartParse.parseSheet(catKey, sheetName, grid, { allowAutoDetect: false });
        if (rows && rows.length) {
          rows.forEach(r => { r._sourceFile = file.name; r._sourceSheet = sheetName; });
          perCategory[catKey].rows.push(...rows);
          perCategory[catKey].matchedSheets.push(`${file.name} / ${sheetName}`);
          perCategory[catKey].sourceFiles.add(file.name);
          if (isPdf) perCategory[catKey].hasPdfSource = true;
          matchedCat = catKey;
          anyMatchInFile = true;
          break; // a sheet belongs to exactly one category
        }
      }
      if (!matchedCat) {
        // don't list every unrecognized junk sheet individually across all files; tallied below instead
      }
    }
    if (!anyMatchInFile) unrecognizedFiles.push(file.name);
  }

  // build per-category "sites" groupings the same way parseWorkbook does
  AUTO_DETECT_CATEGORIES.forEach(catKey => {
    const rows = perCategory[catKey].rows;
    const sites = {};
    rows.forEach((row, i) => {
      const key = row._siteCode || row._rawLocation || `row${i}`;
      if (!sites[key]) sites[key] = { siteCode: row._siteCode || '', rawLocation: row._rawLocation || '', rowIndices: [] };
      sites[key].rowIndices.push(i);
    });
    perCategory[catKey].sites = sites;
  });

  const queue = AUTO_DETECT_CATEGORIES
    .filter(c => perCategory[c].rows.length > 0)
    .map(c => ({ catKey: c, result: perCategory[c] }));

  const summary = queue.map(q => `${CATEGORIES[q.catKey].label}：${q.result.rows.length} 筆`).join('、');
  statusEl.textContent = queue.length
    ? `判讀完成：${summary}。${unrecognizedFiles.length ? `無法判斷類別的檔案／工作表：${unrecognizedFiles.join('、')}` : ''}`
    : `無法從所選檔案中辨識出任何已支援的報告格式。${unrecognizedFiles.length ? `（${unrecognizedFiles.join('、')}）` : ''}`;

  if (queue.length === 0) return;

  // Lock the project for the whole batch, same reasoning as openImportModal — the
  // entire multi-category queue must land in whichever project was active when the
  // person picked these files, not wherever they happen to be looking by the time
  // each category in the queue gets confirmed.
  state.importProjectId = state.currentProjectId;
  state.batchQueue = queue;
  state.batchQueueTotal = queue.length;
  closeBatchImportModal();
  processNextBatchItem();
}

function processNextBatchItem() {
  if (!state.batchQueue || state.batchQueue.length === 0) {
    state.batchQueue = null;
    return;
  }
  const next = state.batchQueue[0];
  const doneCount = state.batchQueueTotal - state.batchQueue.length;
  state.importCatKey = next.catKey;
  state.importMode = 'smart';
  state.smartResult = next.result;
  state.currentImportSourceLabel = next.result.sourceFiles ? [...next.result.sourceFiles].join('、') : '批次匯入';
  document.getElementById('importModalTitle').textContent =
    `批次匯入（${doneCount + 1}/${state.batchQueueTotal}）：${CATEGORIES[next.catKey].label}`;
  document.getElementById('importModal').classList.remove('hidden');
  renderSmartImportPreview();
}

// ---------- import modal ----------
function openImportModal(catKey) {
  state.importCatKey = catKey;
  // Lock the project this import is for at the moment the modal opens — every
  // subsequent step (preview render, confirm, history learning) uses THIS id, not
  // a fresh getCurrentProject() lookup. Without this, switching projects in the
  // sidebar while the import modal is still open (before confirming) would silently
  // commit the parsed file into whichever project happens to be selected at confirm
  // time — confirmed as a real, serious bug: project A's file ended up entirely
  // inside project B's data, not just compared against the wrong history.
  state.importProjectId = state.currentProjectId;
  state.importParsed = null;
  state.importMode = null;
  state.smartResult = null;
  state.itemSelection = null;
  state.statSelection = null;
  state.excludedRowIndices = null;
  state.rowDetailColumnFilters = {};
  state.importPeriod = '';
  document.getElementById('importModalTitle').textContent = `匯入${CATEGORIES[catKey].label}監測資料`;
  // Per-category note about what this category's import can and can't do — set here
  // rather than hard-coded in the HTML, since the answer differs by category.
  const noticeEl = document.getElementById('importCategoryNotice');
  if (noticeEl) {
    if (catKey === 'eco') {
      noticeEl.className = 'warning warning-strong';
      noticeEl.innerHTML = `🌿 ${ECO_IMPORT_ONLY_NOTE}<br>`
        + `匯入完成版後，資料會照常進入下方表格、可編輯、可匯出成申報用的 Excel。`;
    } else if (SMART_PARSE_CATEGORIES.includes(catKey)) {
      noticeEl.className = 'notice-info';
      noticeEl.innerHTML = `本類別可直接讀原始檢測報告。若是系統沒看過的報告格式，會改用自動偵測，`
        + `盡量找出<strong>日期(起)／日期(迄)／時間(起)／時間(迄)／檢測項目／監測數值／檢測單位</strong>這幾個必要欄位，`
        + `並在下一步明確標示哪些是「猜」出來的、哪些沒找到。也可以直接匯入上一季的完成版 Excel。`;
    } else {
      noticeEl.className = '';
      noticeEl.innerHTML = '';
    }
  }
  document.getElementById('importFileInput').value = '';
  if (state._importStaging) state._importStaging.reset();
  document.getElementById('importStep1').classList.remove('hidden');
  document.getElementById('importStep2').classList.add('hidden');
  document.getElementById('importStep3').classList.add('hidden');
  document.getElementById('importPdfWarning').classList.add('hidden');
  document.getElementById('btnImportConfirm').classList.add('hidden');
  document.getElementById('importModal').classList.remove('hidden');
}
function closeImportModal() {
  document.getElementById('importModal').classList.add('hidden');
  if (state.batchQueue) {
    const remaining = state.batchQueue.length;
    state.batchQueue = null;
    if (remaining > 0) alert(`已取消批次匯入，還有 ${remaining} 個類別未匯入。`);
  }
}

// Categories whose raw monitoring reports the app will try to read directly.
// 生態 is deliberately absent: ecological surveys are filed by a separate contractor
// and this project never receives their raw survey report, only the finished
// 生態調查資料填寫.xlsx — so that category imports the completed file through the
// ordinary column-mapping path and says so plainly (see ECO_IMPORT_ONLY_NOTE).
const SMART_PARSE_CATEGORIES = ['noise', 'water', 'air', 'geo'];

const ECO_IMPORT_ONLY_NOTE = '生態分類目前<strong>只能匯入「已填寫完成」的生態調查資料填寫.xlsx（完成版）</strong>做資料彙整，'
  + '無法判讀生態監測／調查報告原始檔。生態調查通常由其他公司負責填報，請向對方索取填好的完成版 Excel 後再匯入本系統。';
const ECO_IMPORT_ONLY_TEXT = '生態分類目前只能匯入「已填寫完成」的生態調查資料填寫.xlsx（完成版）做資料彙整，無法判讀生態監測／調查報告原始檔。';

/**
 * Accepts one File, an array of Files, or a FileList (multi-select is now supported
 * for a single category too — e.g. importing this quarter's several monthly 放流水
 * reports at once). For smart-parse categories every selected file is read and
 * aggregated into one preview (one item checklist, one site-profile table) before
 * anything is committed. Non-smart-parse categories still take exactly one file,
 * since combining several raw tables via column-mapping isn't well-defined.
 */
async function handleImportFile(fileOrFiles) {
  if (!fileOrFiles) return;
  const files = fileOrFiles instanceof FileList ? [...fileOrFiles]
    : Array.isArray(fileOrFiles) ? fileOrFiles : [fileOrFiles];
  if (files.length === 0) return;
  const catKey = state.importCatKey;
  state.currentImportSourceLabel = files.map(f => f.name).join('、');

  const cat = CATEGORIES[catKey];

  try {
    if (SMART_PARSE_CATEGORIES.includes(catKey)) {
      const aggregate = { rows: [], matchedSheets: [], skippedSheets: [], sourceFiles: new Set() };
      let anyTemplateLike = false;
      const templateLikeFiles = [];
      for (const file of files) {
        const isSpreadsheet = /\.(xlsx|xls)$/i.test(file.name);
        const isPdf = /\.pdf$/i.test(file.name);
        if (!isSpreadsheet && !isPdf) { aggregate.skippedSheets.push(`${file.name}（不支援的格式）`); continue; }
        let grids;
        try {
          grids = isSpreadsheet ? await ImportEngine.readWorkbookGrids(file) : await ImportEngine.readPdfAsGrids(file);
        } catch (err) {
          console.error(err);
          aggregate.skippedSheets.push(`${file.name}（讀取失敗：${err.message}）`);
          continue;
        }
        // A file that is ALREADY in the official template's shape — a previous
        // season's completed filing, or this app's own export that the person
        // edited — must go through the ordinary column-mapping importer, which
        // reads it field-for-field. Running a report-form parser (or the layout
        // guesser) over it would re-derive values that are already correct.
        const templateLike = Object.values(grids).some(g => ImportEngine.gridSchemaMatchRatio(g, cat.fields) >= 0.5);
        if (templateLike) {
          anyTemplateLike = true;
          templateLikeFiles.push(file.name);
          continue;
        }
        // TWO PASSES PER FILE, and the second one only if the first found nothing.
        //
        // Pass 1 runs the template-specific parsers alone. If ANY sheet in the file
        // is recognized, the file's report format is known — and the sheets that
        // didn't match are then working sheets, not another report: a 氣象 tab, a
        // 工作表1 scratch pad, a 主辦 summary the consultant hand-typed for the
        // client. Turning the layout guesser loose on those invents items that
        // duplicate readings already read correctly from the real report sheets
        // (a 主辦 summary produced five bogus "(7～20)"-style rows next to the
        // 均能音量(Leq) rows the noise parser had already got right).
        //
        // Pass 2 — the guesser — therefore runs only when pass 1 came back empty for
        // the whole file, which is exactly the "unfamiliar lab's report" case it
        // exists for.
        const sheetEntries = Object.entries(grids);
        let fileRows = [];
        const unmatchedSheets = [];
        sheetEntries.forEach(([sheetName, grid]) => {
          const rows = SmartParse.parseSheet(catKey, sheetName, grid, { allowAutoDetect: false });
          if (rows && rows.length) {
            rows.forEach(r => { r._sourceSheet = sheetName; });
            fileRows.push(...rows);
            aggregate.matchedSheets.push(`${file.name} / ${sheetName}`);
          } else {
            unmatchedSheets.push(sheetName);
          }
        });
        if (fileRows.length === 0) {
          unmatchedSheets.splice(0).forEach(sheetName => {
            const rows = AutoDetect.parseSheet(catKey, sheetName, grids[sheetName]);
            if (rows && rows.length) {
              rows.forEach(r => { r._sourceSheet = sheetName; });
              fileRows.push(...rows);
              aggregate.matchedSheets.push(`${file.name} / ${sheetName}`);
              aggregate.hasAutoDetected = true;
            } else {
              unmatchedSheets.push(sheetName);
            }
          });
        }
        unmatchedSheets.forEach(sheetName => aggregate.skippedSheets.push(`${file.name} / ${sheetName}`));
        if (fileRows.length) {
          fileRows.forEach(r => { r._sourceFile = file.name; });
          aggregate.rows.push(...fileRows);
          aggregate.sourceFiles.add(file.name);
          if (isPdf) aggregate.hasPdfSource = true;
        }
      }
      if (anyTemplateLike && aggregate.rows.length === 0) {
        // Nothing here but files already in the official template's shape: hand over
        // to the column-mapping importer below, which reads them field-for-field.
      } else {
        // A template-format file mixed in with raw reports can't be shown on the
        // report-preview screen, so name the file that was left out rather than
        // dropping it silently.
        templateLikeFiles.forEach(n => aggregate.skippedSheets.push(`${n}（已是範本／完成版格式，請單獨匯入以進行欄位比對）`));
      }
      if (aggregate.rows.length > 0) {
        const sites = {};
        aggregate.rows.forEach((row, i) => {
          const key = row._siteCode || row._rawLocation || `row${i}`;
          if (!sites[key]) sites[key] = { siteCode: row._siteCode || '', rawLocation: row._rawLocation || '', rowIndices: [] };
          sites[key].rowIndices.push(i);
        });
        aggregate.sites = sites;
        state.importMode = 'smart';
        state.smartResult = aggregate;
        renderSmartImportPreview();
        return;
      }
      // fall through to generic single-file import if nothing recognized
    }

    if (files.length > 1) {
      alert('這個類別目前無法自動判讀多個檔案的原始報告格式，因此只能一次匯入一個檔案（一般欄位比對）。請改為逐一匯入，或確認檔案是否為可自動判讀的報告格式。');
      return;
    }

    const file = files[0];
    // Pass the target category's own field names in, so a workbook that already
    // matches the official template locks onto the right sheet and the right header
    // row instead of whichever sheet merely happens to be longest.
    const parsed = await ImportEngine.readFile(file, cat.fields.map(f => f.key));
    if (!parsed.rows || parsed.rows.length === 0) {
      alert(catKey === 'eco'
        ? `無法從此檔案讀取到任何資料列。\n\n${ECO_IMPORT_ONLY_TEXT}\n\n請確認您選的是填寫完成的「生態調查資料填寫.xlsx」，其中應有「生態檢測項目」工作表與「日期(起)／調查地點／學名／中文名…」等欄位標題。`
        : '無法從此檔案讀取到任何資料列，請確認檔案內容或改用 Excel 格式。');
      return;
    }

    // 生態: warn (but don't block) when the chosen file clearly isn't the completed
    // template — the person may still want to map columns by hand, but they should
    // know the app can't read a raw ecological survey report.
    if (catKey === 'eco') {
      const ratio = ImportEngine.schemaMatchRatio(parsed.headers, cat.fields);
      if (ratio < 0.4) {
        const ok = confirm(
          `${ECO_IMPORT_ONLY_TEXT}\n\n`
          + `這個檔案看起來不像填寫完成的「生態調查資料填寫.xlsx」（只比對到 ${Math.round(ratio * 100)}% 的欄位名稱）。\n\n`
          + `按「確定」仍可進入欄位對應畫面自行對應，按「取消」則停止匯入。`
        );
        if (!ok) return;
      }
    }

    state.importMode = 'generic';
    state.importParsed = parsed;
    document.getElementById('importPdfWarning').classList.toggle('hidden', !parsed.isPdfBestEffort);
    renderMappingStep();
  } catch (err) {
    console.error(err);
    alert('讀取檔案時發生錯誤：' + err.message);
  }
}

function renderMappingStep() {
  const project = getImportProject();
  const cat = CATEGORIES[state.importCatKey];
  const { headers, rows } = state.importParsed;
  const mapping = ImportEngine.suggestMapping(headers, cat.fields);
  state.excludedRowIndices = new Set();

  const tbody = document.getElementById('mappingTableBody');
  tbody.innerHTML = cat.fields.map(f => {
    const options = ['<option value="">（不匯入）</option>']
      .concat(headers.map(h => `<option value="${escapeAttr(h)}" ${mapping[f.key] === h ? 'selected' : ''}>${escapeHtml(h)}</option>`));
    return `<tr>
      <td>${escapeHtml(f.label)}${f.required ? '<span class="req">＊</span>' : ''}</td>
      <td class="${mapping[f.key] ? '' : 'unmatched'}">
        <select data-target-field="${f.key}">${options.join('')}</select>
      </td>
    </tr>`;
  }).join('');

  document.getElementById('importRowCount').textContent = `檔案共讀取到 ${rows.length} 筆資料列${state.importParsed.usedSheet ? `（工作表：${state.importParsed.usedSheet}）` : ''}。`;
  document.getElementById('importStep1').classList.add('hidden');
  document.getElementById('importStep2').classList.remove('hidden');
  document.getElementById('importStep3').classList.add('hidden');
  document.getElementById('btnImportConfirm').classList.remove('hidden');
  document.getElementById('btnImportConfirm').textContent = '確認匯入';

  const refreshGenericItemChecklist = () => {
    state.itemSelection = null; // re-seed from scratch when the mapped column changes
    const itemSel = tbody.querySelector(`select[data-target-field="${cat.itemField}"]`);
    const mappedHeader = itemSel ? itemSel.value : '';
    const wrap = document.getElementById('genericImportItemsWrap');
    if (!mappedHeader) { wrap.innerHTML = ''; document.getElementById('genericImportRowDetailWrap').innerHTML = ''; return; }
    const shimRows = rows.map(r => ({ [cat.itemField]: r[mappedHeader] }));
    renderItemChecklist(wrap, shimRows, cat.itemField, refreshGenericRowDetail);
    refreshGenericRowDetail();
  };
  const refreshGenericRowDetail = () => {
    const currentMapping = {};
    tbody.querySelectorAll('select[data-target-field]').forEach(sel => { currentMapping[sel.dataset.targetField] = sel.value || null; });
    const mappedRows = ImportEngine.applyMapping(rows, currentMapping, cat.fields);
    mappedRows.forEach((r, i) => { r._rowUid = i; });
    const itemFilteredRows = state.itemSelection
      ? mappedRows.filter(r => state.itemSelection.has((r[cat.itemField] || '').trim() || '（未標示）'))
      : mappedRows;
    renderRowDetailTable(document.getElementById('genericImportRowDetailWrap'), itemFilteredRows, cat);

    // Same method-diff notice as the smart-parse path: when the file explicitly
    // supplies a method that differs from what's remembered for that item, trust
    // the file (it's authoritative for this quarter's actual lab work), but flag it
    // so the person can confirm it's an intentional change rather than a mapping
    // mistake. This re-runs on every mapping change since which column feeds
    // 檢測方法 can change what gets compared.
    const noticeEl = document.getElementById('genericImportMethodDiffNotice');
    if (cat.methodField && noticeEl) {
      const memory = DataStore.getItemMemory(project.id, cat.key);
      const diffsByItem = {};
      mappedRows.forEach(row => {
        const mem = memory[row[cat.itemField]];
        if (mem && mem.method && row[cat.methodField] && row[cat.methodField] !== mem.method) {
          diffsByItem[row[cat.itemField]] = { reportMethod: row[cat.methodField], memoryMethod: mem.method };
        }
      });
      const diffs = Object.entries(diffsByItem).map(([item, d]) => ({ item, ...d }));
      noticeEl.innerHTML = diffs.length === 0 ? '' : `
        <div class="warning" style="background:#e8f0fe;border-color:#a8c7fa;">
          ℹ️ 以下項目本次檔案的檢測方法跟先前記錄不同，系統已依「本次檔案」為準。若並非刻意更換方法，請確認是否為判讀誤差：<br>
          ${diffs.map(d => `・${escapeHtml(d.item)}：本次「${escapeHtml(d.reportMethod)}」，先前記錄為「${escapeHtml(d.memoryMethod)}」`).join('<br>')}
        </div>`;
    }
  };
  refreshGenericItemChecklist();
  const itemSel = tbody.querySelector(`select[data-target-field="${cat.itemField}"]`);
  if (itemSel) itemSel.addEventListener('change', refreshGenericItemChecklist);

  const dateSel = tbody.querySelector(`select[data-target-field="日期(起)"]`);
  const refreshPeriodPicker = () => {
    const mappedHeader = dateSel ? dateSel.value : '';
    const shimRows = mappedHeader ? rows.map(r => ({ '日期(起)': normalizeDateString(r[mappedHeader]) })) : [];
    renderPeriodPicker('genericPeriodWrap', shimRows);
  };
  refreshPeriodPicker();
  if (dateSel) dateSel.addEventListener('change', refreshPeriodPicker);
}

// ---------- smart import (report-form parser) preview ----------
function renderSmartImportPreview() {
  const project = getImportProject();
  const catKey = state.importCatKey;
  const cat = CATEGORIES[catKey];
  const result = state.smartResult;
  const profileFields = SmartParse.SITE_PROFILE_FIELDS[catKey] || [];
  const savedAliases = DataStore.getSiteAliases(project.id, catKey);

  // Fill in method/unit from this project's remembered values (learned from past
  // confirmed imports, any season) wherever the report itself didn't supply one —
  // done once per parsed result, not on every checklist toggle re-render. When the
  // report DOES supply a method but it differs from what's remembered, the report
  // wins (it's authoritative for this quarter's actual lab work — e.g. dissolved
  // oxygen legitimately measured by either 電極法/NIEA W455 or 碘定量法/NIEA W422
  // depending on which the lab used that season), but it's flagged so the person
  // can confirm it's an intentional change rather than a parsing fluke.
  //
  // Separately, fill in any OTHER blank field from the historical full-row
  // snapshot for this exact (location, item-identity) combo — e.g. this quarter's
  // report only supplied date/time/value for an item that's otherwise unchanged,
  // but coordinates/管制標準/etc were present last season; carry those forward
  // rather than leaving them permanently blank, since only date/time/value should
  // genuinely differ quarter to quarter for an otherwise-unchanged monitoring point.
  if (!result._memoryApplied) {
    const memory = DataStore.getItemMemory(project.id, catKey);
    const siteHistory = DataStore.getSiteItemHistory(project.id, catKey);
    const methodDiffsByItem = {};
    result.rows.forEach(row => {
      const mem = memory[row[cat.itemField]];
      if (mem) {
        if (cat.methodField && !row[cat.methodField] && mem.method) {
          row[cat.methodField] = mem.method; row._methodFromMemory = true;
        } else if (cat.methodField && row[cat.methodField] && mem.method && row[cat.methodField] !== mem.method) {
          methodDiffsByItem[row[cat.itemField]] = { reportMethod: row[cat.methodField], memoryMethod: mem.method };
        }
        if (cat.unitField && !row[cat.unitField] && mem.unitCode) {
          row[cat.unitField] = mem.unitCode; row._unitFromMemory = true;
        }
      }
      const loc = row[cat.locationField];
      if (loc) {
        const histEntry = (siteHistory[loc] || {})[itemIdentityKey(row, cat)];
        if (histEntry && histEntry.snapshot) {
          Object.entries(histEntry.snapshot).forEach(([k, v]) => { if (v && !row[k]) row[k] = v; });
        }
      }
      // No history to fall back on doesn't have to mean typing the unit code by
      // hand — if the parsed unit field holds readable text (e.g. "mg/L") rather
      // than a valid code (some report parsers, e.g. air quality, don't always
      // resolve this themselves), try reverse-matching it against the unit code
      // table. A code that's already valid is left untouched.
      if (cat.unitField && row[cat.unitField] && !UNIT_CODES[row[cat.unitField]]) {
        const lookup = SmartParse.reverseUnitLookup(row[cat.unitField], row[cat.itemField]);
        if (lookup.code) {
          row[cat.unitField] = lookup.code;
          if (!lookup.confident) row._uncertainUnit = true;
        }
      }
    });
    result._methodDiffs = Object.entries(methodDiffsByItem).map(([item, d]) => ({ item, ...d }));
    result._memoryApplied = true;
  }
  if (!result._rowUidsAssigned) {
    result.rows.forEach((row, i) => { row._rowUid = i; });
    result._rowUidsAssigned = true;
  }

  const siteEntries = Object.entries(result.sites); // [key, {siteCode, rawLocation, rowIndices}]
  state.itemSelection = null; // reset so renderItemChecklist re-seeds with "all checked"
  state.excludedRowIndices = new Set(); // reset row-level exclusions for a fresh parse
  state.statSelection = null; // reset so the statistic picker re-seeds with its default

  // Rows whose 日期(起) is blank start UNTICKED. A real quarterly workbook routinely
  // carries leftover sheets from an earlier round that the lab never deleted — same
  // site names, same layout, different (older) numbers, and no 監測日期 filled in.
  // Those parse perfectly happily and would land in the filing next to this
  // quarter's real data with an empty date. Requiring an explicit tick to include
  // them keeps the data visible and recoverable without letting it in by accident.
  const datelessUids = result.rows.filter(r => !r['日期(起)']).map(r => r._rowUid);
  datelessUids.forEach(uid => state.excludedRowIndices.add(uid));

  renderPeriodPicker('smartPeriodWrap', result.rows);

  const updateCounts = () => {
    const selectedTotal = filterRowsBySelection(result.rows, cat.itemField).length;
    document.getElementById('smartImportSummary').textContent =
      `系統辨識出 ${siteEntries.length} 個測站、共 ${result.rows.length} 筆資料（來自 ${result.matchedSheets.length} 個工作表），目前已勾選 ${selectedTotal} 筆將匯入。`
      + ` 部分欄位無法從報告本身取得，請在下方為每個測站補充一次，之後匯入同一測站會自動套用。`;
    siteEntries.forEach(([key, site]) => {
      const n = filterRowsBySelection(site.rowIndices.map(i => result.rows[i]), cat.itemField).length;
      const cell = document.querySelector(`#smartSitesBody tr[data-site-key="${CSS.escape(key)}"] .site-row-count`);
      if (cell) cell.textContent = n;
    });
    document.getElementById('btnImportConfirm').textContent = `確認匯入 ${selectedTotal} 筆資料`;
  };
  const updateCountsAndRowDetail = () => {
    updateCounts();
    // The row-detail table itself is what excludedRowIndices comes from — show
    // rows filtered only by statistic and item type (not by row exclusion),
    // otherwise a row the person unchecked would vanish from the list and they
    // could never re-check it.
    const statRows = filterRowsByStat(result.rows);
    const itemFilteredRows = state.itemSelection
      ? statRows.filter(r => state.itemSelection.has((r[cat.itemField] || '').trim() || '（未標示）'))
      : statRows;
    renderRowDetailTable(document.getElementById('smartImportRowDetailWrap'), itemFilteredRows, cat, updateCounts);
  };

  // The statistic picker comes FIRST: which statistic is chosen decides what the
  // item checklist below it even contains.
  const renderItemsForCurrentStat = () => {
    renderItemChecklist(document.getElementById('smartImportItemsWrap'),
      filterRowsByStat(result.rows), cat.itemField, updateCountsAndRowDetail);
    updateCountsAndRowDetail();
  };
  renderStatChecklist(document.getElementById('smartImportStatWrap'), result.rows, renderItemsForCurrentStat);
  renderItemsForCurrentStat();

  const existingSkippedWarning = document.getElementById('smartImportSkippedItemsWarning');
  if (existingSkippedWarning) existingSkippedWarning.remove();
  const skippedItems = [...new Set(result.rows.flatMap(r => r._skippedPlaceholderItems || []))];
  if (skippedItems.length > 0) {
    const warn = document.createElement('div');
    warn.id = 'smartImportSkippedItemsWarning';
    warn.className = 'warning';
    warn.innerHTML = `ℹ️ 偵測到報告 Excel 檔案裡以下欄位被設定為「隱藏」（欄寬為0），代表這個測站實際上沒有監測這些項目，報告只是沿用共用範本、保留欄位但不對外顯示，已自動略過未匯入：${escapeHtml(skippedItems.join('、'))}。如果您確認這個測站其實有測這些項目，請直接用「＋新增一筆」手動補上。`;
    document.getElementById('smartImportItemsWrap').after(warn);
  }

  const existingMethodWarning = document.getElementById('smartImportMethodWarning');
  if (existingMethodWarning) existingMethodWarning.remove();
  if (cat.methodField) {
    const missingMethodItems = [...new Set(result.rows.filter(r => !r[cat.methodField]).map(r => r[cat.itemField]))];
    if (missingMethodItems.length > 0) {
      const warn = document.createElement('div');
      warn.id = 'smartImportMethodWarning';
      warn.className = 'warning';
      warn.innerHTML = `⚠️ 以下項目報告中未提供檢測方法、系統也沒有先前記憶的資料，匯入後請至「🧪 檢測方法管理」補充：${escapeHtml(missingMethodItems.join('、'))}`;
      document.getElementById('smartImportItemsWrap').after(warn);
    }
  }

  const existingMethodDiffNotice = document.getElementById('smartImportMethodDiffNotice');
  if (existingMethodDiffNotice) existingMethodDiffNotice.remove();
  if (result._methodDiffs && result._methodDiffs.length > 0) {
    const notice = document.createElement('div');
    notice.id = 'smartImportMethodDiffNotice';
    notice.className = 'warning';
    notice.style.background = '#e8f0fe';
    notice.style.borderColor = '#a8c7fa';
    notice.innerHTML = `ℹ️ 以下項目本次報告使用的檢測方法跟先前記錄不同，系統已依「本次報告」為準（例如溶氧的電極法 NIEA W455 與碘定量法 NIEA W422 都是合法方法，不同季節實驗室可能採用不同方法）。若並非實驗室刻意更換方法，請確認是否為判讀誤差：<br>` +
      result._methodDiffs.map(d => `・${escapeHtml(d.item)}：本次「${escapeHtml(d.reportMethod)}」，先前記錄為「${escapeHtml(d.memoryMethod)}」`).join('<br>');
    document.getElementById('smartImportItemsWrap').after(notice);
  }

  // ---- rows with no sampling date --------------------------------------------
  const existingDatelessWarning = document.getElementById('smartImportDatelessWarning');
  if (existingDatelessWarning) existingDatelessWarning.remove();
  if (datelessUids.length > 0) {
    const sheets = [...new Set(result.rows.filter(r => !r['日期(起)']).map(r => r._sourceSheet).filter(Boolean))];
    const warn = document.createElement('div');
    warn.id = 'smartImportDatelessWarning';
    warn.className = 'warning warning-strong';
    warn.innerHTML = `📅 有 <strong>${datelessUids.length} 筆資料讀不到採樣日期</strong>（報告上的「監測日期／採樣日期」是空白的），`
      + `已<strong>預設不匯入</strong>。<br>`
      + `這通常代表該工作表是上一次監測留在檔案裡、忘了刪除的舊版本——內容看起來很像，但數值是舊的。`
      + `${sheets.length ? `<br>來源工作表：${escapeHtml(sheets.join('、'))}` : ''}`
      + `<br><span class="hint">若確定這些資料是本次的，請展開下方「📋 詳細資料列表」勾選回來，匯入後再自行補上日期。</span>`;
    document.getElementById('smartImportItemsWrap').after(warn);
  }

  // ---- unrecognized layout: say so loudly ----------------------------------
  // Rows carrying `_autoDetected` didn't come from a parser written for this report
  // template; they came from the layout-agnostic reader that hunts for the seven
  // required fields by their column headings. That is far better than importing
  // nothing at all, but it IS a guess — so the preview names exactly which of the
  // seven it found and which it couldn't, and asks for the values to be checked.
  const existingAutoWarning = document.getElementById('smartImportAutoDetectWarning');
  if (existingAutoWarning) existingAutoWarning.remove();
  const autoRows = result.rows.filter(r => r._autoDetected);
  if (autoRows.length > 0) {
    const info = autoRows.find(r => r._autoDetectInfo)?._autoDetectInfo || { found: [], missing: [], extra: [] };
    const warn = document.createElement('div');
    warn.id = 'smartImportAutoDetectWarning';
    warn.className = 'warning warning-strong';
    warn.innerHTML = `🔎 <strong>這份檔案不是系統內建認得的報告格式</strong>，共 ${autoRows.length} 筆資料是系統<strong>自行從欄位標題「猜」出來的</strong>，`
      + `請務必逐筆核對後再匯入。<br>`
      + `・已自動找到：${info.found.length ? escapeHtml(info.found.join('、')) : '（無）'}`
      + `${info.extra && info.extra.length ? `（另外還找到：${escapeHtml(info.extra.join('、'))}）` : ''}<br>`
      + `・<strong>找不到、需要您自行補上：${info.missing.length ? escapeHtml(info.missing.join('、')) : '（無，七個必要欄位都找到了）'}</strong><br>`
      + `<span class="hint">小技巧：先匯入上一季同一個測站的資料，系統就會把座標、檢測方法、單位代碼等不會逐季改變的欄位自動沿用過來，這次只需確認日期／時間／數值。</span>`;
    document.getElementById('smartImportItemsWrap').after(warn);
  }

  // ---- 檢測類別 left blank (地質 always starts blank by design) --------------
  const existingCategoryWarning = document.getElementById('smartImportCategoryWarning');
  if (existingCategoryWarning) existingCategoryWarning.remove();
  if (cat.fields.some(f => f.key === '檢測類別') && result.rows.some(r => !r['檢測類別'])) {
    const warn = document.createElement('div');
    warn.id = 'smartImportCategoryWarning';
    warn.className = 'warning';
    warn.innerHTML = `⚠️ 有資料的「檢測類別」還是空白（這個欄位是申報必填）。請在下方測站設定表的「檢測類別」欄位選擇，`
      + `選好後會套用到該測站的所有資料；日後在表格中修改單筆時，系統也會詢問是否同步到<strong>同一份檔案、同一測站、同樣日期(起)(迄)</strong>的其他資料。`;
    document.getElementById('smartImportItemsWrap').after(warn);
  }

  const existingPdfWarning = document.getElementById('smartImportPdfWarning');
  if (existingPdfWarning) existingPdfWarning.remove();
  if (result.hasPdfSource) {
    const pdfWarn = document.createElement('div');
    pdfWarn.id = 'smartImportPdfWarning';
    pdfWarn.className = 'warning';
    pdfWarn.innerHTML = '⚠️ 這批資料包含從 PDF 判讀而來的內容。PDF 是逐行文字重組出來的表格，準確度低於 Excel（欄位對齊、數值可能有誤判），匯入後請務必逐筆核對，尤其是數值與座標欄位。';
    document.getElementById('smartImportItemsWrap').after(pdfWarn);
  }

  const uncertainCount = result.rows.filter(r => r._uncertainUnit).length;
  const existingWarning = document.getElementById('smartImportUnitWarning');
  if (existingWarning) existingWarning.remove();
  if (uncertainCount > 0) {
    const warn = document.createElement('div');
    warn.id = 'smartImportUnitWarning';
    warn.className = 'warning';
    const items = [...new Set(result.rows.filter(r => r._uncertainUnit).map(r => r[cat.itemField]))];
    warn.innerHTML = `⚠️ 有 ${uncertainCount} 筆資料（${escapeHtml(items.join('、'))}）的單位代碼是系統自動比對、非完全確定，匯入後請至「單位代碼表」核對並視需要手動修正。 <button type="button" class="btn btn-ghost btn-sm" id="btnOpenUnitRefFromWarning">開啟單位代碼表</button>`;
    document.getElementById('smartImportItemsWrap').after(warn);
    document.getElementById('btnOpenUnitRefFromWarning').addEventListener('click', openUnitRefModal);
  }

  const locField = cat.locationField;

  // Compare each LOCATION's items in THIS import against the project's site-item
  // history (accumulated from every past confirmed import for this category) — if
  // history has items this batch doesn't, offer to add them as blank rows (every
  // field carried over from the last confirmed snapshot except date/time/value,
  // which the person fills in themselves — the source is often a PDF this app
  // can't extract numbers from at all).
  //
  // Grouped by CONFIRMED location name, not raw site key: a category like noise
  // can have several distinct raw "sites" (e.g. a 環境噪音 report and a separate
  // 振動 report) that all resolve to the same physical location once the person
  // confirms the official site name — comparing history against just one of those
  // site keys' items would wrongly "miss" every item that actually belongs to the
  // other sub-report at the same location.
  //
  // Comparison and history are both keyed by itemIdentityKey (item name + 監測時段
  // when the category has that field), not bare item name — otherwise noise's
  // separate 日/晚/夜 rows for the same 音源發聲特性 collapse into one indistinguishable
  // entry and only one of the three ever gets suggested/rebuilt.
  const existingMissingWrap = document.getElementById('smartImportMissingItemsWrap');
  const siteHistory = DataStore.getSiteItemHistory(project.id, catKey);
  const itemMemoryForSuggestion = DataStore.getItemMemory(project.id, catKey);

  // Fuzzy location-name matching against history — catches the case where this
  // season's raw location text differs from a historically-known location only by
  // punctuation/whitespace/full-width-vs-half-width or a likely typo, which would
  // otherwise silently look like "this site vanished" + "a brand new site appeared"
  // instead of being recognized as the same physical site (a class of bug found
  // and fixed several times over the course of this project). Asked once per
  // parsed result (result._fuzzyLocationChecked), not on every checklist re-render,
  // and skipped for any site that already has a confirmed saved alias (nothing
  // ambiguous left to ask about there).
  if (!result._fuzzyLocationChecked) {
    const historicalLocs = Object.keys(siteHistory);
    const overrides = {};
    siteEntries.forEach(([key, site]) => {
      if (!site.rawLocation) return;
      const saved = savedAliases[siteAliasKey(site, result, cat)] || {};
      if (saved[locField]) return;
      const found = findSimilarHistoricalLocation(site.rawLocation, historicalLocs);
      if (!found) return;
      const useHistorical = confirm(
        `本次報告的測站「${site.rawLocation}」，跟歷史記錄裡的「${found.match}」名稱很相似（可能是標點符號、空格不同，或打字有誤）。\n\n` +
        `是否要視為同一個測站，改用歷史記錄的名稱「${found.match}」？（選「確定」才能正確沿用這個測站上一季的座標、管制標準等記憶內容）\n\n` +
        `若這其實是不同的測站，請選「取消」，維持原名稱「${site.rawLocation}」另建監測地點。`
      );
      if (useHistorical) overrides[key] = found.match;
    });
    result._fuzzyLocationOverrides = overrides;
    result._fuzzyLocationChecked = true;
  }
  const fuzzyLocationOverrides = result._fuzzyLocationOverrides || {};

  const itemsByLoc = {}; // confirmedLoc -> Map(identityKey -> count)
  const siteKeysByLoc = {}; // confirmedLoc -> [siteKey, ...] (may span multiple raw sites)
  siteEntries.forEach(([key, site]) => {
    const saved = savedAliases[siteAliasKey(site, result, cat)] || {};
    const confirmedLoc = saved[locField] || fuzzyLocationOverrides[key] || site.rawLocation;
    if (!itemsByLoc[confirmedLoc]) { itemsByLoc[confirmedLoc] = new Map(); siteKeysByLoc[confirmedLoc] = []; }
    site.rowIndices.forEach(i => {
      const idKey = itemIdentityKey(result.rows[i], cat);
      const m = itemsByLoc[confirmedLoc];
      m.set(idKey, (m.get(idKey) || 0) + 1);
    });
    siteKeysByLoc[confirmedLoc].push(key);
  });

  // tolerate the older, narrower history formats (array of item names, or
  // { item: category }) by simply treating them as having nothing usable —
  // there's no snapshot to rebuild from anyway, so nothing to suggest from them.
  const historyEntriesFor = (historyForLoc) => {
    if (!historyForLoc || Array.isArray(historyForLoc)) return [];
    return Object.entries(historyForLoc).filter(([, v]) => v && typeof v === 'object' && v.snapshot);
  };

  const suggestions = [];
  Object.entries(itemsByLoc).forEach(([confirmedLoc, currentCounts]) => {
    // Compare COUNTS, not just presence — a site sampled both 平日 and 假日 in the
    // same quarter legitimately has 2 rows sharing the same item+監測時段 identity;
    // if history remembers 2 but this import only has 1 (or 0), that's 1 (or 2)
    // still missing, not "already covered because at least one exists".
    const missing = historyEntriesFor(siteHistory[confirmedLoc])
      .map(([identityKey, entry]) => {
        const historicalCount = entry.count || 1; // tolerate pre-count history entries
        const missingCount = historicalCount - (currentCounts.get(identityKey) || 0);
        return missingCount > 0 ? { identityKey, ...entry, missingCount } : null;
      })
      .filter(Boolean);
    if (missing.length > 0) {
      suggestions.push({ siteKeys: siteKeysByLoc[confirmedLoc], location: confirmedLoc, missingItems: missing });
    }
  });
  // A whole location can be entirely absent from this import — e.g. this quarter's
  // report simply doesn't cover site A or B at all, even though they were sampled
  // last quarter. Those locations never show up in itemsByLoc above (there's no
  // current row for them at all), so they need a separate pass over the full
  // history to be offered — with no row from this import to copy shared fields
  // from, buildSuggestedRows rebuilds the whole row from the remembered snapshot
  // and leaves date/time/value blank for the person to fill in themselves. Every
  // historically-remembered occurrence (count) is offered, not just one.
  Object.entries(siteHistory).forEach(([histLoc, historyForLoc]) => {
    if (itemsByLoc[histLoc]) return; // already handled above — this location DID appear
    const historicalEntries = historyEntriesFor(historyForLoc);
    if (historicalEntries.length === 0) return;
    suggestions.push({
      siteKeys: [], location: histLoc, entirelyAbsent: true,
      missingItems: historicalEntries.map(([identityKey, entry]) => ({ identityKey, ...entry, missingCount: entry.count || 1 })),
    });
  });
  state.missingItemSuggestions = suggestions;

  if (suggestions.length === 0) {
    existingMissingWrap.innerHTML = '';
  } else {
    existingMissingWrap.innerHTML = `
      <div class="warning" style="background:#e8f0fe;border-color:#a8c7fa;">
        📋 系統比對過去記錄，以下測站過去曾出現、但本次報告未出現的測項。若要一併新增（檢測方法／單位／項目名稱及對應的檢測類別會依過去記錄先幫您填好，檢測數值請自行輸入），請勾選：
        <details style="margin-top:6px">
          <summary style="cursor:pointer;font-size:12.5px;color:var(--text-muted)">ℹ️ 這是跟哪一份資料比對的？（點此展開說明）</summary>
          <p class="hint" style="margin:6px 0 0 0">
            系統一律以「您最近一次確認匯入的內容」為比對基準，不是依季別標籤強制對應到「上一季」——只是一般情況下季度是照順序匯入，所以看起來就像是跟前一季比對。<br>
            如果您想指定要跟哪一份資料比對（例如想跳過中間某幾季，直接參考更早以前的完整版資料），只要<strong>先把那份資料匯入並確認</strong>，之後才匯入這次真正要處理的資料，系統就會以最後確認的那份為準。
          </p>
        </details>
        <div style="display:flex;gap:8px;align-items:center;margin:8px 0;flex-wrap:wrap">
          <input type="text" id="missingItemsSearchInput" placeholder="🔍 搜尋測站名稱，快速篩選" style="flex:1;min-width:180px;max-width:320px;padding:6px 9px;border:1px solid var(--border);border-radius:5px;font-family:inherit;font-size:13px">
          <button type="button" id="btnMissingSelectAllVisible" class="btn btn-ghost btn-sm">全選（僅目前顯示的測站）</button>
          <button type="button" id="btnMissingClearAllVisible" class="btn btn-ghost btn-sm">全不選（僅目前顯示的測站）</button>
        </div>
        <div id="missingItemsList" style="margin-top:8px">
          ${suggestions.map((s, i) => {
            // "Entirely absent" locations default to UNCHECKED — unlike "this site
            // is still here but missing a few items" (safe, low-risk, same-site
            // correction), a whole missing site more often means the person is
            // importing an unrelated batch/project's data as "this quarter" and the
            // old site genuinely doesn't belong there anymore. Bulk-importing every
            // vanished site by default risked silently mixing unrelated data across
            // batches; requiring an explicit opt-in here is the safer default.
            const defaultChecked = !s.entirelyAbsent;
            return `
            <div class="missing-item-loc-group" data-search-text="${escapeAttr(s.location.toLowerCase())}" style="margin-top:6px${s.entirelyAbsent ? ';padding:6px 8px;background:#fff6e0;border-radius:6px' : ''}">
              <label style="font-weight:700">
                <input type="checkbox" class="missing-item-group" data-group-idx="${i}" ${defaultChecked ? 'checked' : ''}> ${escapeHtml(s.location)}${s.entirelyAbsent ? ' <span class="hint">（本次報告完全沒有這個測站——若這是不同專案/不相關的批次資料，請保持不勾選；確定要帶回本測站的舊資料才勾選。建議新增的資料日期需要您自行填寫）</span>' : ''}
              </label>
              <div style="margin-left:22px">
                ${s.missingItems.map(({ identityKey, itemName, timeSegment, category, snapshot, missingCount }) => {
                  const displayName = timeSegment ? `${itemName}（${timeSegment}）` : itemName;
                  const countNote = missingCount > 1 ? `<strong>（缺 ${missingCount} 筆，例如平日／假日各一次，將新增 ${missingCount} 筆空白列）</strong>` : '';
                  const mem = itemMemoryForSuggestion[itemName];
                  const methodNote = snapshot?.[cat.methodField] || mem?.method;
                  const unitNote = snapshot?.[cat.unitField] || mem?.unitCode;
                  const memParts = [category ? `檢測類別：${category}` : '', methodNote, unitNote ? `單位代碼${unitNote}` : ''].filter(Boolean);
                  const memNote = memParts.length ? `（已記憶：${memParts.join('，')}）` : '（無先前記憶的方法/單位，需另外補上）';
                  return `<label style="display:block">
                    <input type="checkbox" class="missing-item-single" data-group-idx="${i}" data-identity-key="${escapeAttr(identityKey)}" ${defaultChecked ? 'checked' : ''}>
                    ${escapeHtml(displayName)} ${countNote} <span class="hint">${memNote}</span>
                  </label>`;
                }).join('')}
              </div>
            </div>
          `;
          }).join('')}
        </div>
      </div>
    `;
    document.querySelectorAll('.missing-item-group').forEach(cb => {
      cb.addEventListener('change', () => {
        const idx = cb.dataset.groupIdx;
        document.querySelectorAll(`.missing-item-single[data-group-idx="${idx}"]`).forEach(sub => { sub.checked = cb.checked; });
      });
    });
    const missingSearchInput = document.getElementById('missingItemsSearchInput');
    missingSearchInput.addEventListener('input', () => {
      const q = missingSearchInput.value.trim().toLowerCase();
      document.querySelectorAll('.missing-item-loc-group').forEach(div => {
        div.classList.toggle('row-hidden', !(!q || div.dataset.searchText.includes(q)));
      });
    });
    // Bulk-toggle every group checkbox that's currently visible under the search
    // filter — dispatches a real 'change' event so the existing group→sub-item sync
    // listener above fires normally, keeping the two layers consistent.
    const toggleVisibleGroups = (checked) => {
      document.querySelectorAll('.missing-item-loc-group').forEach(div => {
        if (div.classList.contains('row-hidden')) return;
        const groupCb = div.querySelector('.missing-item-group');
        if (groupCb.checked !== checked) {
          groupCb.checked = checked;
          groupCb.dispatchEvent(new Event('change', { bubbles: true }));
        }
      });
    };
    document.getElementById('btnMissingSelectAllVisible').addEventListener('click', () => toggleVisibleGroups(true));
    document.getElementById('btnMissingClearAllVisible').addEventListener('click', () => toggleVisibleGroups(false));
  }

  const wrap = document.getElementById('smartImportSitesWrap');
  wrap.innerHTML = `<table class="mapping-table">
    <thead><tr>
      <th>報告測點編號 / 原始記錄</th>
      ${profileFields.map(f => `<th>${escapeHtml(f.label)}</th>`).join('')}
      <th>筆數</th>
      <th>記住此設定</th>
    </tr></thead>
    <tbody id="smartSitesBody">
      ${siteEntries.map(([key, site]) => {
        const saved = savedAliases[siteAliasKey(site, result, cat)] || {};
        const firstRow = result.rows[site.rowIndices[0]] || {};
        const defaults = { [locField]: site.rawLocation, ...firstRow };
        if (fuzzyLocationOverrides[key]) defaults[locField] = fuzzyLocationOverrides[key];
        return `<tr data-site-key="${escapeAttr(key)}">
          <td>${escapeHtml(site.siteCode || site.rawLocation || key)}${site.siteCode && site.rawLocation ? `<br><span class="hint">${escapeHtml(site.rawLocation)}</span>` : ''}</td>
          ${profileFields.map(f => {
            // A blank remembered value must NOT win over a non-blank default — e.g.
            // the very first time this site was ever confirmed, a field like
            // coordinates may have been blank (the raw report simply doesn't have
            // them yet) and got saved as an empty alias; a later import can then
            // have that same field correctly auto-filled from the site-item history
            // snapshot (see _memoryApplied) into `firstRow`/`defaults` — that filled
            // value has to take priority, or the stale "remembered blank" would
            // silently blank it out again the moment the person hits confirm.
            const savedVal = saved[f.key];
            const val = (savedVal !== undefined && savedVal !== '') ? savedVal : (defaults[f.key] || '');
            if (f.type === 'select') {
              const opts = f.options.map(o => {
                const label = (f.optionLabels && f.optionLabels[o]) || o || '（未選擇）';
                return `<option value="${escapeAttr(o)}" ${o === val ? 'selected' : ''}>${escapeHtml(label)}</option>`;
              }).join('');
              return `<td><select data-site-field="${f.key}">${opts}</select></td>`;
            }
            return `<td><input type="text" data-site-field="${f.key}" value="${escapeAttr(val)}"></td>`;
          }).join('')}
          <td class="site-row-count">${site.rowIndices.length}</td>
          <td style="text-align:center"><input type="checkbox" data-site-remember checked></td>
        </tr>`;
      }).join('')}
    </tbody>
  </table>`;

  document.getElementById('smartImportSkipped').textContent =
    result.skippedSheets.length ? `略過 ${result.skippedSheets.length} 個無法辨識的工作表：${result.skippedSheets.join('、')}` : '';

  document.getElementById('importStep1').classList.add('hidden');
  document.getElementById('importStep2').classList.add('hidden');
  document.getElementById('importStep3').classList.remove('hidden');
  const confirmBtn = document.getElementById('btnImportConfirm');
  confirmBtn.classList.remove('hidden');
  updateCounts();
}

// ---------- duplicate detection ----------
// ---------- cross-season item memory (unit code / test method, remembered per project) ----------
function learnItemMemoryFromRows(projectId, catKey, cat, rows) {
  if (!cat.methodField && !cat.unitField) return;
  const updates = {};
  rows.forEach(r => {
    const itemName = r[cat.itemField];
    if (!itemName) return;
    const fields = {};
    if (cat.methodField && r[cat.methodField]) fields.method = r[cat.methodField];
    if (cat.unitField && r[cat.unitField]) fields.unitCode = r[cat.unitField];
    if (Object.keys(fields).length) updates[itemName] = { ...(updates[itemName] || {}), ...fields };
  });
  if (Object.keys(updates).length) DataStore.updateItemMemory(projectId, catKey, updates);
}

/** Records which items have ever been confirmed for each location — this is what
 *  lets a later import notice "上次這個測站還有 def 三項" even though this quarter's
 *  raw report only covers abc. Called after every successful import commit. */
/** Whether item identity for cross-season comparison needs to fold in 監測時段
 *  (day/evening/night) — currently only noise reuses the same item name (音源發聲
 *  特性, e.g. "均能音量(Leq)") across three separate rows distinguished only by
 *  their 監測時段, so without this the three would collapse into one indistinguishable
 *  entry both in "what items does this site currently have" and in history. */
function hasTimeSegmentField(cat) {
  return cat.fields.some(f => f.key === '監測時段');
}
/** A stable identity for one "kind of measurement" at a site — item name, plus
 *  監測時段 when the category has that field (see hasTimeSegmentField). */
function itemIdentityKey(row, cat) {
  return hasTimeSegmentField(cat) ? `${row[cat.itemField]}::${row['監測時段'] || ''}` : row[cat.itemField];
}
// 比較關係 is DERIVED from 檢測數值/監測數值 (see deriveComparisonRelation below),
// not an independent site/method attribute — it must never be historically
// inherited the way 座標/管制標準/檢測方法 are. Confirmed as a real bug: a site
// whose value was ND last season (比較關係="ND") got a real number this season,
// but the history-fill mechanism kept re-stamping 比較關係 back to "ND" because it
// was blank on the freshly-parsed row and got silently filled from the old
// snapshot — producing an internally-inconsistent row (a real number "10" tagged
// as 比較關係="ND") that would be wrong on an official filing.
/** Recomputes 比較關係 (comparison relation) from a value the person just typed
 *  into 檢測數值/監測數值, mirroring the exact rule SmartParse.parseValueCell uses
 *  when reading the same kind of value out of a lab report: a plain number ->
 *  比較關係 blank; "<0.3" or ">100" -> 比較關係 '<' or '>' (the value field keeps
 *  just the number, not the symbol); "ND" -> 比較關係 'ND'. Always returns a
 *  result (never null/undefined) — clearing the value to blank must also clear a
 *  stale comparison symbol left over from before, not leave one behind. */
function deriveComparisonRelation(rawValue) {
  const s = String(rawValue ?? '').trim();
  if (s === '') return { cmp: '', val: '', note: '' };
  // ND: 比較關係="ND", 檢測數值 stays BLANK — same rule as SmartParse.parseValueCell.
  if (/^ND$/i.test(s)) return { cmp: 'ND', val: '', note: '' };
  // NA/未檢測: neither 比較關係 nor 檢測數值 carries this — it belongs in 備註.
  if (/^(NA|未檢測|N\.A\.?)$/i.test(s)) return { cmp: '', val: '', note: '未檢測' };
  const m = s.match(/^([<>])\s*([\d.]+)$/);
  if (m) return { cmp: m[1], val: m[2], note: '' };
  return { cmp: '', val: s, note: '' };
}
const SNAPSHOT_EXCLUDED_FIELDS = new Set(['日期(起)', '時間(起)', '日期(迄)', '時間(迄)', '檢測數值', '監測數值', '比較關係']);
/** Every field's value except the row's own location/item identity and its actual
 *  measurement (date/time/value) — this is what gets carried forward wholesale when
 *  reconstructing a historically-known-but-currently-absent row, so a person only
 *  has to fill in this quarter's actual date/time/reading, not re-type coordinates,
 *  method, unit, 管制標準, remarks, etc. all over again. */
function buildFieldSnapshot(row, cat) {
  const exclude = new Set([cat.itemField, cat.locationField, ...SNAPSHOT_EXCLUDED_FIELDS]);
  const snap = {};
  cat.fields.forEach(f => { if (!exclude.has(f.key)) snap[f.key] = row[f.key] || ''; });
  return snap;
}

function learnSiteItemHistory(projectId, catKey, cat, rows) {
  const locField = cat.locationField, itemField = cat.itemField;
  // Recompute from the FULL current DataStore contents for whichever locations were
  // touched — not from `rows` directly — so a partial call (e.g. commit() passing
  // just the one row that was edited, or coordinate-manager passing only the rows
  // it changed) can't corrupt the remembered COUNT for a site that legitimately has
  // multiple rows sharing the same item identity (e.g. 平日/假日 pairs both landing
  // on "均能音量(Leq)::日間"). `rows` only tells us WHICH locations to re-learn;
  // the actual counting always looks at everything currently in DataStore for
  // those locations, which is the single source of truth.
  const touchedLocations = new Set(rows.filter(r => r[locField]).map(r => r[locField]));
  if (touchedLocations.size === 0) return;
  const allRows = DataStore.getData(projectId, catKey);
  const relevantRows = allRows.filter(r => touchedLocations.has(r[locField]) && r[itemField]);
  const entries = relevantRows.map(r => ({
    location: r[locField],
    identityKey: itemIdentityKey(r, cat),
    itemName: r[itemField],
    timeSegment: hasTimeSegmentField(cat) ? (r['監測時段'] || '') : '',
    itemCategory: r['檢測類別'] || '',
    snapshot: buildFieldSnapshot(r, cat),
  }));
  DataStore.learnSiteItemSnapshots(projectId, catKey, entries);
}

/**
 * Compares candidate rows against what's already in this category, split into:
 *  - brandNew: rows with no existing match at all — nothing to ask about
 *  - conflicts: rows that match an existing row's date/time/location/item exactly,
 *    but differ in some other field — this is the "you re-imported a corrected
 *    version" case, and needs a decision (update vs keep existing), never a silent
 *    guess either way
 * Rows that match an existing row on EVERY field (true duplicates) are silently
 * dropped — there's nothing to gain by asking about them.
 */
function analyzeImportAgainstExisting(existingRows, candidateRows, cat) {
  const locField = cat.locationField;
  const itemField = cat.itemField;
  const keyOf = (r) => [r['日期(起)'], r['時間(起)'], r[locField], r[itemField]].join('\u0001');
  const identityKeys = new Set(['日期(起)', '時間(起)', locField, itemField]);

  const existingByKey = new Map();
  existingRows.forEach((r, idx) => { if (!existingByKey.has(keyOf(r))) existingByKey.set(keyOf(r), idx); });

  const brandNew = [];
  const conflicts = [];

  candidateRows.forEach(candidate => {
    const key = keyOf(candidate);
    const existingIdx = existingByKey.get(key);
    if (existingIdx === undefined) { brandNew.push(candidate); return; }
    const existingRow = existingRows[existingIdx];
    const diffFields = [];
    cat.fields.forEach(f => {
      if (identityKeys.has(f.key)) return;
      const oldVal = existingRow[f.key] || '';
      const newVal = candidate[f.key] || '';
      if (oldVal !== newVal) diffFields.push({ key: f.key, label: f.label, type: f.type, oldVal, newVal });
    });
    if (diffFields.length === 0) return; // truly identical — silently skip, nothing to decide
    conflicts.push({
      existingIndex: existingIdx, candidateRow: candidate, diffFields,
      location: candidate[locField], item: candidate[itemField], date: candidate['日期(起)'],
    });
  });

  return { brandNew, conflicts };
}

/** Shows the conflict list and waits for the person's per-row decision, then calls
 *  onResolve({ useNew: Set<conflictIndex> }) — or doesn't call it at all if they
 *  cancel the whole import. */
function openConflictResolutionModal(conflicts, onResolve) {
  const wrap = document.getElementById('conflictListWrap');
  wrap.innerHTML = conflicts.map((c, i) => `
    <div class="conflict-item">
      <div class="conflict-item-header">
        <label>
          <input type="checkbox" class="conflict-use-new" data-idx="${i}" checked>
          <span><strong>${escapeHtml(c.location)}</strong>／${escapeHtml(c.item)}／${escapeHtml(toDateDisplayValue(c.date))}：套用本次匯入的新版本（取消勾選＝保留原有資料，不更動）</span>
        </label>
      </div>
      <table class="conflict-diff-table">
        <thead><tr><th>欄位</th><th>原有資料</th><th>本次匯入</th></tr></thead>
        <tbody>
          ${c.diffFields.map(d => {
            // show date/time differences the way the person reads them
            // (2026/06/25, 08:30) rather than in the internal storage format
            const fmt = (v) => d.type === 'date' ? (toDateDisplayValue(v) || v)
              : d.type === 'time' ? (toTimeDisplayValue(v) || v) : v;
            return `<tr>
            <td>${escapeHtml(d.label)}</td>
            <td class="diff-old">${escapeHtml(fmt(d.oldVal) || '（空白）')}</td>
            <td class="diff-new">${escapeHtml(fmt(d.newVal) || '（空白）')}</td>
          </tr>`;
          }).join('')}
        </tbody>
      </table>
    </div>
  `).join('');

  document.getElementById('btnConflictUseAllNew').onclick = () => {
    wrap.querySelectorAll('.conflict-use-new').forEach(cb => { cb.checked = true; });
  };
  document.getElementById('btnConflictKeepAllOld').onclick = () => {
    wrap.querySelectorAll('.conflict-use-new').forEach(cb => { cb.checked = false; });
  };
  document.getElementById('btnConflictCancel').onclick = () => {
    document.getElementById('conflictModal').classList.add('hidden');
  };
  document.getElementById('btnConflictConfirm').onclick = () => {
    const useNew = new Set();
    wrap.querySelectorAll('.conflict-use-new:checked').forEach(cb => useNew.add(Number(cb.dataset.idx)));
    document.getElementById('conflictModal').classList.add('hidden');
    onResolve({ useNew });
  };

  document.getElementById('conflictModal').classList.remove('hidden');
}

/** Synthesizes blank rows for any "missing item" suggestions the person checked —
 *  copies date/time/coordinates/category etc. from an existing row at that site (so
 *  it lines up as the same sampling event), fills in method/unit from item memory
 *  when available, and leaves the actual measurement blank for manual entry. */
function buildSuggestedRows(project, catKey, cat, result) {
  // Deliberately re-check the GROUP checkbox here too, rather than trusting only the
  // individual item checkbox's own .checked state — the group→items sync happens via
  // a 'change' listener on the group box, so if that ever doesn't fire for some
  // reason (or a re-render raced with the person's click), an unchecked group whose
  // items still show .checked=true in the DOM would otherwise get imported anyway.
  // Requiring BOTH layers to agree is what "the person unchecked this location" is
  // supposed to mean, and costs nothing when everything is working normally.
  const checkedBoxes = [...document.querySelectorAll('.missing-item-single:checked')].filter(cb => {
    const groupCb = document.querySelector(`.missing-item-group[data-group-idx="${cb.dataset.groupIdx}"]`);
    return !groupCb || groupCb.checked;
  });
  if (checkedBoxes.length === 0) return [];
  const itemMemory = DataStore.getItemMemory(project.id, catKey);

  // Build each checked item's single-row template first (without expanding
  // copiesToAdd yet), then interleave by OCCURRENCE index below rather than
  // finishing one item's copies before moving to the next.
  const perItemRows = []; // { template, copiesToAdd }
  checkedBoxes.forEach(cb => {
    const suggestion = state.missingItemSuggestions?.[Number(cb.dataset.groupIdx)];
    if (!suggestion) return;
    const identityKey = cb.dataset.identityKey;
    const missingEntry = suggestion.missingItems.find(m => m.identityKey === identityKey);
    if (!missingEntry) return;
    const { itemName, timeSegment, category: historicalCategory, snapshot, missingCount } = missingEntry;
    const copiesToAdd = missingCount || 1;

    // A location can span several raw "sites" (e.g. a noise sub-report and a
    // separate vibration sub-report at the same physical site) — gather rows from
    // ALL of them, then prefer one whose 檢測類別 already matches what history says
    // this item belongs to, so shared fields like 管制標準/頻率範圍 come from the
    // right kind of report rather than an arbitrary other category's row.
    const candidateRows = [];
    (suggestion.siteKeys || []).forEach(sk => {
      const site = result.sites[sk];
      if (site) site.rowIndices.forEach(i => candidateRows.push(result.rows[i]));
    });

    let newRow;
    if (candidateRows.length > 0) {
      // Partially-present location: this quarter's own data at the site is the
      // most accurate source for shared fields (coordinates etc, which could
      // genuinely have changed since last season) — use it as the base, then just
      // correct the identity fields to the specific missing item being added.
      const template = candidateRows.find(r => !historicalCategory || r['檢測類別'] === historicalCategory) || candidateRows[0];
      newRow = { ...template };
    } else {
      // Entirely-absent location (suggestion.entirelyAbsent): this quarter's report
      // has no rows for this site at all, so there's no current data to base shared
      // fields on — reconstruct the ENTIRE row from the remembered snapshot (every
      // field it had last season: coordinates, 管制標準, 檢測方法, 單位, 備註, etc),
      // not just a narrow "site profile" subset.
      newRow = {};
      cat.fields.forEach(f => { newRow[f.key] = ''; });
      newRow[cat.locationField] = suggestion.location;
      Object.entries(snapshot || {}).forEach(([k, v]) => { if (v) newRow[k] = v; });
    }
    newRow[cat.itemField] = itemName;
    if (timeSegment && '監測時段' in newRow) newRow['監測時段'] = timeSegment;
    if (historicalCategory && '檢測類別' in newRow) newRow['檢測類別'] = historicalCategory;
    // Only date/time/value are left blank for manual entry — everything else
    // (comparison relation, detection limit, method, unit, coordinates, remarks...)
    // carries over as-is, per the person's explicit request.
    ['日期(起)', '時間(起)', '日期(迄)', '時間(迄)', '檢測數值', '監測數值'].forEach(k => { if (k in newRow) newRow[k] = ''; });
    const mem = itemMemory[itemName];
    if (mem) {
      if (cat.methodField && !newRow[cat.methodField] && mem.method) newRow[cat.methodField] = mem.method;
      if (cat.unitField && !newRow[cat.unitField] && mem.unitCode) newRow[cat.unitField] = mem.unitCode;
    }
    perItemRows.push({ template: newRow, copiesToAdd });
  });

  // A site sampled both 平日 and 假日 historically appears GROUPED BY OCCURRENCE
  // — the weekday's full 日/晚/夜 set together, then the holiday's full set — not
  // grouped by time segment (both 日間 copies together, then both 晚間 copies).
  // Interleaving by occurrence index here reproduces that original grouping, so a
  // person filling in dates afterward sees each sampling occasion's rows already
  // sitting together instead of scattered by measurement type.
  const maxCopies = perItemRows.reduce((max, p) => Math.max(max, p.copiesToAdd), 0);
  const newRows = [];
  for (let occurrence = 0; occurrence < maxCopies; occurrence++) {
    perItemRows.forEach(({ template, copiesToAdd }) => {
      if (occurrence < copiesToAdd) newRows.push({ ...template });
    });
  }
  return newRows;
}

/**
 * Commits a resolved import (after any conflict decisions are made): applies
 * updates to matching existing rows in place, appends brand-new rows, tags
 * everything with batch/period metadata, records history/memory, and finishes the
 * modal/batch-queue flow. Shared by both the smart-parse and generic import paths.
 *  - brandNewRows: rows with no existing match — appended as new rows
 *  - updates: [{ existingIndex, newData }] — existing rows to overwrite in place
 *  - assignBatchId(row): returns the batch id a given row belongs to
 */
function finalizeImportCommit(project, catKey, cat, brandNewRows, updates, assignBatchId, importMode) {
  const periodLabel = (state.importPeriod || '').trim();
  const cleanify = (r) => {
    const out = { _batchId: assignBatchId(r), _period: periodLabel };
    cat.fields.forEach(f => { out[f.key] = r[f.key] || ''; });
    return out;
  };

  const existing = DataStore.getData(project.id, catKey);
  const cleanNew = brandNewRows.map(cleanify);
  const cleanUpdates = updates.map(u => ({ existingIndex: u.existingIndex, row: cleanify(u.newData) }));

  cleanUpdates.forEach(u => { existing[u.existingIndex] = u.row; });
  const merged = existing.concat(cleanNew);
  DataStore.saveData(project.id, catKey, merged);

  const allTouchedRows = cleanNew.concat(cleanUpdates.map(u => u.row));
  learnItemMemoryFromRows(project.id, catKey, cat, allTouchedRows);
  learnSiteItemHistory(project.id, catKey, cat, allTouchedRows);

  // One import-history entry per batch id, so "刪除此批次" in the history list acts
  // on exactly the rows this confirm action touched (new or updated).
  const rowCountByBatchId = {};
  const fileByBatchId = {};
  allTouchedRows.forEach(r => {
    rowCountByBatchId[r._batchId] = (rowCountByBatchId[r._batchId] || 0) + 1;
    if (!fileByBatchId[r._batchId]) fileByBatchId[r._batchId] = state.currentImportSourceLabel || '（未知來源）';
  });
  Object.entries(rowCountByBatchId).forEach(([id, rowCount]) => {
    DataStore.addImportBatch(project.id, catKey, {
      id, timestamp: new Date().toISOString(), sourceLabel: fileByBatchId[id],
      mode: importMode, rowCount, period: periodLabel,
    });
  });

  const summary = [];
  if (cleanNew.length) summary.push(`新增 ${cleanNew.length} 筆`);
  if (cleanUpdates.length) summary.push(`更新 ${cleanUpdates.length} 筆既有資料`);

  if (state.batchQueue && state.batchQueue.length > 0) {
    state.batchQueue.shift();
    renderContent();
    if (state.batchQueue.length > 0) {
      processNextBatchItem();
    } else {
      state.batchQueue = null;
      document.getElementById('importModal').classList.add('hidden');
      alert(`批次匯入完成，共匯入 ${state.batchQueueTotal} 個類別的資料，請至各分類頁面核對內容。`);
    }
    return;
  }

  closeImportModal();
  renderContent();
  alert(`已匯入「${cat.label}」：${summary.join('、') || '沒有變更'}。請於表格中核對內容是否正確，特別是尚未有測站設定的欄位。`);
}

function confirmSmartImport() {
  const project = getImportProject();
  if (!project) { alert('這次匯入鎖定的計畫已經不存在（可能已被刪除），請重新開啟匯入視窗。'); closeImportModal(); return; }
  const catKey = state.importCatKey;
  const cat = CATEGORIES[catKey];
  const result = state.smartResult;
  const profileFields = SmartParse.SITE_PROFILE_FIELDS[catKey] || [];
  const savedAliases = DataStore.getSiteAliases(project.id, catKey);

  // collect edited values per site from the DOM
  document.querySelectorAll('#smartSitesBody tr').forEach(tr => {
    const siteKey = tr.dataset.siteKey;
    const site = result.sites[siteKey];
    const overrides = {};
    tr.querySelectorAll('[data-site-field]').forEach(el => {
      overrides[el.dataset.siteField] = el.value;
    });
    const remember = tr.querySelector('[data-site-remember]').checked;
    const aliasKey = siteAliasKey(site, result, cat);
    if (remember) savedAliases[aliasKey] = overrides;
    else delete savedAliases[aliasKey];

    // apply overrides to every row belonging to this site
    site.rowIndices.forEach(idx => {
      Object.assign(result.rows[idx], overrides);
    });
  });
  DataStore.saveSiteAliases(project.id, catKey, savedAliases);

  let selectedRows = filterRowsBySelection(result.rows, cat.itemField);
  const suggestedRows = buildSuggestedRows(project, catKey, cat, result);
  selectedRows = selectedRows.concat(suggestedRows);
  if (selectedRows.length === 0) { alert('目前沒有勾選任何監測項目，請至少勾選一項再匯入。'); return; }

  // Tag rows with a batch id PER SOURCE FILE, not one shared id for the whole confirm
  // action — a multi-file import (e.g. three months' reports for the same site
  // selected together) must not let the "same file, same site" sync logic treat all
  // three months as one sampling event just because they were imported in one go.
  const batchIdByFile = {};
  const assignBatchId = (r) => {
    const key = r._sourceFile || '(單一檔案)';
    if (!batchIdByFile[key]) batchIdByFile[key] = 'b_' + Date.now() + '_' + Math.random().toString(36).slice(2, 8);
    return batchIdByFile[key];
  };

  const existing = DataStore.getData(project.id, catKey);
  const { brandNew, conflicts } = analyzeImportAgainstExisting(existing, selectedRows, cat);

  const proceed = (useNewSet) => {
    const updates = [];
    conflicts.forEach((c, i) => {
      if (!useNewSet || useNewSet.has(i)) updates.push({ existingIndex: c.existingIndex, newData: c.candidateRow });
    });
    if (brandNew.length === 0 && updates.length === 0) { alert('這批資料跟現有資料完全相同，或您選擇全部保留原有資料，沒有任何變更。'); return; }
    finalizeImportCommit(project, catKey, cat, brandNew, updates, assignBatchId, 'smart');
  };

  if (conflicts.length === 0) {
    if (brandNew.length === 0) { alert('這批資料跟現有資料完全相同，沒有新增或需要處理的內容。'); return; }
    proceed(null);
  } else {
    openConflictResolutionModal(conflicts, ({ useNew }) => proceed(useNew));
  }
}

function confirmImport() {
  if (state.importMode === 'smart') { confirmSmartImport(); return; }

  const project = getImportProject();
  if (!project) { alert('這次匯入鎖定的計畫已經不存在（可能已被刪除），請重新開啟匯入視窗。'); closeImportModal(); return; }
  const catKey = state.importCatKey;
  const cat = CATEGORIES[catKey];
  const mapping = {};
  document.querySelectorAll('#mappingTableBody select').forEach(sel => {
    mapping[sel.dataset.targetField] = sel.value || null;
  });

  const newRows = ImportEngine.applyMapping(state.importParsed.rows, mapping, cat.fields);
  newRows.forEach((r, i) => { r._rowUid = i; });
  // Same history-based fill as the smart-parse path: method/unit from item memory,
  // plus every OTHER blank field (座標/管制標準/管制區/環境音量標準/備註 etc) from
  // the last confirmed full-row snapshot for this exact (location, item-identity)
  // combo — this path (re-importing an already-formatted, schema-matching file, e.g.
  // a prior season's completed export or the government template) is the MAIN way
  // people actually re-import a full season's data, so it needs the same treatment
  // as the raw-report smart-parse path, not just the older method/unit-only fill.
  const memory = DataStore.getItemMemory(project.id, catKey);
  const siteHistoryForFill = DataStore.getSiteItemHistory(project.id, catKey);
  newRows.forEach(row => {
    const mem = memory[row[cat.itemField]];
    if (mem) {
      if (cat.methodField && !row[cat.methodField] && mem.method) row[cat.methodField] = mem.method;
      if (cat.unitField && !row[cat.unitField] && mem.unitCode) row[cat.unitField] = mem.unitCode;
    }
    const loc = row[cat.locationField];
    if (loc) {
      const histEntry = (siteHistoryForFill[loc] || {})[itemIdentityKey(row, cat)];
      if (histEntry && histEntry.snapshot) {
        Object.entries(histEntry.snapshot).forEach(([k, v]) => { if (v && !row[k]) row[k] = v; });
      }
    }
    // No history to fall back on (brand-new project, or this item's first
    // appearance) doesn't have to mean the person types the unit code by hand —
    // if the imported file's own 單位 column holds readable text (e.g. "mg/L")
    // rather than the code the schema expects, try reverse-matching it against the
    // unit code table, the same lookup the report-form parsers already use for
    // noise/water. A code that's already valid is left untouched.
    if (cat.unitField && row[cat.unitField] && !UNIT_CODES[row[cat.unitField]]) {
      const lookup = SmartParse.reverseUnitLookup(row[cat.unitField], row[cat.itemField]);
      if (lookup.code) {
        row[cat.unitField] = lookup.code;
        if (!lookup.confident) row._uncertainUnit = true;
      }
    }
  });
  const selectedRows = filterRowsBySelection(newRows, cat.itemField);
  if (selectedRows.length === 0) { alert('目前沒有勾選任何監測項目，請至少勾選一項再匯入。'); return; }

  const batchId = 'b_' + Date.now() + '_' + Math.random().toString(36).slice(2, 8);
  const assignBatchId = () => batchId; // one file, one batch id — no per-file split needed here

  const existing = DataStore.getData(project.id, catKey);
  const { brandNew, conflicts } = analyzeImportAgainstExisting(existing, selectedRows, cat);

  const proceed = (useNewSet) => {
    const updates = [];
    conflicts.forEach((c, i) => {
      if (!useNewSet || useNewSet.has(i)) updates.push({ existingIndex: c.existingIndex, newData: c.candidateRow });
    });
    if (brandNew.length === 0 && updates.length === 0) { alert('這批資料跟現有資料完全相同，或您選擇全部保留原有資料，沒有任何變更。'); return; }
    finalizeImportCommit(project, catKey, cat, brandNew, updates, assignBatchId, 'generic');
  };

  if (conflicts.length === 0) {
    if (brandNew.length === 0) { alert('這批資料跟現有資料完全相同，沒有新增或需要處理的內容。'); return; }
    proceed(null);
  } else {
    openConflictResolutionModal(conflicts, ({ useNew }) => proceed(useNew));
  }
}

// ---------- backup export/import ----------
const BACKUP_REMINDER_THRESHOLD_DAYS = 14;
function renderBackupReminder() {
  const banner = document.getElementById('backupReminderBanner');
  if (!banner) return;
  const projects = DataStore.getProjects();
  if (projects.length === 0) { banner.innerHTML = ''; return; }
  const lastBackupAt = localStorage.getItem('envapp_lastBackupAt');
  const daysSince = lastBackupAt ? Math.floor((Date.now() - Number(lastBackupAt)) / 86400000) : null;
  if (daysSince !== null && daysSince < BACKUP_REMINDER_THRESHOLD_DAYS) { banner.innerHTML = ''; return; }
  // All data lives only in this browser's local storage — no server-side copy at
  // all — so clearing browser data, switching devices, or an incognito session can
  // permanently lose already-confirmed filing data. This is unrelated to the
  // cross-season comparison/memory feature (that's just a convenience — losing it
  // just means re-typing some fields by hand, nothing is actually destroyed); this
  // reminder is specifically about protecting the real submission data itself.
  const message = daysSince === null
    ? '尚未備份過'
    : `距離上次備份已 ${daysSince} 天`;
  banner.innerHTML = `
    <div class="warning" style="margin-top:10px;font-size:12px;line-height:1.6">
      💾 ${message}。所有資料只存在這個瀏覽器裡，沒有任何伺服器備份，清除瀏覽器資料或換裝置都會導致資料完全遺失，建議定期備份。
      <button class="btn btn-primary btn-sm" id="btnBackupReminderExport" style="margin-top:6px;width:100%">立即備份匯出</button>
    </div>
  `;
  document.getElementById('btnBackupReminderExport').addEventListener('click', backupExport);
}
function backupExport() {
  const data = DataStore.exportAll();
  if (data.projects.length === 0) { alert('目前尚無任何計畫資料可匯出。'); return; }
  const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  const ts = new Date().toISOString().slice(0, 10);
  a.href = url;
  a.download = `環評監測資料備份_${ts}.json`;
  a.click();
  URL.revokeObjectURL(url);
  localStorage.setItem('envapp_lastBackupAt', String(Date.now()));
  renderBackupReminder();
}

async function backupImport(file) {
  if (!file) return;
  try {
    const text = await file.text();
    const payload = JSON.parse(text);
    const mode = confirm('要用備份檔「取代」目前所有資料嗎？\n按「確定」＝取代全部；按「取消」＝合併（新增為額外計畫，不影響現有資料）。') ? 'replace' : 'merge';
    DataStore.importAll(payload, mode);
    renderProjectList();
    state.currentProjectId = null;
    renderContent();
    alert('備份匯入完成。');
  } catch (err) {
    console.error(err);
    alert('匯入備份檔時發生錯誤，請確認檔案格式是否正確。');
  }
}

// ---------- init ----------
function init() {
  renderVersionBadge();
  renderProjectList();
  renderContent();

  document.getElementById('versionBadge').addEventListener('click', openChangelogModal);
  document.getElementById('btnChangelogClose').addEventListener('click', () => document.getElementById('changelogModal').classList.add('hidden'));

  document.getElementById('btnNewProject').addEventListener('click', () => openProjectModal(null));
  document.getElementById('btnProjectCancel').addEventListener('click', closeProjectModal);
  document.getElementById('btnProjectSave').addEventListener('click', saveProjectModal);

  document.getElementById('btnBackupExport').addEventListener('click', backupExport);
  document.getElementById('btnBackupImport').addEventListener('click', () => document.getElementById('backupFileInput').click());
  document.getElementById('backupFileInput').addEventListener('change', (e) => backupImport(e.target.files[0]));

  const importStaging = wireFileStaging({
    dropzoneId: 'importDropzone', fileInputId: 'importFileInput', listId: 'importStagedList',
    confirmBtnId: 'btnImportStagedConfirm', onConfirm: (files) => handleImportFile(files),
  });
  state._importStaging = importStaging; // so openImportModal can reset it when the modal (re)opens
  document.getElementById('btnImportCancel').addEventListener('click', closeImportModal);
  document.getElementById('btnImportConfirm').addEventListener('click', confirmImport);

  document.getElementById('btnCoordCancel').addEventListener('click', closeCoordModal);
  document.getElementById('btnCoordSave').addEventListener('click', saveCoordModal);

  document.getElementById('btnMethodCancel').addEventListener('click', closeMethodModal);
  document.getElementById('btnMethodSave').addEventListener('click', saveMethodModal);

  document.getElementById('btnUnitCodeRef').addEventListener('click', openUnitRefModal);
  document.getElementById('btnUnitRefClose').addEventListener('click', () => document.getElementById('unitRefModal').classList.add('hidden'));
  document.getElementById('unitRefSearch').addEventListener('input', (e) => renderUnitRefTable(e.target.value));

  document.getElementById('btnAgencyCodeRef').addEventListener('click', openAgencyRefModal);
  document.getElementById('btnAgencyRefClose').addEventListener('click', () => document.getElementById('agencyRefModal').classList.add('hidden'));
  document.getElementById('agencyRefSearch').addEventListener('input', (e) => renderAgencyRefTable(e.target.value));

  renderEiasLink();
  loadSiteConfig(); // async — re-renders link/template button once config.json loads
  document.getElementById('btnEiasLinkEdit').addEventListener('click', openEiasLinkEditModal);
  document.getElementById('btnEiasLinkCancel').addEventListener('click', () => document.getElementById('eiasLinkEditModal').classList.add('hidden'));
  document.getElementById('btnEiasLinkSave').addEventListener('click', saveEiasLinkEdit);
  document.getElementById('btnEiasLinkReset').addEventListener('click', () => {
    // "Reset" clears the PERSONAL override and returns to the shared config.json
    // value — not a hardcoded constant — so if the site owner already updated the
    // shared link, resetting correctly lands on that, not on a stale built-in URL.
    // Takes effect immediately (not just in the input box) so no leftover personal
    // override lingers in localStorage if the person doesn't also click Save.
    localStorage.removeItem('envapp_eiasUrl');
    document.getElementById('eiasLinkInput').value = getSharedEiasUrl();
    renderEiasLink();
  });

  document.getElementById('btnTemplateDownload').addEventListener('click', openTemplateDownloadModal);
  document.getElementById('btnTemplateDownloadClose').addEventListener('click', () => document.getElementById('templateDownloadModal').classList.add('hidden'));

  document.getElementById('btnCustomItemAdd').addEventListener('click', () => {
    const catKey = state.commonItemsCatKey;
    const cat = CATEGORIES[catKey];
    const name = document.getElementById('customItemName').value.trim();
    const method = document.getElementById('customItemMethod').value.trim();
    const unitCode = document.getElementById('customItemUnit').value.trim();
    if (!name) { alert('請輸入測項名稱。'); return; }
    if (state.commonItemsEntries.some(e => e.itemName === name)) { alert('這個測項名稱已經在清單裡了，請直接勾選它，或使用不同的名稱。'); return; }
    state.commonItemsEntries.push({ itemName: name, variants: [{ method, unitCode }] });
    renderCommonItemsList(cat);
    const newIdx = state.commonItemsEntries.length - 1;
    const cb = document.querySelector(`.common-item-check[data-idx="${newIdx}"]`);
    if (cb) cb.checked = true;
    document.getElementById('customItemName').value = '';
    document.getElementById('customItemMethod').value = '';
    document.getElementById('customItemUnit').value = '';
    document.getElementById('customItemName').focus();
  });
  document.getElementById('btnCommonItemsCancel').addEventListener('click', () => document.getElementById('commonItemsModal').classList.add('hidden'));
  document.getElementById('btnCommonItemsApply').addEventListener('click', () => {
    const project = getCurrentProject();
    const catKey = state.commonItemsCatKey;
    const cat = CATEGORIES[catKey];
    const checkedBoxes = [...document.querySelectorAll('.common-item-check:checked')];
    if (checkedBoxes.length === 0) { alert('請至少勾選一個測項。'); return; }
    const quantity = Math.max(1, Math.min(50, Number(document.getElementById('commonItemsQuantity').value) || 1));

    pushUndoSnapshot(project.id, catKey, `常用測項新增（${checkedBoxes.length}項${quantity > 1 ? `×${quantity}份` : ''}）`);
    const rows = DataStore.getData(project.id, catKey);
    const memoryUpdates = {};
    const presetItems = []; // { itemName, method } — remembered as a quick-reapply preset after this

    // Build each checked item's chosen variant once, up front (not inside the
    // per-group loop), so picking a method doesn't get re-evaluated per group.
    const picks = checkedBoxes.map(cb => {
      const idx = Number(cb.dataset.idx);
      const entry = state.commonItemsEntries[idx];
      const methodSelect = document.querySelector(`.common-item-method-select[data-idx="${idx}"]`);
      const variant = methodSelect ? entry.variants[Number(methodSelect.value)] : entry.variants[0];
      return { entry, variant };
    });
    picks.forEach(({ entry, variant }) => {
      presetItems.push({ itemName: entry.itemName, method: variant.method || '' });
      if (variant.method || variant.unitCode) {
        memoryUpdates[entry.itemName] = { method: variant.method, unitCode: variant.unitCode };
      }
    });

    // "幾份" (quantity) creates N groups of rows — outer loop over groups, inner
    // loop over the checked items — so all items for group 1 land together, then
    // all items for group 2, etc. This matters for how the table reads afterward:
    // grouped-by-station means the person can fill in one 地點 for a contiguous
    // block of rows (or select that block and batch-edit 地點 in one go), rather
    // than having the same item's N copies scattered together and every OTHER
    // item's copies elsewhere, which is awkward to assign locations against.
    for (let g = 0; g < quantity; g++) {
      picks.forEach(({ entry, variant }) => {
        const blank = {};
        cat.fields.forEach(f => { blank[f.key] = ''; });
        blank[cat.itemField] = entry.itemName;
        if (cat.methodField && variant.method) blank[cat.methodField] = variant.method;
        if (cat.unitField && variant.unitCode) blank[cat.unitField] = variant.unitCode;
        rows.push(blank);
      });
    }
    DataStore.saveData(project.id, catKey, rows);
    if (Object.keys(memoryUpdates).length) DataStore.updateItemMemory(project.id, catKey, memoryUpdates);
    rememberItemPreset(catKey, presetItems);
    document.getElementById('commonItemsModal').classList.add('hidden');
    renderContent();
  });

  const batchStaging = wireFileStaging({
    dropzoneId: 'batchDropzone', fileInputId: 'batchFileInput', listId: 'batchStagedList',
    confirmBtnId: 'btnBatchStagedConfirm', onConfirm: (files) => handleBatchFiles(files),
  });
  state._batchStaging = batchStaging;
  document.getElementById('btnBatchCancel').addEventListener('click', closeBatchImportModal);

  document.getElementById('btnBatchHistoryClose').addEventListener('click', () => document.getElementById('batchHistoryModal').classList.add('hidden'));

  document.getElementById('btnExportSelectCancel').addEventListener('click', closeExportSelectModal);
  document.getElementById('btnExportSelectConfirm').addEventListener('click', confirmExportSelect);

  // Universal modal escape hatch: clicking the dark overlay outside the modal box,
  // or pressing Escape, closes whichever modal is currently open. This is a safety
  // net independent of each modal's own Cancel/Close button, so a modal can never
  // become a dead end.
  //
  // The import modal is deliberately EXCLUDED: importing is a multi-step flow where
  // the person may have already picked a period, adjusted site names/categories, or
  // checked/unchecked items — an accidental click just outside the modal (easy to do,
  // since the modal doesn't fill the screen) would silently discard all of that. Now
  // that the Cancel button itself is reliably visible, there's no longer a need for
  // this safety net there; closing the import modal requires the explicit Cancel
  // button (or finishing/cancelling a batch import via its own controls).
  document.querySelectorAll('.modal-overlay').forEach(overlay => {
    if (overlay.id === 'importModal') return;
    overlay.addEventListener('click', (e) => {
      if (e.target === overlay) overlay.classList.add('hidden');
    });
  });
  document.addEventListener('keydown', (e) => {
    if (e.key !== 'Escape') return;
    document.querySelectorAll('.modal-overlay').forEach(overlay => {
      if (overlay.id === 'importModal') return;
      if (!overlay.classList.contains('hidden')) overlay.classList.add('hidden');
    });
  });
}

document.addEventListener('DOMContentLoaded', init);
