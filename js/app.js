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
  rowSearch: {}, // { [catKey]: free-text search box contents } — kept in state so a
  // re-render (a sort click, an accepted sync, a column filter) can't silently clear
  // the search box while the ▾ filters still look active. That mismatch made
  // "select all visible + delete" act on far more rows than the person could see.
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
function pushUndoSnapshot(projectId, catKey, description, rowsOverride) {
  const key = undoKey(projectId, catKey);
  if (!state.undoStack[key]) state.undoStack[key] = [];
  /*
   * rowsOverride：呼叫端已經把資料改掉了，但它知道「改之前」長什麼樣。
   * 表格裡改一格是這種情況——值在每一次按鍵就寫進去了，等到失焦才知道
   * 這是一次真正的編輯，那時候只能靠 baseline 把那一格還原成舊值來拍快照。
   */
  const rows = rowsOverride || DataStore.getData(projectId, catKey);
  const batches = DataStore.getImportBatches(projectId, catKey);
  const presets = getItemPresets(catKey);
  state.undoStack[key].push({ description, rows: deepCopy(rows), batches: deepCopy(batches), presets: deepCopy(presets) });
  if (state.undoStack[key].length > UNDO_STACK_LIMIT) state.undoStack[key].shift();
  state.redoStack[key] = [];
}
/** 丟掉最上面那個復原點——動作最後發現「其實什麼都沒改」時用，
 *  免得復原清單裡多一個按了沒反應的項目。 */
function popUndoSnapshot(projectId, catKey) {
  const stack = state.undoStack[undoKey(projectId, catKey)];
  if (stack && stack.length > 0) stack.pop();
}
/** 事後補上更精確的說明——例如「修改『座標』」在同步之後變成
 *  「修改『座標』並同步 4 筆」。一次使用者動作只留一個復原點。 */
function retitleUndoSnapshot(projectId, catKey, description) {
  const stack = state.undoStack[undoKey(projectId, catKey)];
  if (stack && stack.length > 0) stack[stack.length - 1].description = description;
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

/*
 * 採樣日期不屬於本次期別的資料列，預設不勾選（v4.39）
 * ──────────────────────────────────────────────────
 * 為什麼要這一條：選期別**不會**排除異常日期的資料，它只是把整批資料標上同一個
 * 期別標籤。實測使用者的 13545NV7：128 列裡有 6 列是 113年第4季（2024 年）的
 * 舊資料——審查報告的人沒有刪掉的舊工作表——選了「115年第1季」之後，
 * 那 6 列的 _period 也一起變成「115年第1季」，就這樣跟著送出去。
 *
 * 舊版只有一句提醒，要使用者自己到「詳細資料列表」把它們取消勾選。
 * 現在改成**預設就不勾**：資料還在、列表上看得到、隨時可以勾回來，
 * 但不會因為沒看到那句提醒就默默送出去。和「沒有日期的列預設不勾」同一個原則。
 *
 * ⚠️ 只動「這一輪自動排除的」那些 uid（記在 autoExcludedByPeriod）。
 * 使用者自己動過的勾選永遠優先——他把某一列勾回來之後，
 * 就算之後又改了期別，也不會再被自動取消勾選。
 */
function applyPeriodExclusions(rows) {
  if (!state.excludedRowIndices) state.excludedRowIndices = new Set();
  if (!state.autoExcludedByPeriod) state.autoExcludedByPeriod = new Set();
  if (!state.manualRowChoices) state.manualRowChoices = new Set();
  // 先把上一輪自動排除的還原，否則改期別之後舊的排除會留著
  state.autoExcludedByPeriod.forEach(uid => state.excludedRowIndices.delete(uid));
  state.autoExcludedByPeriod.clear();

  const target = normalizePeriodShorthand(state.importPeriod || '').trim();
  if (!target) return 0;
  rows.forEach((r) => {
    if (r._rowUid === undefined) return;
    /*
     * 使用者自己勾過的那幾列，這裡一律不碰。
     * 只從 autoExcludedByPeriod 移除是不夠的——下面重算時它照樣符合
     * 「期別不同」的條件，又會被排除一次，等於使用者的決定被無聲推翻。
     */
    if (state.manualRowChoices.has(r._rowUid)) return;
    const own = guessPeriodFromRows([r]);
    if (!own || own === target) return;   // 猜不出期別的不動（沒有日期的另有規則）
    state.excludedRowIndices.add(r._rowUid);
    state.autoExcludedByPeriod.add(r._rowUid);
  });
  return state.autoExcludedByPeriod.size;
}

function renderPeriodPicker(containerId, rows, onPeriodChange) {
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
  const target = normalizePeriodShorthand(state.importPeriod || '').trim();
  const offCount = periodEntries.filter(([p]) => p !== target).reduce((n, [, c]) => n + c, 0);
  const outlierWarning = periodEntries.length > 1
    ? `⚠️ 偵測到本次資料的採樣日期橫跨不同期別：${periodEntries.map(([p, c]) => `${escapeHtml(p)}（${c}筆）`).join('、')}。`
      + (offCount > 0
        ? `<br><strong>不屬於「${escapeHtml(target)}」的 ${offCount} 筆已經預設取消勾選、不會匯入</strong>（常見原因是報告裡夾帶了沒刪乾淨的舊工作表或參考值）。資料仍然列在下方「詳細資料列表」，確定要送的話展開勾回來即可。`
        : `如果其中有不屬於本次要送件季度的資料（例如報告裡夾帶的舊資料或參考值），建議到下方「詳細資料列表」展開後取消勾選排除，避免誤送。`)
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

  // 改了期別就要重算「哪些列不屬於這一期」，否則畫面上的排除是照舊期別算的。
  const notifyPeriodChanged = () => { if (onPeriodChange) onPeriodChange(); };

  const applyFromDropdown = () => {
    if (yearInput.value && quarterInput.value) {
      state.importPeriod = `${yearInput.value}年第${quarterInput.value}季`;
      textInput.value = state.importPeriod;
      notifyPeriodChanged();
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
    notifyPeriodChanged();
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
    // No Math.max(1, …) floor. With it, ANY two names differing by one character
    // were offered as "the same site typed differently" — and 大甲溪上游 vs 大甲溪下游,
    // 民生國小 vs 民權國小, 測站1 vs 測站2 are exactly how distinct monitoring points
    // are named. Accepting the prompt merged two physically different stations into
    // one in the filing. The ratio rule now governs, so a 1-edit match needs names of
    // at least 10 characters.
    // One allowed edit per 10 characters. A ratio of 0.2 still let 5-character names
    // match on a single edit, which is exactly 大甲溪上游 vs 大甲溪下游 — the pair this
    // guard exists to keep apart. Tier 1 above (punctuation/whitespace-insensitive
    // equality) still catches the genuinely-same-name-typed-differently case.
    const threshold = Math.floor(Math.max(rawLoc.length, h.length) / 10);
    if (threshold >= 1 && dist <= threshold && dist < bestDist) { best = h; bestDist = dist; }
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
  migrateNightVibrationItem(id);
  renderProjectList();
  renderContent();
}

// ── v4.30 一次性處理：夜間振動的 Lvd(10) → Lvn(10) ──────────────────────────
//
// 分成兩半，理由不同：
//
//  1. **內部帳本（跨季測站記憶、測項方法記憶）一律自動轉換，不問。**
//     那是使用者看不到的東西，而且不轉換一定會出事：舊記憶裡的
//     「Lvd(10)::夜間」在本版匯入時會被當成「這個測項不見了」，
//     而缺少測項的建議是預設打勾的——使用者不取消就會多出一筆空白列。
//
//  2. **看得見的資料列要先問過才改。** 那是使用者已經整理好、甚至已經申報出去的
//     內容，程式不該自己動手。每個計畫只問一次，回答「不要」就記住不再問。
const VIB_MIGRATION_FLAG_KEY = 'envapp_vibnight_migrated_v1';

function vibMigrationDone(projectId) {
  try {
    const done = JSON.parse(localStorage.getItem(VIB_MIGRATION_FLAG_KEY)) || {};
    return !!done[projectId];
  } catch (e) { return false; }
}
function markVibMigrationDone(projectId) {
  try {
    const done = JSON.parse(localStorage.getItem(VIB_MIGRATION_FLAG_KEY)) || {};
    done[projectId] = true;
    localStorage.setItem(VIB_MIGRATION_FLAG_KEY, JSON.stringify(done));
  } catch (e) { /* 記不住就下次再問一次，不影響資料 */ }
}

/** 內部帳本的轉換。沒有畫面、不需要同意，而且每次開啟計畫跑都是安全的（冪等）。 */
function migrateVibBookkeeping(projectId) {
  const cat = CATEGORIES.noise;
  // 跨季測站記憶：識別碼與記下來的測項名稱一起改。
  const history = DataStore.getSiteItemHistory(projectId, 'noise');
  let historyChanged = false;
  Object.keys(history).forEach(loc => {
    const entries = history[loc];
    if (!entries || typeof entries !== 'object' || Array.isArray(entries)) return;
    Object.keys(entries).forEach(key => {
      const entry = entries[key];
      if (!entry || typeof entry !== 'object') return;
      const seg = entry.timeSegment || '';
      const fixedName = canonicalVibItemName(entry.itemName, seg);
      /*
       * 只處理「夜間的 Lvd(10)」這一種，其餘一律原封不動。
       *
       * 第一版是先組出 `${fixedName}::${seg}` 再跟原本的 key 比對，結果把所有
       * 「key 的長相剛好不是這個格式」的舊資料也一起改寫、刪除——包含更早期
       * 沒有時段的記憶格式，以及跟本次修正毫無關係的測項。函式註解寫著只動一種
       * 組合，就該只動一種。
       */
      if (fixedName === entry.itemName) return;
      const fixedKey = `${fixedName}::${seg}`;
      const existing = entries[fixedKey];
      const merged = { ...entry, itemName: fixedName };
      if (existing) {
        /*
         * 兩個識別碼同時存在（使用者在舊版就手動改過一部分）。筆數取兩者的較大值、
         * 日期取聯集，**不相加**——相加會把同一次量測算成兩筆，下次匯入就會出現
         * 「缺 N 筆」的假警報。少算不會有人受害（頂多不提醒），多算會。
         */
        merged.count = Math.max(existing.count || 1, entry.count || 1);
        merged.dates = [...new Set([...(existing.dates || []), ...(entry.dates || [])])];
        merged.snapshot = entry.snapshot || existing.snapshot;
      }
      entries[fixedKey] = merged;
      delete entries[key];
      historyChanged = true;
    });
  });
  if (historyChanged) DataStore.saveSiteItemHistory(projectId, 'noise', history);

  // 測項方法記憶是用純測項名稱當 key（沒有時段），所以不能改名，
  // 而是把 Lvd(10) 記住的方法與單位「複製」一份給 Lvn(10)——兩者本來就同方法同單位。
  const memory = DataStore.getItemMemory(projectId, 'noise');
  if (memory[VIB_LV10_DAY] && !memory[VIB_LV10_NIGHT]) {
    DataStore.updateItemMemory(projectId, 'noise', { [VIB_LV10_NIGHT]: { ...memory[VIB_LV10_DAY] } });
  }
  return cat;
}

/** 找出所有「夜間 ＋ 振動 ＋ Lvd(10)」的資料列——依定義那就是錯的組合。 */
function findNightVibrationRowsToFix(projectId) {
  const cat = CATEGORIES.noise;
  const rows = DataStore.getData(projectId, 'noise');
  const hits = [];
  rows.forEach((r, idx) => {
    if (r['檢測類別'] === '振動' && r['監測時段'] === '夜間' && r[cat.itemField] === VIB_LV10_DAY) {
      hits.push({ idx, location: r[cat.locationField] || '（未填地點）', date: r['日期(起)'] || '' });
    }
  });
  return hits;
}

function migrateNightVibrationItem(projectId) {
  if (!projectId) return;
  migrateVibBookkeeping(projectId); // 無論如何都要做，而且冪等
  if (vibMigrationDone(projectId)) return;
  const hits = findNightVibrationRowsToFix(projectId);
  /*
   * 沒有要改的就直接離開，**不要記旗標**。
   *
   * 旗標的意思是「這個人說過不要」，不是「檢查過了」。第一版連「沒東西可改」
   * 也一併記起來，於是日後把上一季用舊版產生的完成版申報檔匯回來——那裡面帶著
   * 夜間的 Lvd(10)——就再也不會有人提醒他了。
   */
  if (hits.length === 0) return;

  const locs = [...new Set(hits.map(h => h.location))];
  const locText = locs.slice(0, 6).join('、') + (locs.length > 6 ? ` 等 ${locs.length} 個測站` : '');
  const ok = confirm(
    `【夜間振動的音源發聲特性】

`
    + `官方的音源發聲特性，環境振動分成日間 Lvd(10) 與夜間 Lvn(10) 兩種`
    + `（d 是 day、n 是 night）。舊版程式不分日夜，一律填 Lvd(10)，夜間那一筆是錯的。

`
    + `這個計畫裡有 ${hits.length} 筆「夜間 + 振動」的資料寫著 Lvd(10)：
`
    + `　${locText}

`
    + `要幫您改成 Lvn(10) 嗎？
`
    + `・只改「音源發聲特性」這一欄，日期、數值、地點等其他欄位一律不動。
`
    + `・日間的 Lvd(10)、營建振動的 Lveq／Lvmax 都不會被動到。
`
    + `・改之前建議先「匯出備份」；本程式沒有復原按鈕。

`
    + `按「取消」就維持原樣，之後不再詢問（您仍可在表格上自行修改）。`
  );
  if (!ok) { markVibMigrationDone(projectId); return; } // 只記住「不要」

  const cat = CATEGORIES.noise;
  const rows = DataStore.getData(projectId, 'noise');
  hits.forEach(h => { rows[h.idx][cat.itemField] = VIB_LV10_NIGHT; });
  DataStore.saveData(projectId, 'noise', rows);
  learnSiteItemHistory(projectId, 'noise', cat, rows);
  alert(`已將 ${hits.length} 筆夜間振動的音源發聲特性改為 ${VIB_LV10_NIGHT}。`);
}

// ---------- content area ----------
function renderContent() {
  const content = document.getElementById('content');
  const project = getCurrentProject();
  if (!project) {
    content.innerHTML = '<div class="empty-state"><p>👈 請先在左側建立或選擇一個計畫</p>'
      + '<p style="font-size:14px;margin-top:10px">第一次使用？請先看 '
      + '<a href="使用說明.html" target="_blank" rel="noopener" style="color:var(--primary-dark);font-weight:600">📘 新手使用說明</a></p></div>';
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
    // Blanks are handled OUTSIDE the direction flip. Negating the comparator flipped
    // the blank sentinels too, so a second click (descending) put every empty cell at
    // the TOP — the opposite of Excel, which this deliberately imitates, and exactly
    // wrong for the "sort descending to find the outliers" review pass.
    displayEntries.sort((a, b) => {
      const av = String(a.row[sortState.fieldKey] ?? '').trim();
      const bv = String(b.row[sortState.fieldKey] ?? '').trim();
      if (av === '' || bv === '') return av === bv ? 0 : (av === '' ? 1 : -1);
      return dir * compareForSort(av, bv);
    });
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
        <button class="btn btn-ghost btn-sm" id="btnDecimalSettings" title="設定各類別的數值欄位要顯示幾位小數，以及要用補零還是四捨五入。只影響畫面與匯出，不會改動存起來的原始值。">🔢 小數位數設定</button>
        <button class="btn btn-ghost btn-sm" id="btnBatchHistory">📜 匯入紀錄${batches.length ? ` (${batches.length})` : ''}</button>
        ${showPeriodUI ? `
        <select id="periodFilterSelect" title="篩選要查看／編輯／匯出哪一期的資料；「匯出此類別」會依此篩選範圍匯出">
          <option value="" ${activePeriod === '' ? 'selected' : ''}>全部期別</option>
          ${knownPeriods.map(p => `<option value="${escapeAttr(p)}" ${activePeriod === p ? 'selected' : ''}>${escapeHtml(p)}</option>`).join('')}
          ${hasUnlabeled ? `<option value="__none__" ${activePeriod === '__none__' ? 'selected' : ''}>未標示期別</option>` : ''}
        </select>` : ''}
        <input type="text" id="rowSearchInput" value="${escapeAttr(state.rowSearch[catKey] || '')}" placeholder="🔍 搜尋任何欄位內容（測站、測項、日期…）" title="輸入關鍵字即時篩選畫面顯示的資料列，方便檢查或修正特定資料；不影響匯出範圍（匯出仍依上方期別篩選）">
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
  document.getElementById('btnDecimalSettings').addEventListener('click', () => openDecimalSettingsModal(catKey));
  if (showPeriodUI) {
    document.getElementById('periodFilterSelect').addEventListener('change', (e) => {
      state.periodFilter[catKey] = e.target.value;
      renderContent();
    });
  }
  document.getElementById('btnExportCat').addEventListener('click', () => {
    // Re-read from storage at click time. `displayRows` is a snapshot taken when the
    // table was last rendered, and ordinary cell edits commit to storage WITHOUT
    // re-rendering — so exporting straight after fixing a value wrote the old value
    // to the file while the screen showed the new one.
    const freshRows = DataStore.getData(project.id, catKey).filter(r => {
      if (!activePeriod) return true;
      if (activePeriod === '__none__') return !r._period;
      return r._period === activePeriod;
    });
    if (freshRows.length === 0) { alert('目前篩選範圍內沒有資料可匯出。'); return; }
    if (!confirmNoiseVibRuleBeforeExport(freshRows, cat)) return;
    if (!confirmTimeRangeBeforeExport(freshRows, cat)) return;
    if (!confirmLimitBeforeExport(freshRows, cat)) return;
    if (!confirmRequiredBeforeExport(freshRows, cat)) return;
    ExportEngine.downloadCategory(project, DataStore.getBasicInfo(project.id), catKey, freshRows);
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
      renderContentPreservingScroll();
    });
  }
  const btnRedo = document.getElementById('btnRedo');
  if (btnRedo) {
    btnRedo.addEventListener('click', () => {
      const snapshot = peekRedo(project.id, catKey);
      if (!snapshot) return;
      if (!confirm(`確定要重做「${snapshot.description}」嗎？`)) return;
      popRedoAndRestore(project.id, catKey);
      renderContentPreservingScroll();
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
  if (state.rowSearch[catKey] === undefined) state.rowSearch[catKey] = '';
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
    state.rowSearch[catKey] = searchInput.value; // survives the next re-render
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
    state.rowSearch[catKey] = '';
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
      renderContentPreservingScroll();
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

  /*
   * 批次修改也要守噪音／振動的規矩（v4.41）
   * ────────────────────────────────────────
   * 表格裡改一格會問「要不要同步給振動」，而且管制標準、管制區這些
   * 官方規定必須不同的欄位連問都不問。但**批次修改是另一條路徑**，
   * 它完全繞過那套規則——使用者同時勾了噪音和振動、把「管制區」統一改成
   * 「第3類」，振動那幾筆就這樣變成不合規，而且沒有任何提示。
   *
   * 這裡不擋。使用者是自己一列一列勾的，他可能真的知道自己在做什麼
   * （例如只勾振動、統一改成「無」，那完全正確）。
   * 要擋的是「勾的範圍橫跨兩邊」這件事——那多半是誤選。
   */
  const selectedRows = DataStore.getData(project.id, catKey)
    .filter((_, i) => indices.includes(i));
  const noiseCount = catKey === 'noise' ? selectedRows.filter(r => noiseHalfOf(r) === 'noise').length : 0;
  const vibCount = catKey === 'noise' ? selectedRows.filter(r => noiseHalfOf(r) === 'vib').length : 0;
  const spansBothHalves = noiseCount > 0 && vibCount > 0;
  const OFFICIAL_REASON = {
    '管制標準': '官方規定振動必須填「無」、噪音不能填「無」（115 年 7 月起加強檢核）',
    '管制區': '官方規定振動必須填「無」、噪音不能填「無」（115 年 7 月起加強檢核）',
    '監測單位': '噪音是 dB(A)（代碼 16）、振動是 dB（代碼 159）',
    '監測方法': '噪音是 NIEA P201／P205、振動是 NIEA P204',
    '音源發聲特性': '噪音是 Leq／Lmax／Leq,LF、振動是 Lveq／Lvmax／Lvd(10)／Lvn(10)',
    '頻率範圍': '這是噪音的欄位，振動留空白',
    '環境音量標準': '這是道路交通噪音專用的欄位，振動不適用',
    '監測時段': '噪音分日／晚／夜三段，振動只有日／夜',
    '檢測類別': '這正是用來分辨噪音與振動的欄位',
  };
  const conflictReason = (fieldKey) =>
    (spansBothHalves && NOISE_VIB_DISTINCT_FIELDS.includes(fieldKey))
      ? (OFFICIAL_REASON[fieldKey] || '這個欄位在噪音與振動之間本來就不一樣')
      : '';

  const renderValueInput = () => {
    const field = editableFields.find(f => f.key === fieldSelect.value);
    document.getElementById('batchEditValueWrap').innerHTML = fieldControlHTML(field, '', '', undefined, catKey);
    const warnEl = document.getElementById('batchEditWarning');
    const reason = conflictReason(field.key);
    warnEl.innerHTML = reason
      ? `<div class="warning" style="margin:10px 0 0">⚠️ 您選取的資料<strong>同時包含噪音 ${noiseCount} 筆與振動 ${vibCount} 筆</strong>，`
        + `而<strong>「${escapeHtml(field.label)}」在噪音與振動之間本來就不一樣</strong>——${escapeHtml(reason)}。<br>`
        + `統一改成同一個值，其中一邊會變成不合規。<br>`
        + `<span class="hint">建議取消，改成分兩次做：先只勾噪音那 ${noiseCount} 筆改一次，再只勾振動那 ${vibCount} 筆改一次。</span></div>`
      : (spansBothHalves
        ? `<div class="warning" style="margin:10px 0 0;background:#e8f0fe;border-color:#a8c7fa">`
          + `ℹ️ 您選取的資料同時包含<strong>噪音 ${noiseCount} 筆</strong>與<strong>振動 ${vibCount} 筆</strong>，兩邊都會被改成同一個值。</div>`
        : '');
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

    const reason = conflictReason(field.key);
    if (reason) {
      // 已經在視窗上用紅框講過一次了，這裡是最後一道——要按兩次才過得去是刻意的
      if (!confirm(
        `⚠️ 您選取的資料同時包含噪音 ${noiseCount} 筆與振動 ${vibCount} 筆。\n\n`
        + `「${field.label}」在噪音與振動之間本來就不一樣——${reason}。\n`
        + `統一改成「${newValue || '（空白）'}」，其中一邊會變成不合規、上傳可能被退件。\n\n`
        + `建議按「取消」，改成分兩次做：先只勾噪音那 ${noiseCount} 筆，再只勾振動那 ${vibCount} 筆。\n\n`
        + `仍要兩邊一起改嗎？`
      )) return;
    } else if (spansBothHalves) {
      if (!confirm(
        `您選取的資料同時包含噪音 ${noiseCount} 筆與振動 ${vibCount} 筆，`
        + `兩邊的「${field.label}」都會被改成「${newValue || '（空白）'}」。\n\n`
        + `確定要這樣修改嗎？（可用「↩️ 復原上一步」救回）`
      )) return;
    } else if (!confirm(`確定要把已選取的 ${indices.length} 筆資料的「${field.label}」統一改成「${newValue || '（空白）'}」嗎？（可用「↩️ 復原上一步」救回）`)) return;

    pushUndoSnapshot(project.id, catKey, `批次修改「${field.label}」（${indices.length}筆）`);
    const rows = DataStore.getData(project.id, catKey);
    const touchedRows = [];
    indices.forEach(idx => {
      if (rows[idx]) { rows[idx][field.key] = newValue; touchedRows.push(rows[idx]); }
    });
    DataStore.saveData(project.id, catKey, rows);
    if (touchedRows.length > 0) learnSiteItemHistory(project.id, catKey, cat, touchedRows);
    modal.classList.add('hidden');
    renderContentPreservingScroll();
  };
}

// ---------- import batch history ----------
/*
 * 小數位數設定面板。
 *
 * 為什麼做成使用者可調，而不是把官方規定寫死在程式裡：規定會改。
 * 寫死的話規定一改就得改程式、重新打包、重新上傳；改成設定之後，
 * 使用者自己在畫面上改一改就好，而出廠預設仍然是目前官方辭典的規定。
 *
 * ⚠️ 這個設定是**整台電腦共用**的，不是每個計畫一份——官方規定全國一致，
 * 不會因計畫而異，每開一個新計畫就要重設一次反而容易漏。
 */
function openDecimalSettingsModal(currentCatKey) {
  const modal = document.getElementById('decimalModal');
  const wrap = document.getElementById('decimalSettingsWrap');
  let draft = readDecimalSettings();

  const render = () => {
    const cats = CATEGORY_ORDER.filter(k => (DECIMAL_CONFIGURABLE_FIELDS[k] || []).length);
    wrap.innerHTML = cats.map(catKey => {
      const cat = CATEGORIES[catKey];
      const rows = DECIMAL_CONFIGURABLE_FIELDS[catKey].map(fieldKey => {
        const rule = draft[decimalSettingKey(catKey, fieldKey)] || DECIMAL_OFF;
        const official = OFFICIAL_DECIMAL_RULES[decimalSettingKey(catKey, fieldKey)];
        const modeOpts = Object.entries(DECIMAL_MODES)
          .map(([v, label]) => `<option value="${v}" ${rule.mode === v ? 'selected' : ''}>${escapeHtml(label)}</option>`).join('');
        const digitOpts = Array.from({ length: DECIMAL_MAX_DIGITS + 1 }, (_, d) =>
          `<option value="${d}" ${rule.digits === d ? 'selected' : ''}>${d} 位</option>`).join('');
        // 用這個欄位真實存在的值當例子，比憑空舉 3.5 有說服力
        const sample = sampleValueFor(catKey, fieldKey) || '3.5';
        return `<tr>
          <td>${escapeHtml(fieldKey)}</td>
          <td><select data-dec-mode="${escapeAttr(decimalSettingKey(catKey, fieldKey))}">${modeOpts}</select></td>
          <td><select data-dec-digits="${escapeAttr(decimalSettingKey(catKey, fieldKey))}" ${rule.mode === 'off' ? 'disabled' : ''}>${digitOpts}</select></td>
          <td class="hint">${escapeHtml(sample)} → <strong>${escapeHtml(formatDecimal(sample, rule))}</strong></td>
          <td class="hint">${official ? `官方：${escapeHtml(DECIMAL_MODES[official.mode])} ${official.digits} 位` : '官方無規定'}</td>
        </tr>`;
      }).join('');
      return `<h4 style="margin:12px 0 6px">${escapeHtml(cat.label)}${catKey === currentCatKey ? '（目前這一類）' : ''}</h4>
        <table class="mapping-table"><thead><tr>
          <th>欄位</th><th>呈現方式</th><th>位數</th><th>範例</th><th>官方辭典</th>
        </tr></thead><tbody>${rows}</tbody></table>`;
    }).join('');

    wrap.querySelectorAll('[data-dec-mode]').forEach(sel => {
      sel.onchange = () => {
        const key = sel.dataset.decMode;
        draft[key] = { ...(draft[key] || DECIMAL_OFF), mode: sel.value };
        render();
      };
    });
    wrap.querySelectorAll('[data-dec-digits]').forEach(sel => {
      sel.onchange = () => {
        const key = sel.dataset.decDigits;
        draft[key] = { ...(draft[key] || DECIMAL_OFF), digits: Number(sel.value) };
        render();
      };
    });
  };

  render();
  modal.classList.remove('hidden');

  document.getElementById('btnDecimalCancel').onclick = () => modal.classList.add('hidden');
  document.getElementById('btnDecimalReset').onclick = () => {
    if (!confirm('要回復出廠預設嗎？\n\n只有「噪音－監測數值」會是補零 2 位，其餘欄位全部改成「不處理」（原樣顯示）。')) return;
    draft = { ...DEFAULT_DECIMAL_SETTINGS };
    render();
  };
  document.getElementById('btnDecimalOfficial').onclick = () => {
    if (!confirm('要套用 115 年版官方資料辭典的建議嗎？\n\n噪音的監測數值，以及空氣／水質／地質的採樣深度、採樣水深、採樣地點高度、污染物採樣高度，都會設成補零 2 位。\n辭典沒有規定的欄位（例如各類別的檢測數值）維持你目前的設定。')) return;
    draft = { ...draft, ...OFFICIAL_DECIMAL_RULES };
    render();
  };
  document.getElementById('btnDecimalSave').onclick = () => {
    writeDecimalSettings(draft);
    modal.classList.add('hidden');
    showToast('小數位數設定已儲存。存起來的原始值沒有被改動，只有顯示與匯出會照新設定。');
    renderContentPreservingScroll();
  };
}

/** 從目前資料裡找一個真實的值當範例，找不到就回 null。 */
function sampleValueFor(catKey, fieldKey) {
  const project = getCurrentProject();
  if (!project) return null;
  const rows = DataStore.getData(project.id, catKey) || [];
  for (const r of rows) {
    const v = String(r[fieldKey] ?? '').trim();
    if (v && /^[+-]?\d*\.?\d+$/.test(v)) return v;
  }
  return null;
}

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
        renderContentPreservingScroll();
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
  /*
   * 這一列少了哪些必填欄位，一次算好再傳給每一格——不要每一格各算一次，
   * 一列二十幾欄乘上幾百列會很慢。
   */
  const missingHere = new Map(missingRequiredFields(row, cat).map(m => [m.key, m.why]));
  /*
   * 噪音／振動的連動規則同樣一列算一次（v4.42）。
   * 這是「已經填錯」的檢查——資料可能是別家公司交來的、或上一季匯入的，
   * 不見得經過這個程式的任何一條編輯路徑，所以只能在畫出來的時候看。
   */
  const ruleHere = new Map(noiseVibRuleViolations(row, cat).map(v => [v.key, v.why]));
  const ctl = (f) => fieldControlHTML(f, row[f.key], `data-row="${idx}"`, missingHere.get(f.key), cat.key, ruleHere.get(f.key));
  const pinnedCells = displayFieldOrder(cat).slice(0, 2).map(f => `<td${f.key === cat.itemField ? ' class="col-item"' : f.key === cat.locationField ? ' class="col-loc"' : ''}>${ctl(f)}</td>`).join('');
  const restCells = displayFieldOrder(cat).slice(2).map(f => `<td>${ctl(f)}</td>`).join('');
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

/**
 * @param missingWhy 這一格「該填而沒填」時的原因文字；沒有就是 undefined。
 *   有值時整格畫紅框並附說明，讓使用者看得到要補哪裡——以前完全沒有提示，
 *   而「補回缺少測項」產生的空白列是預設打勾的，很容易就這樣送出去。
 */
function fieldControlHTML(field, value, rowAttr, missingWhy, catKey, ruleWhy) {
  value = value ?? '';
  const base = `${rowAttr} data-field="${field.key}"`;
  const missTip = missingWhy
    ? ` title="${escapeAttr(`「${field.label}」是${missingWhy}，目前是空的。`)}"`
    : '';
  const missCls = missingWhy ? ' cell-required-missing' : '';
  switch (field.type) {
    case 'select': {
      /*
       * 存進來的值有可能不在官方清單裡，而且是真的會發生：
       *   ・使用者自己的 115Q2 噪音申報檔，頻率範圍寫「20 Hz 至 200Hz」
       *     （官方是「20 Hz 至 200 Hz」，中間少一個空格）
       *   ・別家公司填好交來的生態檔，特有性寫「特有／特亞／外來」
       *     （官方只有「特有種」「特有亞種」）
       *
       * 舊寫法只把清單裡的值標 selected，於是這種值畫出來是**一片空白**——
       * 看的人以為欄位沒填。更糟的是那個 <select> 的目前值真的就是空字串，
       * 所以只要它收到一次 change（點開下拉再關掉就會），存起來的值就被清成空白，
       * 而且 頻率範圍 是必填欄位。實測：存的是「20 Hz 至 200Hz」→ 畫面空白 →
       * 碰一下 → 存的變成空白 → 匯出也是空白。整個過程沒有任何提示。
       *
       * 改成：清單外的值照樣列成一個選項並選起來（值不會消失），
       * 同時用紅框標示、滑鼠移上去說明為什麼，讓人看得見要修。
       */
      const value_ = String(value ?? '');
      const known = field.options.includes(value_);
      const opts = field.options.map(o => {
        const label = (field.optionLabels && field.optionLabels[o]) || o || '（未選擇）';
        return `<option value="${escapeAttr(o)}" ${o === value_ ? 'selected' : ''}>${escapeHtml(label)}</option>`;
      }).join('');
      const extra = known || value_ === ''
        ? ''
        : `<option value="${escapeAttr(value_)}" selected>${escapeHtml(value_)}（非標準值）</option>`;
      /*
       * 兩種紅框都可能出現在同一格：值不在官方清單裡（extra），
       * 以及值合法但**和檢測類別對不起來**（ruleWhy，例如振動填了「第2類」）。
       * 後者比較常見也比較難發現，所以說明文字優先講它。
       */
      const badValue = Boolean(extra) || Boolean(ruleWhy);
      const cls = badValue || missCls ? ` class="${badValue ? 'cell-invalid' : ''}${missCls}"` : '';
      const tip = ruleWhy
        ? ` title="${escapeAttr(`${ruleWhy}。這是官方明文規定的連動規則（115 年 7 月起加強檢核），不改申報會被退件。`)}"`
        : extra
          ? ` title="「${escapeAttr(value_)}」不在官方允許的選項裡（${escapeAttr(field.options.filter(Boolean).join('、'))}），申報時可能被退件。請從清單中改選正確的值。"`
          : (missTip || (field.help ? ` title="${escapeAttr(field.help)}"` : ''));
      return `<select ${base}${cls}${tip}>${opts}${extra}</select>`;
    }
    case 'date':
      // Plain text rather than a native <input type=date>: native date pickers render
      // according to browser/OS locale and can't be forced to show "YYYY/MM/DD" —
      // typing "2026/5/12" or "2026-5-12" both work, normalized on blur.
      return `<input type="text" ${base}${missTip} value="${escapeAttr(toDateDisplayValue(value))}" class="date-input${missCls}" placeholder="YYYY/MM/DD" inputmode="numeric" maxlength="10">`;
    case 'time':
      // Plain text rather than a native <input type=time>: native time pickers on many
      // devices show a scroll-wheel that's fiddly to land on an exact second, and on
      // some mobile browsers don't reliably fire change events at all. Typing "1430",
      // "14:30", or "14:30:00" all work — normalized to HH:MM on blur (the official
      // template's own time format has no seconds either).
    {
      /*
       * 官方限 00:00~23:59（勿填 24:00）。超出範圍的值**照樣顯示出來**、
       * 用紅框標示，不會被默默改掉——和上面 select 遇到清單外的值同一個原則：
       * 值不能消失，但要看得見它是錯的。
       */
      const badTime = DateTimeUtil.outOfRangeTimeReason(value);
      const cls = `time-input${badTime ? ' cell-invalid' : ''}${missCls}`;
      const tip = badTime
        ? ` title="${escapeAttr(`「${toTimeDisplayValue(value)}」${badTime}。請改成 23:59（一天的最後一分鐘）或正確的時間。`)}"`
        : missTip;
      return `<input type="text" ${base}${tip} value="${escapeAttr(toTimeDisplayValue(value))}" class="${cls}" placeholder="HH:MM" inputmode="numeric" maxlength="8">`;
    }
    case 'suggest': {
      /*
       * ⚠️ 這個 id 以前是 `suggest-` ＋ 欄位名稱去掉所有非英數字元。
       * 中文欄位名稱去掉之後**一個字都不剩**，所以「比較關係」與「調查項目」
       * 算出來的 id 都是 `suggest-`——兩個完全不同的建議清單會撞在一起，
       * 而且下面的 `if (!getElementById(listId))` 會讓先建立的那一個留著不換。
       * 實際症狀：先開空品（比較關係）再開生態，生態的「調查項目」跳出來的
       * 是 ＞、＜、ND。v4.35 之前只有「比較關係」一個 suggest 欄位，所以撞不到；
       * v4.36 生態的「調查項目」也改成 suggest，這個洞就會踩到。
       *
       * 改成用欄位名稱的字碼組出 id，中文也有唯一值。
       */
      const listId = `suggest-${[...field.key].map(ch => ch.charCodeAt(0).toString(36)).join('')}`;
      if (!document.getElementById(listId)) {
        const dl = document.createElement('datalist');
        dl.id = listId;
        dl.innerHTML = field.options.map(o => `<option value="${escapeAttr(o)}">`).join('');
        document.body.appendChild(dl);
      }
      const maxAttr = field.maxLength ? ` maxlength="${field.maxLength}"` : '';
      return `<input type="text" ${base}${missTip} value="${escapeAttr(value)}" list="${listId}"${maxAttr}${missCls ? ` class="${missCls.trim()}"` : ''}>`;
    }
    case 'unitcode':
      return `<input type="text" ${base} value="${escapeAttr(value)}" class="code-input${missCls}" data-codetype="unit" title="${escapeAttr(missingWhy ? `「${field.label}」是${missingWhy}，目前是空的。` : lookupUnit(value))}" placeholder="代碼">`;
    case 'agencycode':
      return `<input type="text" ${base} value="${escapeAttr(value)}" class="code-input${missCls}" data-codetype="agency" title="${escapeAttr(missingWhy ? `「${field.label}」是${missingWhy}，目前是空的。` : lookupAgency(value))}" placeholder="代碼">`;
    default: {
      /*
       * 依「小數位數設定」把值格式化後顯示。設定是使用者自己調的
       * （工具列的「🔢 小數位數設定」），出廠預設 ＝ 目前官方辭典的規定。
       *
       * ⚠️ 這裡改的只是**畫面上看到的字**，DataStore 裡存的永遠是原始值。
       * 所以設定改來改去都能還原，設錯也救得回來。
       * ⚠️ 一定要**連類別一起看**，不能只比對欄位名稱：官方哪天把水質的
       * 「檢測數值」改名成「監測數值」（噪音現在就叫這個），只看名稱就會
       * 把水質的值也套上噪音的規則。catKey 沒傳進來時一律不處理。
       */
      if (catKey) value = formatFieldValue(catKey, field.key, value);
      const bad = field.key === LIMIT_FIELD && String(value).trim() !== '' && !isPlainNumber(value);
      const cls = bad || missCls ? ` class="${bad ? 'cell-invalid' : ''}${missCls}"` : '';
      const tip = bad ? ' title="「檢測極限」只能填數值。請改成純數字，或清空這一格。"' : missTip;
      return `<input type="text" ${base}${cls}${tip} value="${escapeAttr(value)}">`;
    }
  }
}

/*
 * 同步選擇視窗（v4.37）
 * ─────────────────────
 * 以前所有的同步都是 confirm()——只有「確定／取消」兩個答案。
 * 噪音與振動要能分開選（只改這一筆／只同步給噪音／只同步給振動／兩邊都要），
 * 兩個答案表達不了，所以改用一個小視窗。
 *
 * 只有「真的有兩邊可選」的時候才會用它；其餘情況仍然走原本的 confirm()，
 * 免得單純的是非題也要多開一個視窗。
 *
 * 回傳被按下的那個選項的 key，關掉或按 Esc 回傳 null。
 */
/*
 * 同步視窗開著的時候，不可以再開第二個。
 * ────────────────────────────────────
 * 視窗一打開就會把焦點移到第一顆按鈕上——而焦點離開原本那一格，
 * 又會再觸發一次 focusout，於是同一次編輯開了兩個視窗：
 * 第二個把第一個的按鈕覆蓋掉，而且它已經沒有「改之前的值」可以還原
 * （baseline 在第一次就被清掉了），「取消修改」那顆就這樣消失。
 * 症狀正是使用者回報的「少了取消的功能」。
 */
let syncDialogOpen = false;

/*
 * `dismissKey`：按 Esc 或點視窗外面時，等同按了哪一個選項。
 *
 * 以前是 null（＝只改這一筆），但那樣不合直覺——使用者按 Esc 的意思是
 * 「算了，我不要改」，結果那一格卻默默留著新值，而且沒有任何提示。
 * 那正是這一版一直在避免的事：沒有講出來的資料變更。
 * 現在只要還原得回去，Esc 就等同「取消修改」。
 */
function askSyncChoice({ title, body, options, wide = false, dismissKey = null }) {
  return new Promise((resolve) => {
    syncDialogOpen = true;
    const modal = document.getElementById('syncChoiceModal');
    const actions = document.getElementById('syncChoiceActions');
    // 有列表的時候放寬視窗，免得「會怎麼改」那一欄擠成一直條
    modal.querySelector('.modal').classList.toggle('modal-wide', Boolean(wide));
    document.getElementById('syncChoiceTitle').textContent = title;
    document.getElementById('syncChoiceBody').innerHTML = body;
    actions.innerHTML = '';
    let done = false;
    const finish = (key) => {
      if (done) return;
      done = true;
      syncDialogOpen = false;
      document.removeEventListener('keydown', onKey);
      modal.removeEventListener('click', onBackdrop);
      modal.classList.add('hidden');
      actions.innerHTML = '';
      resolve(key);
    };
    const onKey = (e) => { if (e.key === 'Escape') finish(dismissKey); };
    // 點視窗外面的深色區域也算取消——和 Esc 同一個意思
    const onBackdrop = (e) => { if (e.target === modal) finish(dismissKey); };
    modal.addEventListener('click', onBackdrop);
    options.forEach((opt) => {
      const btn = document.createElement('button');
      btn.className = `btn ${opt.primary ? 'btn-primary' : 'btn-ghost'} btn-sm`;
      btn.textContent = opt.label;
      btn.addEventListener('click', () => finish(opt.key));
      actions.appendChild(btn);
    });
    document.addEventListener('keydown', onKey);
    modal.classList.remove('hidden');
    const first = actions.querySelector('button');
    if (first) first.focus();
  });
}

/** 噪音類別裡，一列屬於「噪音」還是「振動」那一半。 */
function noiseHalfOf(row) {
  return (row && row['檢測類別'] === '振動') ? 'vib' : 'noise';
}

function wireGridEvents(project, catKey, cat) {
  const tbody = document.getElementById('gridBody');
  // COORD_FIELDS 現在定義在 schema.js（v4.37）——匯入那一側也要用同一份清單。

  const commit = (rowIdx, fieldKey, value, { learn = true } = {}) => {
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
    /*
     * 監測數值補成兩位小數。
     *
     * 位置很重要：要在 deriveComparisonRelation 之後，因為那一段可能會把值換掉
     * （輸入「<40」會變成 比較關係="<" ＋ 監測數值="40"），先補零就會被它抵銷。
     * 而且只在**失焦時**（learn=true）做——每一次按鍵都補的話，使用者才打了「3」
     * 就被改成「3.00」，游標跳掉、也打不出 39.2。
     */
    /*
     * ⚠️ 這裡**刻意不把格式化後的值寫回去**。
     * v4.33 以前會寫回去（存的就是 39.20），那在「四捨五入」這個模式下會出事：
     * 設成兩位就把 0.125 永久變成 0.13，設定改回去也救不回原始值。
     * 現在存的永遠是使用者打的／報告上的原始值，格式化只發生在畫面與匯出。
     */
    DataStore.saveData(project.id, catKey, rows);
    // Keep the site-item history snapshot current with manual corrections too — not
    // just at import time — so a value the person fixes by hand (e.g. filling in
    // coordinates a water report never provides) is what gets carried forward next
    // season, not whatever was frozen in at the moment of import.
    //
    // But NOT on every keystroke. History is additive and never forgets, so learning
    // mid-typing filed every prefix of a name as its own remembered site: typing
    // 「大甲溪橋」 left permanent phantom sites 大, 大甲, 大甲溪, which then fed the
    // fuzzy-location prompt and the missing-item suggestions on the next import.
    // The `input` handler passes learn:false; the matching blur below does the
    // learning once, with the finished value.
    if (learn) learnSiteItemHistory(project.id, catKey, cat, [rows[rowIdx]]);
  };

  /**
   * WHAT COUNTS AS "THE SAME SAMPLING EVENT" FOR A SYNC
   * ---------------------------------------------------
   * One physical site can hold several distinct measurement sets at once, and they
   * must not bleed into each other:
   *
   *   - 平日 and 假日 are separate visits to the SAME site, told apart only by their
   *     sampling dates. Correcting the 假日 date must never reach the 平日 rows.
   *   - 噪音 and 振動 are separate sub-reports of the same visit, told apart by
   *     檢測類別. They share the things that come from STANDING IN ONE SPOT AT ONE
   *     MOMENT — 座標 (one tripod, one position) and 日期／時間起迄 (the same visit) —
   *     but nothing else: 管制區, 管制標準, 檢測方法 and so on belong to one or the
   *     other. (v4.36：時間起迄 加入共用清單，依使用者說明「同一地點的噪音與振動，
   *     座標一樣之外，時間起迄也會一樣」。)
   *
   * ⚠️ 這和「同一地點不同檢測項目」是**兩件不同的事**，不要混在一起看：
   *   空品／水質／地質 在同一地點做了 PM10 與 PM2.5，那是**同一個檢測類別**底下
   *   的兩個檢測項目——它們本來就走「同檢測類別」那條路徑，改 PM10 的座標
   *   PM2.5 會跟著。這是既有規則，v4.36 完全沒有動到它。
   *   噪音↔振動是**跨檢測類別**，只有噪音類別有這個需求。
   *
   * So the matching rules differ per field, and each caller says what it needs:
   *
   *   座標         same location + same 日期(起) AND 日期(迄); category ignored (noise
   *                and vibration share them); batch ignored (the two sub-reports can
   *                arrive as two separate files)
   *   日期／時間    same location + the date the row had BEFORE the edit;
   *                噪音類別：category ignored + batch ignored（和座標同一組）
   *                其他類別：same 檢測類別 + same batch（維持既有行為）
   *   檢測類別      same location + same dates + the category the row had BEFORE the edit
   *   其他欄位      same location + same 檢測類別 + same dates（噪音與振動各自獨立）
   *
   * `dateStart`/`dateEnd`/`category` override what to compare against, which is what
   * makes "the rows that matched me before I changed this" expressible: after the
   * edit the source no longer carries the old value, so the group has to be defined
   * by the old one.
   */
  const matchesSyncGroup = (source, r, {
    requireBatch = true, requireCategory = true, requireDate = true,
    dateStart = null, dateEnd = null, category = null,
  } = {}) => {
    const locField = cat.locationField;
    if (requireBatch && r._batchId !== source._batchId) return false;
    if (r[locField] !== source[locField]) return false;
    if (requireCategory && r['檢測類別'] !== (category !== null ? category : source['檢測類別'])) return false;
    // Both ends of the sampling window have to match, not just 日期(起) — a site
    // sampled on the same start date but over a different span (a one-off grab
    // sample vs a 24-hour composite) is a different sampling event and must not be
    // swept up by a sync. This is also what separates 平日 from 假日.
    if (requireDate) {
      const ds = dateStart !== null ? dateStart : source['日期(起)'];
      const de = dateEnd !== null ? dateEnd : source['日期(迄)'];
      if (r['日期(起)'] !== ds) return false;
      if (r['日期(迄)'] !== de) return false;
    }
    return true;
  };

  // If the person corrects a coordinate on one row, offer to sync all three
  // coordinate fields to every other row that (a) came from the same import batch,
  // (b) shares the same sampling location, (c) shares the same 檢測類別, AND
  // (d) shares the same sampling date — e.g. a site sampled in both April and May
  // can genuinely have slightly different coordinates between visits, so syncing
  // must never cross dates. Always asks first, same reasoning as the date/time sync
  // below.
  /*
   * 同步提示裡的「共 X 筆」拆成各檢測類別各幾筆。
   * 使用者要的是看得出「同步給噪音幾筆、同步給振動幾筆」再決定按不按確定——
   * 只講一個總數的話，跨檢測類別的同步等於盲按。
   */
  const countByCategory = (matches) => {
    const counts = new Map();
    matches.forEach(({ r }) => {
      const c = r['檢測類別'] || '（未填檢測類別）';
      counts.set(c, (counts.get(c) || 0) + 1);
    });
    return [...counts.entries()].map(([c, n]) => `${c} ${n} 筆`).join('、');
  };

  /*
   * 把「會被改到的是哪幾筆」逐列列出來。
   * ────────────────────────────────────
   * 只講「同步 X 筆噪音、Y 筆振動」，使用者仍然不知道那幾筆是誰——
   * 按確定之前有權知道自己按下去會動到什麼。所以列出每一列的身分
   * （檢測類別／時段／測項／時間）以及**這一格會從什麼變成什麼**。
   *
   * 身分欄位刻意用該類別自己的 itemField（噪音是音源發聲特性、
   * 水質空品地質是檢測項目、生態是調查項目），不寫死欄位名稱。
   */
  const MAX_LISTED_ROWS = 12;
  const affectedRowsTable = (list, fields, source) => {
    const idParts = (r) => [
      r['檢測類別'], r['監測時段'], r[cat.itemField],
      // 日期或時間跟來源不一樣時才顯示，否則整張表都是同一個值、只是佔位置
      (r['日期(起)'] !== source['日期(起)']) ? toDateDisplayValue(r['日期(起)']) : '',
      (r['時間(起)'] !== source['時間(起)']) ? toTimeDisplayValue(r['時間(起)']) : '',
    ].filter(v => String(v ?? '').trim() !== '').join('　');
    const changeText = (r) => fields
      .filter(f => r[f] !== source[f])
      .map((f) => {
        const before = String(r[f] ?? '').trim();
        const after = String(source[f] ?? '').trim();
        const fmt = (v) => v === '' ? '（空白）' : v;
        // 兩邊都要跳脫——欄位值是使用者打進來的，直接接進 HTML 會壞掉
        return `${escapeHtml(f)}：${escapeHtml(fmt(before))} → <strong>${escapeHtml(fmt(after))}</strong>`;
      })
      .join('<br>');
    const shown = list.slice(0, MAX_LISTED_ROWS);
    const rowsHtml = shown.map(({ r }) => `<tr>`
      + `<td style="padding:3px 8px;border-bottom:1px solid var(--border);white-space:nowrap">${escapeHtml(idParts(r) || '（未標示）')}</td>`
      + `<td style="padding:3px 8px;border-bottom:1px solid var(--border)">${changeText(r)}</td>`
      + `</tr>`).join('');
    return `<div style="max-height:220px;overflow:auto;margin:8px 0;border:1px solid var(--border);border-radius:6px">`
      + `<table style="width:100%;border-collapse:collapse;font-size:12.5px">`
      + `<thead><tr style="background:#eef3f1">`
      + `<th style="padding:4px 8px;text-align:left;white-space:nowrap">是哪一筆</th>`
      + `<th style="padding:4px 8px;text-align:left">會怎麼改</th>`
      + `</tr></thead><tbody>${rowsHtml}</tbody></table></div>`
      + `${list.length > MAX_LISTED_ROWS ? `<div class="hint">…另有 ${list.length - MAX_LISTED_ROWS} 筆未列出（同樣會一併更新）。</div>` : ''}`;
  };

  /*
   * 所有同步共用的最後一段：問使用者要套用到哪些，然後套用。
   * ────────────────────────────────────────────────────
   * 三種同步（座標／日期時間／其他欄位）以前各自寫一段 confirm 文字，
   * 「符合的有幾筆、分別是哪個檢測類別」也各算一次。合成一處之後，
   * 以後要多一種同步，呼叫端只要說「配對條件是什麼、要帶哪些欄位」。
   *
   * 噪音類別而且兩邊都有資料時，給的是**多選**（只改這一筆／只給噪音／
   * 只給振動／兩邊都要）；其餘情況是單純的是非題，維持原本的 confirm()，
   * 免得本來一句話就問完的事情也要多開一個視窗。
   *
   * 回傳 true 代表真的改了東西（呼叫端要重畫）。
   */
  const applySync = async ({ rows, rowIdx, matches, fields, what, note = '', when = '', revert = null }) => {
    /*
     * ⚠️ `rows` 一定要由呼叫端傳進來，不可以在這裡重新 DataStore.getData()。
     * getData 每次都會從 localStorage 重新 JSON.parse，回來的是**全新的物件**——
     * 在這裡重讀一次的話，matches 裡的列屬於呼叫端那一份陣列，改了它們卻存另一份，
     * 結果就是視窗按了、畫面重畫了、資料一格都沒變，而且完全沒有錯誤訊息。
     */
    if (syncDialogOpen) return false;   // 已經有一個同步視窗開著了，不要再問一次
    const source = rows[rowIdx];
    if (!source) return false;
    const differing = matches.filter(({ r }) => fields.some(f => r[f] !== source[f]));
    if (differing.length === 0) return false;

    /*
     * 「取消修改」——把剛剛改的那一格還原回去（v4.40）
     * ────────────────────────────────────────────────
     * 使用者回報：誤觸的時候沒有退路。原本四個選項全都是「這一筆已經改了，
     * 只是要不要擴散出去」，少了「我根本不想改」這個答案；
     * 而且 Esc 關掉視窗等同「只改這一筆」，那一格照樣是新值。
     *
     * ⚠️ 只還原**使用者剛剛動的那一格**。座標／日期時間雖然是整組同步，
     * 但來源這一列只有被編輯的那一格變過，其餘欄位本來就沒動。
     * ⚠️ 沒有 baseline（例如程式內部呼叫）時就不顯示這個選項，
     * 不要拿猜的值去「還原」。
     */
    const canRevert = revert && revert.field !== undefined && revert.previous !== undefined
      && String(revert.previous) !== String(source[revert.field] ?? '');
    const revertLabelValue = canRevert
      ? (String(revert.previous).trim() === '' ? '空白'
        : (fields.includes('日期(起)') || fields.includes('日期(迄)')) && /^\d{4}-\d{2}-\d{2}$/.test(String(revert.previous))
          ? toDateDisplayValue(revert.previous)
          : /^\d{1,2}:\d{2}/.test(String(revert.previous)) ? toTimeDisplayValue(revert.previous)
            : String(revert.previous))
      : '';
    const revertOption = canRevert
      ? [{ key: 'revert', label: `取消修改（還原為 ${revertLabelValue}）` }]
      : [];
    const doRevert = () => {
      source[revert.field] = revert.previous;
      DataStore.saveData(project.id, catKey, rows);
      // 資料回到動手之前了，那個復原點按下去不會有任何變化——不要留一個死項目
      popUndoSnapshot(project.id, catKey);
      return true;
    };

    const locField = cat.locationField;
    const where = `同一個測站「${source[locField]}」${when ? `、${when}` : ''}`;
    const myHalf = noiseHalfOf(source);
    const same = differing.filter(({ r }) => noiseHalfOf(r) === myHalf);
    const other = differing.filter(({ r }) => noiseHalfOf(r) !== myHalf);

    let chosen = null;
    if (catKey === 'noise' && same.length > 0 && other.length > 0) {
      const myName = myHalf === 'vib' ? '振動' : '噪音';
      const otherName = myHalf === 'vib' ? '噪音' : '振動';
      const key = await askSyncChoice({
        title: `要把「${what}」一起改到哪些資料？`,
        body: `偵測到${escapeHtml(where)}還有 <strong>${differing.length}</strong> 筆會被改到`
          + `（<strong>${escapeHtml(countByCategory(same))}</strong>與這一筆同為${myName}、`
          + `<strong>${escapeHtml(countByCategory(other))}</strong>）：`
          + affectedRowsTable(differing, fields, source)
          + `${note ? `<div style="margin-top:10px">${note}</div>` : ''}`
          + `<div style="margin-top:6px">不同採樣日期的資料（例如平日／假日）完全不受影響。</div>`,
        wide: true,
        dismissKey: canRevert ? 'revert' : 'none',
        options: [
          { key: 'all', label: `兩邊都同步（${differing.length} 筆）`, primary: true },
          { key: 'other', label: `只同步給${otherName} ${other.length} 筆` },
          { key: 'same', label: `只同步給${myName} ${same.length} 筆` },
          { key: 'none', label: '只改這一筆' },
          ...revertOption,
        ],
      });
      if (key === 'revert') return doRevert();
      if (!key || key === 'none') return false;
      chosen = key === 'all' ? differing : key === 'other' ? other : same;
    } else {
      /*
       * 單邊的情況也要看得到「是哪幾筆」，不能只講一個數字——
       * 使用者按確定之前有權知道自己按下去會動到什麼。
       * 這裡一樣用視窗（只有兩個選項），不再用 confirm()。
       */
      const key = await askSyncChoice({
        title: `要把「${what}」一起改到其他資料嗎？`,
        body: `偵測到${escapeHtml(where)}還有 <strong>${differing.length}</strong> 筆會被改到`
          + `（${escapeHtml(countByCategory(differing))}）：`
          + affectedRowsTable(differing, fields, source)
          + `${note ? `<div style="margin-top:10px">${note}</div>` : ''}`
          + `<div style="margin-top:6px">不同採樣日期的資料（例如平日／假日）完全不受影響。</div>`,
        wide: true,
        dismissKey: canRevert ? 'revert' : 'none',
        options: [
          { key: 'all', label: `一併同步（${differing.length} 筆）`, primary: true },
          { key: 'none', label: '只改這一筆' },
          ...revertOption,
        ],
      });
      if (key === 'revert') return doRevert();
      if (!key || key === 'none') return false;
      chosen = differing;
    }

    chosen.forEach(({ r }) => { fields.forEach(f => { r[f] = source[f]; }); });
    DataStore.saveData(project.id, catKey, rows);
    learnSiteItemHistory(project.id, catKey, cat, chosen.map(m => m.r).concat([source]));
    /*
     * 復原點在編輯發生時就已經拍好了（snapshotEdit），這裡只是把說明補精確。
     * 不再另外拍一張——否則「改一格＋同步」要按兩次復原才回得去，
     * 而使用者心裡那是一個動作。
     */
    retitleUndoSnapshot(project.id, catKey, `修改「${what}」並同步 ${chosen.length} 筆`);
    return true;
  };

  /** 這一次採樣的文字說明，用在提示裡（「同一次採樣（2026/05/10 ～ 2026/05/11）」）。 */
  const visitLabel = (start, end) => `同一次採樣（${toDateDisplayValue(start)}`
    + `${end && end !== start ? ` ～ ${toDateDisplayValue(end)}` : ''}）`;

  /*
   * 一次使用者動作＝一個復原點（v4.40）
   * ──────────────────────────────────
   * 表格裡改一格、以及接下來可能發生的同步，合起來算**一次**動作：
   * 按一下「↩️ 復原上一步」就要全部回到動手之前，不能按兩次。
   *
   * 值在每一次按鍵就已經寫進去了，所以這裡是把 baseline 塞回那一格，
   * 用「改之前」的樣子拍快照，再把現值留著。這樣只複製一次，
   * 也不必在使用者只是點進格子看一眼的時候就先拍一張。
   */
  const snapshotEdit = (rowIdx, fieldKey, previous, label) => {
    if (previous === undefined) return false;
    const before = DataStore.getData(project.id, catKey);
    if (!before[rowIdx]) return false;
    before[rowIdx][fieldKey] = previous;
    pushUndoSnapshot(project.id, catKey, `修改「${label}」`, before);
    return true;
  };

  const offerCoordSync = (rowIdx, fieldKey, prevValue) => {
    const rows = DataStore.getData(project.id, catKey);
    const source = rows[rowIdx];
    const locField = cat.locationField;
    if (!source || !source[locField] || !source['日期(起)']) return Promise.resolve(false);
    // Coordinates are the ONE thing 噪音 and 振動 share: one visit, one tripod, one
    // position — so this sync deliberately ignores 檢測類別, and ignores the import
    // batch too, because the noise and vibration sub-reports often arrive as two
    // separate files. It still never crosses a sampling date, which is what keeps
    // 平日 and 假日 apart.
    const matches = rows
      .map((r, idx) => ({ r, idx }))
      .filter(({ r, idx }) => idx !== rowIdx
        && matchesSyncGroup(source, r, { requireBatch: false, requireCategory: false }));
    return applySync({
      rows, rowIdx, matches, fields: COORD_FIELDS, what: '座標',
      revert: { field: fieldKey, previous: prevValue },
      when: visitLabel(source['日期(起)'], source['日期(迄)']),
      note: '座標是同一次採樣共用的——一次到場就架一次腳架、就一個位置。其他欄位不會被動到。',
    });
  };

  /**
   * Correcting a date/time on one row offers to fix the rest of the SAME visit.
   *
   * The group can't be defined by the source's current dates — the edit has already
   * changed them, and the siblings still carry the old ones. It used to sidestep that
   * by ignoring dates entirely, which is precisely how correcting the 假日 date also
   * rewrote the 平日 rows at that station: same location, same 檢測類別, nothing left
   * to tell them apart. So the group is defined by the dates the row had BEFORE this
   * edit (`prevValue`, captured on focus), which is exactly "the rows that were part
   * of the same visit as this one".
   */
  const DATE_TIME_FIELDS = ['日期(起)', '時間(起)', '日期(迄)', '時間(迄)'];
  /*
   * 噪音類別：日期／時間跨「檢測類別」同步（v4.36）
   * ------------------------------------------------
   * 使用者的說明：同一地點的噪音與振動，是同一次到場的兩份子報告——
   * 「座標會一樣之外，時間起迄也會一樣」。所以這一組欄位改成和座標同一個群組：
   * 忽略檢測類別、忽略匯入檔案（噪音與振動常常是兩份分開的檔案）。
   *
   * ⚠️ 只有噪音類別。空品／水質／地質在同一地點同一天出現兩個檢測類別
   * （例如水質的「河川」與「地下水」）是兩次不同的採樣，時間本來就不見得一樣，
   * 讓它們互相同步會把不相干的資料改掉。
   *
   * ⚠️ 日期和時間仍然是**同一組**在同步，不是只放寬時間：如果只讓時間跨類別，
   * 那麼修正日期時噪音那幾列的日期變了、振動沒變，兩邊就再也配不成
   * 「同一地點＋同一日期」，下一次同步就找不到彼此了。
   *
   * 群組仍然是用「這一列**改之前**的日期」定義的，所以平日／假日依舊分得開。
   */
  const crossCategoryDateTime = catKey === 'noise';
  const offerDateTimeSync = (rowIdx, fieldKey, prevValue) => {
    const rows = DataStore.getData(project.id, catKey);
    const source = rows[rowIdx];
    const locField = cat.locationField;
    if (!source || !source[locField]) return Promise.resolve(false);
    if (!crossCategoryDateTime && !source._batchId) return Promise.resolve(false);
    const prevStart = (fieldKey === '日期(起)' && prevValue !== undefined) ? prevValue : source['日期(起)'];
    const prevEnd = (fieldKey === '日期(迄)' && prevValue !== undefined) ? prevValue : source['日期(迄)'];
    if (!prevStart) return Promise.resolve(false);
    const matches = rows
      .map((r, idx) => ({ r, idx }))
      .filter(({ r, idx }) => idx !== rowIdx
        && matchesSyncGroup(source, r, {
          dateStart: prevStart,
          dateEnd: prevEnd,
          requireBatch: !crossCategoryDateTime,
          requireCategory: !crossCategoryDateTime,
        }));
    return applySync({
      rows, rowIdx, matches, fields: DATE_TIME_FIELDS, what: '採樣日期／時間',
      revert: { field: fieldKey, previous: prevValue },
      when: `原本同樣是 ${toDateDisplayValue(prevStart)}`
        + `${prevEnd && prevEnd !== prevStart ? ` ～ ${toDateDisplayValue(prevEnd)}` : ''} 這一次採樣的`,
      note: crossCategoryDateTime
        ? '同一地點的噪音與振動是同一次到場，座標與時間起迄共用。管制標準、管制區、監測數值等各自不同的欄位不會被動到。'
        : '只會套用到「原本跟這一筆同一次採樣」的資料，不同檢測類別的資料不會被動到。',
    });
  };

  // 檢測類別 sync follows the same rule as coordinates (per the person's own
  // clarification): same batch (file) + same site + same sampling date. Unlike
  // date/time sync, this must NOT cross dates — a site's category classification for
  // one visit shouldn't silently overwrite a different visit's classification. This
  // is the one sync that can't require "same category" as a matching criterion,
  // since 檢測類別 is the very field being synced.
  const offerCategorySync = (rowIdx, prevCategory) => {
    const rows = DataStore.getData(project.id, catKey);
    const source = rows[rowIdx];
    const locField = cat.locationField;
    if (!source || !source._batchId || !source[locField] || !source['日期(起)']) return false;
    // Same trick as the date sync: the group is "the rows that had the SAME category
    // as this one before the edit". Matching on anything else would sweep in the
    // 振動 rows sitting at the same site on the same day and reclassify them too.
    if (prevCategory === undefined || prevCategory === source['檢測類別']) return false;
    const matches = rows
      .map((r, idx) => ({ r, idx }))
      .filter(({ r, idx }) => idx !== rowIdx && matchesSyncGroup(source, r, { category: prevCategory }));
    if (matches.length === 0) return false;
    const anyDiff = matches.some(({ r }) => r['檢測類別'] !== source['檢測類別']);
    if (!anyDiff) return false;
    const ok = confirm(
      `偵測到同一份檔案、同一天（${toDateDisplayValue(source['日期(起)'])}）、同一個測站「${source[locField]}」，`
      + `原本同樣是「${prevCategory || '（空白）'}」的還有 ${matches.length} 筆資料。\n`
      + `是否要將這些資料的檢測類別一併同步更新為「${source['檢測類別']}」？\n\n`
      + `（只會套用到原本就是「${prevCategory || '（空白）'}」的資料，其他檢測類別（例如振動）與不同採樣日期的資料不受影響。`
      + `選擇「取消」則只修改目前這一筆。）`
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
  // 比較關係 is DERIVED per row from that row's own value (deriveComparisonRelation),
  // so it must never be propagated sideways: syncing a blank one across a site turned
  // every ND row there into a row carrying neither a value nor an ND marker.
  const SYNC_EXCLUDED_FIELDS = new Set([
    cat.itemField, cat.locationField, '檢測數值', '監測數值', '比較關係',
    ...COORD_FIELDS, ...DATE_TIME_FIELDS, '檢測類別',
  ]);
  /*
   * 噪音類別：這一格能不能同步給「另一半」？（v4.37）
   * ────────────────────────────────────────────────
   * 預設是**可以，而且會先問**——這樣以後多一個欄位不用再改一次程式。
   * 例外寫在 schema.js 的 NOISE_VIB_DISTINCT_FIELDS，那些欄位在噪音與振動
   * 之間本來就不一樣（管制標準、管制區、監測單位、監測方法…），
   * 同步過去就是填錯，所以連問都不問。
   */
  const canCrossNoiseVib = (fieldKey) =>
    catKey === 'noise' && !NOISE_VIB_DISTINCT_FIELDS.includes(fieldKey);

  const offerGenericFieldSync = (rowIdx, fieldKey, prevValue) => {
    const rows = DataStore.getData(project.id, catKey);
    const source = rows[rowIdx];
    const locField = cat.locationField;
    if (!source || !source[locField] || !source['日期(起)']) return Promise.resolve(false);
    const cross = canCrossNoiseVib(fieldKey);
    // 跨噪音／振動時不要求同一份匯入檔案——兩者常常是分開的兩份報告。
    if (!cross && !source._batchId) return Promise.resolve(false);
    const matches = rows
      .map((r, idx) => ({ r, idx }))
      .filter(({ r, idx }) => idx !== rowIdx
        && matchesSyncGroup(source, r, { requireBatch: !cross, requireCategory: !cross }));
    const fieldLabel = (cat.fields.find(f => f.key === fieldKey) || {}).label || fieldKey;
    return applySync({
      rows, rowIdx, matches, fields: [fieldKey],
      revert: { field: fieldKey, previous: prevValue },
      what: `「${fieldLabel}」（${source[fieldKey] || '空白'}）`,
      when: `同一天（${toDateDisplayValue(source['日期(起)'])}）`,
      note: cross
        ? '這個欄位噪音與振動可以共用，所以兩邊都能選。'
        : `「${fieldLabel}」在噪音與振動之間本來就不一樣（官方規定或報告本身就不同），所以只會套用到同一個檢測類別。`,
    });
  };

  // Remember what a control held when it gained focus, so the sync prompts can tell
  // "the person edited this" from "the person merely clicked in and out". Without it,
  // clicking a 檢測方法 cell just to read it and then clicking away offered to
  // overwrite that field on every other row at the site.
  tbody.addEventListener('focusin', (e) => {
    const t = e.target;
    if (t && t.dataset && t.dataset.field !== undefined) t.dataset.syncBaseline = t.value;
  });
  const valueUnchanged = (t) => !!t && !!t.dataset && t.dataset.syncBaseline !== undefined
    && t.dataset.syncBaseline === t.value;

  tbody.addEventListener('input', (e) => {
    const t = e.target;
    if (!t.dataset.field) return;
    if (t.classList.contains('code-input')) {
      t.title = t.dataset.codetype === 'unit' ? lookupUnit(t.value) : lookupAgency(t.value);
    }
    if (t.dataset.field === LIMIT_FIELD) {
      const bad = t.value.trim() !== '' && !isPlainNumber(t.value);
      t.classList.toggle('cell-invalid', bad);
      t.title = bad ? '「檢測極限」只能填數值。請改成純數字，或清空這一格。' : '';
    }
    if (t.tagName === 'SELECT') return; // handled by change
    commit(Number(t.dataset.row), t.dataset.field, t.value, { learn: false });
  });
  tbody.addEventListener('change', async (e) => {
    const t = e.target;
    if (!t.dataset.field) return;
    if (t.tagName === 'SELECT') {
      // What the dropdown held before this change — the 檢測類別 sync needs it to
      // work out which rows were in the same group as this one beforehand.
      const previous = t.dataset.syncBaseline;
      commit(Number(t.dataset.row), t.dataset.field, t.value);
      t.dataset.syncBaseline = t.value; // a second change compares against this one
      const fieldKey = t.dataset.field;
      const rowIdx = Number(t.dataset.row);
      const label = (cat.fields.find(f => f.key === fieldKey) || {}).label || fieldKey;
      const snapped = snapshotEdit(rowIdx, fieldKey, previous, label);
      // 同步視窗是非同步的（要等使用者按按鈕），所以這裡要 await。
      if (COORD_FIELDS.includes(fieldKey) && await offerCoordSync(rowIdx, fieldKey, previous)) { renderContentPreservingScroll(); return; }
      if (fieldKey === '檢測類別' && offerCategorySync(rowIdx, previous)) { renderContentPreservingScroll(); return; }
      if (!SYNC_EXCLUDED_FIELDS.has(fieldKey) && await offerGenericFieldSync(rowIdx, fieldKey, previous)) { renderContentPreservingScroll(); return; }
      if (snapped) renderContentPreservingScroll(); // 讓「↩️ 復原上一步」按鈕出現
    }
  });
  // use focusout (bubbles) rather than blur to catch this via delegation; only
  // re-render on blur (not every keystroke) so typing isn't interrupted.
  tbody.addEventListener('focusout', async (e) => {
    const t = e.target;
    if (!t.dataset.field || t.tagName === 'SELECT') return;
    const rowIdx = Number(t.dataset.row);
    const fieldKey = t.dataset.field;

    // The two branches below already commit the CANONICAL value (2026-07-01,
    // 09:15:00) while putting the display form (2026/07/01, 09:15) in the box. The
    // generic re-commit further down must therefore skip them, or it writes the
    // display string back into storage — which then exports as a text cell instead
    // of a real Excel date, and stops matching on re-import so the same rows come
    // back as duplicates.
    const isDateTimeInput = t.classList.contains('time-input') || t.classList.contains('date-input');

    if (t.classList.contains('time-input')) {
      let normalized = normalizeTimeString(t.value); // canonical HH:MM:00 for storage
      /*
       * 打了 24:00 就當場問（v4.37）。
       *
       * 官方五份資料辭典都寫「勿輸入 24:00，僅能輸入 00:00~23:59」。
       * 舊版是**匯出的時候**把 24:00 悄悄寫成 00:00——畫面上是一天的結尾、
       * 檔案裡卻是同一天的開頭，而且完全沒有提示。
       *
       * ⚠️ 按「取消」時**不會**幫他改，值原樣留著、格子畫紅框。
       * 使用者可能是照著報告上的寫法先打進來，之後才要回頭確認到底是幾點；
       * 幫他挑一個數字等於又做了一次「靜靜改掉輸入」，那正是這一版要修的事。
       */
      const badTime = DateTimeUtil.outOfRangeTimeReason(normalized);
      if (badTime && confirm(
        `「${toTimeDisplayValue(normalized) || normalized}」${badTime}。\n\n`
        + `要改成 23:59（一天的最後一分鐘）嗎？\n\n`
        + `按「確定」改成 23:59；按「取消」保留現在填的內容，\n`
        + `這一格會用紅框標示，匯出前也會再提醒一次。`
      )) {
        normalized = DateTimeUtil.clampToDayEnd(normalized);
      }
      // slice(0,5) blindly chopped anything it couldn't parse — a value the person
      // typed, or an imported "連續24小時", lost its last characters on every blur.
      t.value = toTimeDisplayValue(normalized) || normalized;
      commit(rowIdx, fieldKey, normalized);
      /*
       * 紅框要當場出現，不能等下一次整頁重繪——按了「取消」之後表格不會重畫，
       * 使用者看到的還是原來那一格。標示如果要捲開才看得到就不算標示。
       */
      const stillBad = DateTimeUtil.outOfRangeTimeReason(normalized);
      t.classList.toggle('cell-invalid', Boolean(stillBad));
      t.title = stillBad
        ? `「${toTimeDisplayValue(normalized) || normalized}」${stillBad}。請改成 23:59（一天的最後一分鐘）或正確的時間。`
        : '';
    }
    if (t.classList.contains('date-input')) {
      const normalized = normalizeDateString(t.value); // ISO YYYY-MM-DD for storage
      t.value = toDateDisplayValue(normalized) || normalized; // YYYY/MM/DD for display
      commit(rowIdx, fieldKey, normalized);
    }

    // The `input` handler deliberately skips history learning (see commit's `learn`
    // flag). Do it once here, when the field is finished, so the remembered snapshot
    // reflects the completed value rather than every prefix of it.
    const edited = !valueUnchanged(t);
    if (edited && !isDateTimeInput) {
      commit(rowIdx, fieldKey, t.value, { learn: true });
      // 補零之後要把畫面上的輸入框也一起換過來，否則存的是 39.20、
      // 格子裡卻還寫著 39.2，直到下一次整頁重繪才對得上。
      // 失焦時把輸入框裡的字換成格式化後的樣子，讓它跟表格其他列一致。
      // 存起來的仍然是原始值（見 commit 裡的說明），所以這只是顯示。
      const shown = formatFieldValue(catKey, fieldKey, t.value);
      if (shown !== t.value) t.value = shown;
    }
    // What this field held before the edit, in STORAGE form — the date/time sync
    // needs it to identify "the rows that were part of the same visit as this one",
    // which the source itself can no longer say once its own date has changed.
    const baseline = t.dataset.syncBaseline;
    const prevStored = baseline === undefined ? undefined
      : (t.classList.contains('date-input') ? normalizeDateString(baseline)
        : t.classList.contains('time-input') ? normalizeTimeString(baseline) : baseline);
    delete t.dataset.syncBaseline;
    if (!edited) return; // focused and left without editing — nothing to sync

    const label = (cat.fields.find(f => f.key === fieldKey) || {}).label || fieldKey;
    const snapped = snapshotEdit(rowIdx, fieldKey, prevStored, label);
    if (COORD_FIELDS.includes(fieldKey) && await offerCoordSync(rowIdx, fieldKey, prevStored)) { renderContentPreservingScroll(); return; }
    if (DATE_TIME_FIELDS.includes(fieldKey) && await offerDateTimeSync(rowIdx, fieldKey, prevStored)) { renderContentPreservingScroll(); return; }
    if (!SYNC_EXCLUDED_FIELDS.has(fieldKey) && await offerGenericFieldSync(rowIdx, fieldKey, prevStored)) { renderContentPreservingScroll(); return; }
    if (snapped) renderContentPreservingScroll(); // 讓「↩️ 復原上一步」按鈕出現
  });
  tbody.addEventListener('click', (e) => {
    const btn = e.target.closest('.row-del-btn');
    if (!btn) return;
    if (!confirm('確定要刪除這一列資料嗎？（可用「↩️ 復原上一步」救回）')) return;
    pushUndoSnapshot(project.id, catKey, '刪除 1 筆資料');
    const rows = DataStore.getData(project.id, catKey);
    rows.splice(Number(btn.dataset.row), 1);
    DataStore.saveData(project.id, catKey, rows);
    renderContentPreservingScroll();
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

/**
 * 手動新增一列時要帶上期別。
 *
 * _period 以前只有匯入那條路徑會寫（finalizeImportCommit），手動新增與
 * 「常用測項新增」都沒有。後果：那些列變成「未標示期別」，而畫面正在篩某一季時
 * 它們**當場就從表格上消失**，依季度匯出時也不會被帶進去——使用者補了一筆
 * 漏掉的測項，結果那一筆從來沒有被送出去，而且沒有任何提示。
 * 而且期別在畫面上沒有地方可以編輯，所以事後也救不回來。
 *
 * 規則：目前正在篩某一季就用那一季；沒有篩選（顯示全部）時，用這個類別
 * 最後一次匯入的期別——那幾乎一定是使用者現在正在整理的那一季。
 */
function currentPeriodForNewRow(project, catKey) {
  const filtered = state.periodFilter[catKey];
  if (filtered && filtered !== '__none__') return filtered;
  const batches = DataStore.getImportBatches(project.id, catKey);
  for (let i = batches.length - 1; i >= 0; i--) if (batches[i].period) return batches[i].period;
  const rows = DataStore.getData(project.id, catKey);
  for (let i = rows.length - 1; i >= 0; i--) if (rows[i]._period) return rows[i]._period;
  return '';
}

function addEmptyRow(project, catKey) {
  const cat = CATEGORIES[catKey];
  const rows = DataStore.getData(project.id, catKey);
  const blank = { _period: currentPeriodForNewRow(project, catKey) };
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
  // The table body is never cleared on close, and renderCommonItemsTableBody
  // deliberately re-harvests checked state from the existing DOM so that sorting and
  // filtering don't lose a selection. Opening the modal for a DIFFERENT category then
  // re-applied the previous category's ticks BY ROW INDEX — three 水質 ticks became
  // three unrelated 空氣品質 items, added to the filing without the person choosing
  // them, and written into the project's item memory.
  const staleBody = document.getElementById('commonItemsTableBody');
  if (staleBody) staleBody.innerHTML = '';
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

  // The other two openers of this shared popup register this; this one didn't, so
  // clicking anywhere else left the dropdown open — and closing the modal underneath
  // left it stranded over the data grid, blocking clicks, with no way to dismiss it.
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
  renderContentPreservingScroll();
  alert('已套用座標到符合的資料列（僅限相同測站＋相同採樣日期的資料）。');
}

// ---------- method manager (test method + unit code, remembered across seasons) ----------
function openMethodModal(project, catKey) {
  const cat = CATEGORIES[catKey];
  const rows = DataStore.getData(project.id, catKey);
  const memory = DataStore.getItemMemory(project.id, catKey);
  const itemField = cat.itemField;

  const groups = {}; // itemName -> { indices, method, unitCode, methodSet, unitSet }
  rows.forEach((row, idx) => {
    const item = (row[itemField] || '').trim() || '（未命名項目）';
    if (!groups[item]) {
      groups[item] = {
        indices: [],
        method: row[cat.methodField] || (memory[item] && memory[item].method) || '',
        unitCode: cat.unitField ? (row[cat.unitField] || (memory[item] && memory[item].unitCode) || '') : null,
        // 同一個測項目前實際存在幾種不同的方法／單位。跨季是合法的差異：
        // 溶氧 115Q1 用 NIEA W455、115Q2 用 W422，兩個都對。這裡要讓使用者
        // 看得到「按下去會把它們統一成一種」。
        methodSet: new Set(),
        unitSet: new Set(),
      };
    }
    groups[item].indices.push(idx);
    if (row[cat.methodField]) groups[item].methodSet.add(row[cat.methodField]);
    if (cat.unitField && row[cat.unitField]) groups[item].unitSet.add(row[cat.unitField]);
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
        ${entries.map(([item, g]) => {
          const mixMethod = g.methodSet.size > 1 ? [...g.methodSet].join('、') : '';
          const mixUnit = g.unitSet.size > 1 ? [...g.unitSet].join('、') : '';
          const warn = mixMethod || mixUnit
            ? `<div class="hint" style="color:#b45309;margin-top:4px">⚠️ 目前這個測項有不只一種${mixMethod ? `檢測方法（${escapeHtml(mixMethod)}）` : ''}${mixMethod && mixUnit ? '、' : ''}${mixUnit ? `單位代碼（${escapeHtml(mixUnit)}）` : ''}。跨季本來就可能不同，按「儲存」會把它們全部改成同一種。</div>`
            : '';
          return `<tr data-item="${escapeAttr(item)}">
          <td>${escapeHtml(item)}${!g.method ? ' <span class="req" title="尚未有檢測方法">＊</span>' : ''}${warn}</td>
          <td><input type="text" data-method-field="method" value="${escapeAttr(g.method)}" data-seeded="${escapeAttr(g.method)}" placeholder="例：NIEA W417"></td>
          ${cat.unitField ? `<td><input type="text" data-method-field="unitCode" value="${escapeAttr(g.unitCode || '')}" data-seeded="${escapeAttr(g.unitCode || '')}" class="code-input" data-codetype="unit" title="${escapeAttr(lookupUnit(g.unitCode))}" placeholder="代碼"></td>` : ''}
          <td>${g.indices.length}</td>
        </tr>`;
        }).join('')}
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

  /*
   * ⚠️ 只寫「使用者真的改過」的欄位。
   *
   * 舊寫法是把畫面上每一格的值無條件套用到所有同名測項的資料列，而畫面上的值
   * 是拿「第一筆遇到的」當種子。後果：打開這個視窗、什麼都沒改、按一下儲存，
   * 就會把其他季度不同的檢測方法一起改成第一筆那個。實測：溶氧 115Q1 是
   * NIEA W455、115Q2 是 NIEA W422，按下儲存之後兩季都變成 W455——
   * 而 Q2 的申報就帶著實驗室從來沒用過的方法。
   *
   * data-seeded 存的是開窗當下填進去的值；和現在的值相同就代表沒動過，跳過。
   */
  document.querySelectorAll('#methodItemsBody tr').forEach(tr => {
    const item = tr.dataset.item;
    const values = {};
    tr.querySelectorAll('[data-method-field]').forEach(el => {
      const targetField = el.dataset.methodField === 'method' ? cat.methodField : cat.unitField;
      if (!targetField) return;
      if (el.value === (el.dataset.seeded ?? '')) return; // 沒改過就不要動任何資料列
      values[targetField] = el.value;
    });
    if (Object.keys(values).length === 0) return;
    rows.forEach(row => {
      const rowItem = (row[itemField] || '').trim() || '（未命名項目）';
      if (rowItem === item) { Object.assign(row, values); touchedRows.push(row); }
    });
    const memFields = {};
    if (values[cat.methodField]) memFields.method = values[cat.methodField];
    if (cat.unitField && values[cat.unitField]) memFields.unitCode = values[cat.unitField];
    if (Object.keys(memFields).length) memoryUpdates[item] = memFields;
  });

  if (touchedRows.length === 0 && Object.keys(memoryUpdates).length === 0) {
    popUndoSnapshot(projectId, catKey); // 什麼都沒改，不要留一個空的復原點
    closeMethodModal();
    alert('沒有任何欄位被修改，資料維持原樣。');
    return;
  }
  DataStore.saveData(projectId, catKey, rows);
  if (Object.keys(memoryUpdates).length) DataStore.updateItemMemory(projectId, catKey, memoryUpdates);
  // Also refresh the FULL per-location snapshot memory, not just the flat item-name
  // memory above — otherwise a correction made here wouldn't carry through to next
  // season's "entirely absent location" reconstruction, which prefers the fuller
  // snapshot over the flat item memory whenever both exist, and would keep
  // resurrecting the pre-correction method/unit indefinitely.
  if (touchedRows.length > 0) learnSiteItemHistory(projectId, catKey, cat, touchedRows);
  closeMethodModal();
  renderContentPreservingScroll();
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
  // Collect first, check 檢測極限 for every chosen category, and only then start
  // downloading — so the person isn't answering a warning after half the files
  // have already been saved.
  const jobs = [];
  checked.forEach(catKey => {
    let rows = DataStore.getData(project.id, catKey);
    if (period === '__none__') rows = rows.filter(r => !r._period);
    else if (period) rows = rows.filter(r => r._period === period);
    if (rows.length === 0) return; // nothing for this category in the chosen period
    jobs.push({ catKey, rows });
  });
  for (const job of jobs) {
    if (!confirmNoiseVibRuleBeforeExport(job.rows, CATEGORIES[job.catKey])) return;
    if (!confirmTimeRangeBeforeExport(job.rows, CATEGORIES[job.catKey])) return;
    if (!confirmLimitBeforeExport(job.rows, CATEGORIES[job.catKey])) return;
    if (!confirmRequiredBeforeExport(job.rows, CATEGORIES[job.catKey])) return;
  }
  let exportedCount = 0;
  jobs.forEach(({ catKey, rows }) => {
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

// ---------- 檢測極限 must hold a bare number ----------
// The filing only accepts a plain number in 檢測極限. Reports write all sorts of
// things there — "<0.001", "--", "─", "N/A", "0.001 mg/L" — and any of them would be
// rejected (or silently misread) by the receiving system. Rather than guessing, the
// import preview shows exactly what was found and asks what to do with it.

/** A value that 檢測極限 can legally hold: empty, or a bare number. */
function isPlainNumber(v) {
  const s = String(v ?? '').trim();
  return s !== '' && /^[+-]?(\d+(\.\d+)?|\.\d+)([eE][+-]?\d+)?$/.test(s);
}
/** The first number embedded in a value ("<0.001" -> "0.001", "0.001 mg/L" -> "0.001"),
 *  or '' when there is no number in there at all ("--", "N/A", "無"). */
function extractPlainNumber(v) {
  const m = String(v ?? '').match(/[+-]?(\d+(\.\d+)?|\.\d+)([eE][+-]?\d+)?/);
  return m ? String(Number(m[0])) : '';
}
const LIMIT_FIELD = '檢測極限';

/** Groups every offending 檢測極限 value in `rows`, most frequent first. */
function findBadLimitValues(rows, cat) {
  if (!cat.fields.some(f => f.key === LIMIT_FIELD)) return [];
  const counts = new Map();
  rows.forEach(r => {
    const v = String(r[LIMIT_FIELD] ?? '').trim();
    if (v === '' || isPlainNumber(v)) return;
    if (!counts.has(v)) counts.set(v, { value: v, count: 0, numeric: extractPlainNumber(v), items: new Set() });
    const e = counts.get(v);
    e.count++;
    if (e.items.size < 4) e.items.add(r[cat.itemField] || '（未標示）');
  });
  return [...counts.values()].sort((a, b) => b.count - a.count);
}

/**
 * Renders the "檢測極限 isn't a number" prompt into `containerEl`, or clears it when
 * everything is clean. The person's choice lives in `state.limitFixMode` and is
 * applied by applyLimitFixMode at confirm time.
 */
function renderLimitWarning(containerEl, rows, cat, onChange) {
  if (!containerEl) return;
  const bad = findBadLimitValues(rows, cat);
  if (bad.length === 0) { containerEl.innerHTML = ''; state.limitFixMode = 'blank'; return; }
  if (!state.limitFixMode) state.limitFixMode = 'blank';
  const total = bad.reduce((n, b) => n + b.count, 0);
  const anyNumeric = bad.some(b => b.numeric !== '');
  const opt = (val, label, hint) => `
    <label class="limit-fix-option">
      <input type="radio" name="limitFixMode" value="${val}" ${state.limitFixMode === val ? 'checked' : ''}>
      <span><b>${label}</b><br><span class="hint">${hint}</span></span>
    </label>`;
  containerEl.innerHTML = `
    <div class="warning warning-strong">
      ⚠️ <b>「檢測極限」欄位只能填數值</b>，但這批資料裡有 ${total} 筆不是純數值：
      <ul style="margin:6px 0 8px 0">
        ${bad.slice(0, 8).map(b => `<li><code>${escapeHtml(b.value)}</code> — ${b.count} 筆`
          + `（${escapeHtml([...b.items].join('、'))}${b.items.size >= 4 ? '…' : ''}）`
          + `${b.numeric !== '' ? `　→ 只取數字會變成 <code>${escapeHtml(b.numeric)}</code>` : '　→ 裡面沒有數字'}</li>`).join('')}
        ${bad.length > 8 ? `<li class="hint">（另有 ${bad.length - 8} 種其他值未列出）</li>` : ''}
      </ul>
      請選擇要怎麼處理，再按下方的確認匯入：
      <div class="limit-fix-options">
        ${opt('blank', '清成空白，不匯入這個值（建議）', '其他欄位照常匯入，只有「檢測極限」留白。申報時空白是允許的，填錯字元則會被退件。')}
        ${anyNumeric ? opt('number', '只保留數字部分', '例如「&lt;0.001」會變成「0.001」。適合報告只是多寫了符號或單位的情況。') : ''}
        ${opt('keep', '照原樣匯入', '原封不動帶進表格。您可以之後在表格裡自行修改，但送件前請確認已改成純數值。')}
      </div>
    </div>`;
  containerEl.querySelectorAll('input[name="limitFixMode"]').forEach(rb => {
    rb.addEventListener('change', () => {
      if (rb.checked) state.limitFixMode = rb.value;
      if (onChange) onChange();
    });
  });
}

/** Applies the person's choice to the rows about to be committed. */
function applyLimitFixMode(rows, cat) {
  if (!cat.fields.some(f => f.key === LIMIT_FIELD)) return;
  const mode = state.limitFixMode || 'blank';
  if (mode === 'keep') return;
  rows.forEach(r => {
    const v = String(r[LIMIT_FIELD] ?? '').trim();
    if (v === '' || isPlainNumber(v)) return;
    r[LIMIT_FIELD] = mode === 'number' ? extractPlainNumber(v) : '';
  });
}

/**
 * Last-chance guard before a file leaves the app. A non-numeric 檢測極限 is rejected
 * by the receiving system, and by this point it can only have got there by hand or by
 * an explicit "照原樣匯入" — so ask rather than silently exporting it.
 * Returns true if the export should go ahead.
 */
/**
 * 匯出前檢查必填欄位（依 115 年版官方資料辭典，見 schema.js 的 missingRequiredFields）。
 *
 * 為什麼一定要有這一關：
 *  ・「補回缺少的測項」產生的空白列，日期(起)、數值、單位全是空的，而且**預設打勾**。
 *    確認之後那些列就躺在表格裡，匯出時原封不動送出去。
 *  ・地質報告本身沒寫檢測類別，程式不猜（那是對的），但它是必填欄位。
 *  ・從報告讀不到座標的類別（水質、地質常見）也是必填。
 * 這些以前完全沒有人擋，只有「檢測極限不是純數值」有檢查。
 *
 * ⚠️ 只提醒，不阻擋。114 年以前的舊格式有些欄位當時不是必填，使用者自己最清楚
 * 這一份是不是真的需要補；所以列出來讓他決定「回去修正」還是「照樣匯出」。
 */
function confirmRequiredBeforeExport(rows, cat) {
  const counts = new Map(); // 欄位 → 缺的筆數
  let badRows = 0;
  rows.forEach((r) => {
    const missing = missingRequiredFields(r, cat);
    if (missing.length === 0) return;
    badRows += 1;
    missing.forEach((m) => {
      const key = `${m.label}｜${m.why}`;
      counts.set(key, (counts.get(key) || 0) + 1);
    });
  });
  if (badRows === 0) return true;
  const list = [...counts.entries()]
    .sort((a, b) => b[1] - a[1])
    .slice(0, 8)
    .map(([k, n]) => {
      const [label, why] = k.split('｜');
      return `　・${label}：${n} 筆未填${why === '必填' ? '' : `（${why}）`}`;
    })
    .join('\n');
  return confirm(
    `「${cat.label}」有 ${badRows} 筆資料的必填欄位還沒填：\n${list}`
    + `${counts.size > 8 ? `\n　…另有 ${counts.size - 8} 種欄位` : ''}\n\n`
    + `這些是官方資料辭典列為必填的欄位，空白可能導致上傳被退件。\n`
    + `（在表格中這些格子會用紅框標示，滑鼠移上去有說明）\n\n`
    + `按「確定」仍要照原樣匯出，按「取消」回去修正。`
  );
}

/*
 * 匯出前清點超出 00:00~23:59 的時間（v4.37）。
 *
 * 為什麼要有這一道：畫面上的紅框只有捲到那一格才看得到，而一份季報有上百列。
 * 這是最後一道——舊版在這個位置什麼都不做，24:00 就這樣被寫成 00:00 送出去。
 */
function findOutOfRangeTimes(rows, cat) {
  const timeFields = cat.fields.filter(f => f.type === 'time').map(f => f.key);
  const counts = new Map(); // 「欄位｜值」 → 筆數
  rows.forEach((r) => {
    timeFields.forEach((k) => {
      if (!DateTimeUtil.outOfRangeTimeReason(r[k])) return;
      const key = `${k}｜${toTimeDisplayValue(r[k]) || r[k]}`;
      counts.set(key, (counts.get(key) || 0) + 1);
    });
  });
  return [...counts.entries()].sort((a, b) => b[1] - a[1]);
}

function confirmTimeRangeBeforeExport(rows, cat) {
  const bad = findOutOfRangeTimes(rows, cat);
  if (bad.length === 0) return true;
  const total = bad.reduce((n, [, c]) => n + c, 0);
  const list = bad.slice(0, 6)
    .map(([k, n]) => { const [f, v] = k.split('｜'); return `　・${f}＝${v}（${n} 筆）`; })
    .join('\n');
  return confirm(
    `「${cat.label}」有 ${total} 筆資料的時間超出 00:00~23:59：\n${list}`
    + `${bad.length > 6 ? `\n　…另有 ${bad.length - 6} 種` : ''}\n\n`
    + `官方規定時間限 24 小時制的 00:00~23:59，明講「勿輸入 24:00」。\n`
    + `一天的結尾請填 23:59。\n\n`
    + `（在表格中這些格子會用紅框標示，滑鼠移上去有說明）\n`
    + `按「確定」仍要照原樣匯出——這些格子會以文字寫入，不會被改成 00:00；\n`
    + `按「取消」回去修正。`
  );
}

/*
 * 匯出前清點「噪音／振動連動規則」填錯的資料（v4.42）
 * ────────────────────────────────────────────────────
 * 表格上的紅框只有捲到那一格才看得到，而一份季報有上百列。
 * 這是最後一道——尤其是別家公司交來、或上一季匯入的資料，
 * 它們從來沒有經過這個程式的任何一條編輯路徑，紅框可能一次都沒被看到過。
 */
function findNoiseVibRuleViolations(rows, cat) {
  const counts = new Map(); // 「檢測類別｜欄位｜值」 → 筆數
  rows.forEach((r) => {
    noiseVibRuleViolations(r, cat).forEach((v) => {
      const key = `${String(r['檢測類別'] ?? '').trim()}｜${v.key}｜${v.value}`;
      counts.set(key, (counts.get(key) || 0) + 1);
    });
  });
  return [...counts.entries()].sort((a, b) => b[1] - a[1]);
}

function confirmNoiseVibRuleBeforeExport(rows, cat) {
  const bad = findNoiseVibRuleViolations(rows, cat);
  if (bad.length === 0) return true;
  const total = bad.reduce((n, [, c]) => n + c, 0);
  const list = bad.slice(0, 8).map(([k, n]) => {
    const [category, field, value] = k.split('｜');
    return `　・檢測類別「${category}」的「${field}」＝「${value}」：${n} 筆`
      + `（${category === '振動' ? '振動必須填「無」' : '噪音不能填「無」'}）`;
  }).join('\n');
  return confirm(
    `「${cat.label}」有 ${total} 筆資料的管制標準／管制區和檢測類別對不起來：\n${list}`
    + `${bad.length > 8 ? `\n　…另有 ${bad.length - 8} 種` : ''}\n\n`
    + `官方規定：檢測類別是「振動」時，管制標準與管制區必須填「無」；\n`
    + `是噪音時則不能填「無」。115 年 7 月起加強檢核這個組合。\n\n`
    + `（在表格中這些格子會用紅框標示，可用欄位篩選 ▾ 快速找到）\n`
    + `按「確定」仍要照原樣匯出，按「取消」回去修正。`
  );
}

function confirmLimitBeforeExport(rows, cat) {
  const bad = findBadLimitValues(rows, cat);
  if (bad.length === 0) return true;
  const total = bad.reduce((n, b) => n + b.count, 0);
  const list = bad.slice(0, 6).map(b => `　・${b.value}（${b.count} 筆）`).join('\n');
  return confirm(
    `「${cat.label}」有 ${total} 筆資料的「檢測極限」不是純數值：\n${list}`
    + `${bad.length > 6 ? `\n　…另有 ${bad.length - 6} 種` : ''}\n\n`
    + `這個欄位只能填數值，含有「<」「>」等字元可能導致上傳被退件。\n`
    + `（在表格中這些格子會用紅框標示，可用欄位篩選 ▾ 快速找到）\n\n`
    + `按「確定」仍要照原樣匯出，按「取消」回去修正。`
  );
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
      // Deliberately NOT resetting state.excludedRowIndices. _rowUid indexes
      // result.rows and stays valid across statistic changes, and rows belonging to a
      // de-selected statistic are already filtered out by filterRowsByStat. Clearing
      // it silently re-included the rows the app had auto-excluded for having no
      // sampling date — the guard against a workbook's leftover older sheets — and
      // threw away any row the person had unticked by hand.
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
function renderRowDetailTable(containerEl, rows, cat, onChange) {
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
  // The caller passes a callback that refreshes the confirm button, the summary line
  // and the per-site counts. It used to be declared away (the function took only
  // three parameters), so unticking rows here changed what would be imported while
  // every number on screen kept claiming the original count — the wrong number in the
  // unsafe direction, shown at the exact moment the person checks their exclusions.
  const notifyChanged = () => { syncCheckAllState(); if (onChange) onChange(); };
  getRowChecks().forEach(cb => {
    cb.addEventListener('change', () => {
      const uid = Number(cb.dataset.rowUid);
      if (cb.checked) state.excludedRowIndices.delete(uid); else state.excludedRowIndices.add(uid);
      // 使用者自己動過的勾選永遠優先：從自動排除名單移除，之後改期別也不會再被蓋掉
      if (state.autoExcludedByPeriod) state.autoExcludedByPeriod.delete(uid);
      if (state.manualRowChoices) state.manualRowChoices.add(uid);
      notifyChanged();
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
      if (state.autoExcludedByPeriod) state.autoExcludedByPeriod.delete(uid);
      if (state.manualRowChoices) state.manualRowChoices.add(uid);
    });
    notifyChanged();
  });
  containerEl.querySelector('#rowDetailSelectAllVisible').addEventListener('click', () => {
    getVisibleRowChecks().forEach(cb => { cb.checked = true; const u = Number(cb.dataset.rowUid); state.excludedRowIndices.delete(u); if (state.autoExcludedByPeriod) state.autoExcludedByPeriod.delete(u); if (state.manualRowChoices) state.manualRowChoices.add(u); });
    notifyChanged();
  });
  containerEl.querySelector('#rowDetailClearAllVisible').addEventListener('click', () => {
    getVisibleRowChecks().forEach(cb => { cb.checked = false; const u = Number(cb.dataset.rowUid); state.excludedRowIndices.add(u); if (state.autoExcludedByPeriod) state.autoExcludedByPeriod.delete(u); if (state.manualRowChoices) state.manualRowChoices.add(u); });
    notifyChanged();
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
    // 噪音的座標帶給同一次採樣的振動（見 smartparse.js 的說明）。
    // 一定要在下面分測站之前做——測站設定的欄位是拿列上的值去帶的。
    perCategory[catKey].filledFromNoise = fillVibrationSharedFromNoise(rows, catKey);
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
  // Each queued category is its own import. Without clearing these, category 2
  // inherited category 1's 期別 label — a Q3 地質 report filed as Q2, which then
  // vanished from a Q3 export — and category 1's 地點/測項 column filter, which left
  // the detail list rendering empty.
  state.importPeriod = '';
  state.limitFixMode = null;
  state.rowDetailColumnFilters = {};
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
  state.limitFixMode = null; // "what to do with a non-numeric 檢測極限" — asked per import
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
        aggregate.filledFromNoise = fillVibrationSharedFromNoise(aggregate.rows, catKey);
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

  // Keeps the confirm button's count honest as items/rows are ticked and unticked,
  // the same way the report-preview screen does.
  let lastMappedRows = [];
  const updateGenericCount = () => {
    const n = filterRowsBySelection(lastMappedRows, cat.itemField).length;
    document.getElementById('btnImportConfirm').textContent = `確認匯入 ${n} 筆資料`;
  };

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
    lastMappedRows = mappedRows;
    renderRowDetailTable(document.getElementById('genericImportRowDetailWrap'), itemFilteredRows, cat, updateGenericCount);
    updateGenericCount();
    renderLimitWarning(document.getElementById('genericImportLimitWarning'),
      filterRowsBySelection(mappedRows, cat.itemField), cat);

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

  /*
   * 採樣日期不屬於本次期別的列，同樣預設不勾（見 applyPeriodExclusions 的說明）。
   * 順序有講究：要先讓 renderPeriodPicker 把 state.importPeriod 定下來
   * （它會在沒有值的時候填入猜出來的期別），才算得出哪些列不屬於這一期；
   * 而提示文字又要講出「已預設排除幾筆」，所以算完之後再畫一次選擇器。
   */
  state.autoExcludedByPeriod = new Set();
  state.manualRowChoices = new Set(); // 新的一次解析＝重新開始，之前的手動勾選不再適用
  const onPeriodChanged = () => { applyPeriodExclusions(result.rows); updateCountsAndRowDetail(); };
  renderPeriodPicker('smartPeriodWrap', result.rows, onPeriodChanged); // 先定下 state.importPeriod
  applyPeriodExclusions(result.rows);                                  // 再算哪些列不屬於這一期
  renderPeriodPicker('smartPeriodWrap', result.rows, onPeriodChanged); // 重畫，讓提示講得出「已排除幾筆」

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
    // Only warn about rows that are actually going to be imported — an offending
    // 檢測極限 sitting on a test item the person has unticked is not their problem.
    renderLimitWarning(document.getElementById('smartImportLimitWarning'),
      filterRowsBySelection(result.rows, cat.itemField), cat);
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

  /*
   * 從噪音帶給振動的欄位要**講出來**。
   * 沒有講出來的自動填值，和悄悄改掉使用者的資料只差一線——
   * 他要能知道振動那幾筆的座標與機構代碼是誰給的，才有辦法核對。
   */
  const existingFilledNotice = document.getElementById('smartImportFilledFromNoise');
  if (existingFilledNotice) existingFilledNotice.remove();
  const filled = result.filledFromNoise;
  if (filled && filled.rows > 0) {
    const detail = Object.entries(filled.byField).map(([f, n]) => `${f} ${n} 筆`).join('、');
    const note = document.createElement('div');
    note.id = 'smartImportFilledFromNoise';
    note.className = 'warning';
    note.innerHTML = `ℹ️ 報告的振動那一半通常不寫座標與檢測機構，已依「同一測站、同一次採樣」`
      + `<strong>從噪音那一半帶入 ${filled.rows} 筆振動資料</strong>（${escapeHtml(detail)}）。`
      + `<br><span class="hint">只填原本空白的格子，不會覆蓋報告本身已經有的值。請在下方詳細資料列表核對一下。</span>`;
    document.getElementById('smartImportItemsWrap').after(note);
  }

  /*
   * 匯入進來就已經違反連動規則的資料，在預覽畫面就講（v4.42）。
   * 別家公司交來的檔案、報告本身填錯的，最好在按下「確認匯入」之前就知道，
   * 而不是等一季之後匯出被退件才回頭找。
   */
  const existingRuleWarning = document.getElementById('smartImportRuleWarning');
  if (existingRuleWarning) existingRuleWarning.remove();
  const ruleBad = findNoiseVibRuleViolations(result.rows, cat);
  if (ruleBad.length > 0) {
    const total = ruleBad.reduce((n, [, c]) => n + c, 0);
    const warn = document.createElement('div');
    warn.id = 'smartImportRuleWarning';
    warn.className = 'warning';
    warn.innerHTML = `⚠️ 這份資料有 <strong>${total} 筆</strong>的管制標準／管制區和檢測類別對不起來：<br>`
      + ruleBad.slice(0, 6).map(([k, n]) => {
        const [category, field, value] = k.split('｜');
        return `　・檢測類別「${escapeHtml(category)}」的「${escapeHtml(field)}」＝「${escapeHtml(value)}」：${n} 筆`;
      }).join('<br>')
      + `${ruleBad.length > 6 ? `<br>　…另有 ${ruleBad.length - 6} 種` : ''}`
      + `<br><span class="hint">官方規定：檢測類別是「振動」時管制標準與管制區必須填「無」，是噪音時不能填「無」`
      + `（115 年 7 月起加強檢核）。匯入後在表格中會用紅框標示，可用欄位篩選 ▾ 快速找到並修正。</span>`;
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

  const datesByLoc = {};  // confirmedLoc -> Set of 日期(起) present in THIS import
  const itemsByLoc = {}; // confirmedLoc -> Map(identityKey -> count)
  const siteKeysByLoc = {}; // confirmedLoc -> [siteKey, ...] (may span multiple raw sites)
  siteEntries.forEach(([key, site]) => {
    const saved = savedAliases[siteAliasKey(site, result, cat)] || {};
    const confirmedLoc = saved[locField] || fuzzyLocationOverrides[key] || site.rawLocation;
    if (!itemsByLoc[confirmedLoc]) { itemsByLoc[confirmedLoc] = new Map(); siteKeysByLoc[confirmedLoc] = []; datesByLoc[confirmedLoc] = new Set(); }
    site.rowIndices.forEach(i => {
      const idKey = itemIdentityKey(result.rows[i], cat);
      const m = itemsByLoc[confirmedLoc];
      m.set(idKey, (m.get(idKey) || 0) + 1);
      if (result.rows[i]['日期(起)']) datesByLoc[confirmedLoc].add(result.rows[i]['日期(起)']);
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
    // TWO VERY DIFFERENT SITUATIONS, WHICH USED TO BE TREATED AS ONE
    //
    // (a) The item is ENTIRELY ABSENT this time — history has it at this site, this
    //     report has none at all. That is the case worth flagging and ticking by
    //     default: an item genuinely dropped out of the report.
    //
    // (b) The item IS here, just fewer times than history remembers. Almost always
    //     this only means the two imports cover a different number of sampling
    //     occasions — comparing a whole quarter's completed filing (Jan + Feb + Mar,
    //     3 rows per item) against a single month's report (1 row) reported every
    //     item as "缺 2 筆" and, because suggestions were ticked by default,
    //     silently appended two blank rows per item on confirm. So (b) is now
    //     surfaced UNTICKED, with the actual dates on both sides shown as evidence,
    //     and it says outright that a single-month import is the usual explanation.
    const currentDates = [...(datesByLoc[confirmedLoc] || [])].sort();
    const missing = historyEntriesFor(siteHistory[confirmedLoc])
      .map(([identityKey, entry]) => {
        const historicalCount = entry.count || 1; // tolerate pre-count history entries
        const currentCount = currentCounts.get(identityKey) || 0;
        if (historicalCount - currentCount <= 0) return null;
        // How many blank rows to offer. For an item that's entirely absent, that is
        // one per sampling occasion THIS import actually covers — not however many
        // the history happened to hold. A single-month report missing an item needs
        // one blank row; a quarter's history having three of them doesn't change
        // that. (A same-quarter 平日/假日 pair still gets 2, because this import
        // genuinely has 2 sampling dates.)
        const occasionsNow = Math.max(1, currentDates.length);
        const missingCount = currentCount === 0
          ? Math.max(1, Math.min(historicalCount, occasionsNow))
          : historicalCount - currentCount;
        return {
          identityKey, ...entry, missingCount, currentCount, historicalCount,
          historicalDates: Array.isArray(entry.dates) ? entry.dates.slice().sort() : [],
          currentDates,
          // absent = nothing of this item at all this time; shortfall = fewer than before
          absent: currentCount === 0,
        };
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
        📋 系統比對過去記錄，列出以下差異。若要一併新增空白列（檢測方法／單位／項目名稱及對應的檢測類別會依過去記錄先幫您填好，檢測數值請自行輸入），請勾選：
        <div class="hint" style="margin:6px 0 0 0">
          ・<strong>本次完全沒有這個測項</strong>：預設<strong>已勾選</strong>（過去有測、這次報告裡整個不見了，通常需要補）。<br>
          ・<strong>本次有、但比過去少幾筆</strong>：預設<strong>不勾選</strong>。最常見的原因是兩邊涵蓋的採樣次數不同——例如過去比對的是一整季完成版（1、2、3 月各一次共 3 筆），而這次只匯入單月報告（1 筆），並不是真的少了資料。每個項目下方都會列出兩邊實際的採樣日期供您判斷。
        </div>
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
            // A group is ticked only if something inside it is — with shortfalls now
            // unticked, a group of nothing but shortfalls must not look "selected".
            const defaultChecked = !s.entirelyAbsent && s.missingItems.some(m => m.absent);
            return `
            <div class="missing-item-loc-group" data-search-text="${escapeAttr(s.location.toLowerCase())}" style="margin-top:6px${s.entirelyAbsent ? ';padding:6px 8px;background:#fff6e0;border-radius:6px' : ''}">
              <label style="font-weight:700">
                <input type="checkbox" class="missing-item-group" data-group-idx="${i}" ${defaultChecked ? 'checked' : ''}> ${escapeHtml(s.location)}${s.entirelyAbsent ? ' <span class="hint">（本次報告完全沒有這個測站——若這是不同專案/不相關的批次資料，請保持不勾選；確定要帶回本測站的舊資料才勾選。建議新增的資料日期需要您自行填寫）</span>' : ''}
              </label>
              <div style="margin-left:22px">
                ${s.missingItems.map((m) => {
                  const { identityKey, itemName, timeSegment, category, snapshot, missingCount } = m;
                  const displayName = timeSegment ? `${itemName}（${timeSegment}）` : itemName;
                  // A shortfall (the item IS in this report, just fewer times) is
                  // untick-by-default and explains itself with the actual dates on
                  // both sides — that is what tells apart "a reading really went
                  // missing" from "this file is one month, that history was a
                  // whole quarter".
                  const itemChecked = s.entirelyAbsent ? false : !!m.absent;
                  const fmtDates = (ds) => (ds || []).map(d => toDateDisplayValue(d) || d).join('、');
                  let countNote = '';
                  if (!m.absent && m.historicalCount !== undefined) {
                    const histDates = fmtDates(m.historicalDates);
                    const curDates = fmtDates(m.currentDates);
                    countNote = `<span class="shortfall-note">（本次有 ${m.currentCount} 筆${curDates ? `：${escapeHtml(curDates)}` : ''}；`
                      + `過去記錄有 ${m.historicalCount} 筆${histDates ? `：${escapeHtml(histDates)}` : ''}。`
                      + `<strong>若本次只是匯入單月／單次報告，其他月份稍後才會匯入，請不要勾選</strong>——`
                      + `勾選會新增 ${missingCount} 筆空白列。）</span>`;
                  } else if (missingCount > 1) {
                    countNote = `<strong>（缺 ${missingCount} 筆，將新增 ${missingCount} 筆空白列）</strong>`;
                  }
                  const mem = itemMemoryForSuggestion[itemName];
                  const methodNote = snapshot?.[cat.methodField] || mem?.method;
                  const unitNote = snapshot?.[cat.unitField] || mem?.unitCode;
                  const memParts = [category ? `檢測類別：${category}` : '', methodNote, unitNote ? `單位代碼${unitNote}` : ''].filter(Boolean);
                  const memNote = memParts.length ? `（已記憶：${memParts.join('，')}）` : '（無先前記憶的方法/單位，需另外補上）';
                  return `<label style="display:block">
                    <input type="checkbox" class="missing-item-single" data-group-idx="${i}" data-identity-key="${escapeAttr(identityKey)}" ${itemChecked ? 'checked' : ''}>
                    ${escapeHtml(displayName)} ${countNote} <span class="hint">${escapeHtml(memNote)}</span>
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
              return `<td><select data-site-field="${f.key}" data-seeded="${escapeAttr(val)}">${opts}</select></td>`;
            }
            return `<td><input type="text" data-site-field="${f.key}" data-seeded="${escapeAttr(val)}" value="${escapeAttr(val)}"></td>`;
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
 *  監測時段 when the category has that field (see hasTimeSegmentField).
 *
 *  夜間振動的音源發聲特性在 v4.29 以前被寫成 Lvd(10)（正確應為 Lvn(10)）。
 *  這裡把那個**錯誤組合**正規化，讓舊資料與本版新匯入的資料算同一個測項——
 *  否則舊記憶裡的 Lvd(10)::夜間 會被當成「這個測項不見了」而預設打勾補一筆空白列，
 *  重新匯入同一份報告也會因為配不到舊列而多附加一筆。詳見 schema.js。 */
function itemIdentityKey(row, cat) {
  if (!hasTimeSegmentField(cat)) return row[cat.itemField];
  const seg = row['監測時段'] || '';
  return `${canonicalVibItemName(row[cat.itemField], seg)}::${seg}`;
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
    // 記下來的名稱也要正規化。這份 itemName 是日後「補回缺少的測項」時用來重建
    // 資料列的（buildSuggestedRows），留著舊的 Lvd(10) 會把錯誤再種回去一次。
    itemName: hasTimeSegmentField(cat)
      ? canonicalVibItemName(r[itemField], r['監測時段'] || '')
      : r[itemField],
    timeSegment: hasTimeSegmentField(cat) ? (r['監測時段'] || '') : '',
    itemCategory: r['檢測類別'] || '',
    date: r['日期(起)'] || '',
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
  // The identity of a row must include 監測時段 wherever the category has one.
  // Without it, noise's 日間/晚間/夜間 readings were indistinguishable: the parser
  // gives all three the same date, the same 時間(起) ("00:00:00" for a 24-hour
  // report), the same location and the same 音源發聲特性. Re-importing a file then
  // matched all three to the SAME existing row and wrote them into that one slot in
  // turn, so the daytime reading was silently replaced by the night one — different
  // statutory limits, wrong number filed, and the only feedback was "更新 N 筆".
  const keyOf = (r) => [r['日期(起)'], r['時間(起)'], r[locField], itemIdentityKey(r, cat)].join('\u0001');
  const identityKeys = new Set(['日期(起)', '時間(起)', locField, itemField,
    ...(hasTimeSegmentField(cat) ? ['監測時段'] : [])]);

  // A queue per key, not a single index: when a site legitimately has two rows with
  // the same identity (平日 and 假日 sampled on one date), each candidate consumes a
  // different existing row instead of both piling onto the first.
  const existingByKey = new Map();
  existingRows.forEach((r, idx) => {
    const k = keyOf(r);
    if (!existingByKey.has(k)) existingByKey.set(k, []);
    existingByKey.get(k).push(idx);
  });

  const brandNew = [];
  const conflicts = [];

  candidateRows.forEach(candidate => {
    const key = keyOf(candidate);
    const bucket = existingByKey.get(key);
    if (!bucket || bucket.length === 0) { brandNew.push(candidate); return; }
    const existingIdx = bucket.shift();
    const existingRow = existingRows[existingIdx];
    const diffFields = [];
    cat.fields.forEach(f => {
      /*
       * 識別欄位原則上不列進差異（配對成功就代表它們一樣），但有一個例外：
       * 音源發聲特性經過 canonicalVibItemName 正規化之後，「夜間的 Lvd(10)」與
       * Lvn(10) 會配到同一列——**字面上卻不一樣**。不把它列出來，使用者在衝突
       * 視窗看到的差異表會說「只差小數點寫法」，按下確定卻連測項名稱一起被換掉。
       * 要換可以，但必須寫在畫面上讓他看見。
       */
      const isIdentity = identityKeys.has(f.key);
      if (isIdentity && f.key !== itemField) return;
      const oldVal = existingRow[f.key] || '';
      const newVal = candidate[f.key] || '';
      if (oldVal === newVal) return;
      /*
       * 「39.2」與「39.20」是同一個數字，只是寫法不同（v4.30 起監測數值一律補成
       * 兩位小數；舊資料、完成版範本匯入的資料則不一定）。把它算成差異的話，
       * 每一筆舊資料在下次匯入時都會跳出一列衝突，而衝突視窗承諾的是
       * 「內容完全相同的部分會自動忽略」。數字相同就不算差異。
       */
      if (sameNumericValue(oldVal, newVal)) return;
      diffFields.push({ key: f.key, label: f.label, type: f.type, oldVal, newVal });
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
      // The template is ANOTHER item's row at the same site — good for shared,
      // site-level fields (coordinates, 管制標準), wrong for anything item-specific.
      // Overlay this item's own remembered snapshot so it doesn't inherit, say,
      // 水溫's NIEA W217 and unit 4. The preview literally promises "檢測方法／單位…
      // 會依過去記錄先幫您填好"; before this it showed one thing and wrote another.
      Object.entries(snapshot || {}).forEach(([k, v]) => { if (v) newRow[k] = v; });
      // If history had no method/unit for this item, blank the template's so the
      // item-memory fallback below can actually fire instead of being shadowed.
      if (cat.methodField && !(snapshot || {})[cat.methodField]) newRow[cat.methodField] = '';
      if (cat.unitField && !(snapshot || {})[cat.unitField]) newRow[cat.unitField] = '';
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
    // 比較關係 and 檢測極限 describe the TEMPLATE row's measurement, not this one —
    // carrying them over would tag a blank row as ND, or give it another item's
    // detection limit.
    ['日期(起)', '時間(起)', '日期(迄)', '時間(迄)', '檢測數值', '監測數值', '比較關係', '檢測極限']
      .forEach(k => { if (k in newRow) newRow[k] = ''; });
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
  /*
   * 從噪音帶給振動的欄位一定要講出來（見 smartparse.js）。
   * 智慧判讀那條路徑在匯入預覽上已經先講過一次，這裡是「已是官方格式」
   * 那條路徑的唯一機會——它沒有預覽提示區。
   */
  const filledNote = state.filledFromNoiseNote;
  state.filledFromNoiseNote = '';

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
  alert(`已匯入「${cat.label}」：${summary.join('、') || '沒有變更'}。`
    + `${filledNote ? `\n\n${filledNote}` : ''}`
    + `\n請於表格中核對內容是否正確，特別是尚未有測站設定的欄位。`);
}

/**
 * 找出「晚間 ＋ 振動 ＋ Lv10」的資料列。
 *
 * 官方的振動音源發聲特性只有 Lvd(10)（日間）與 Lvn(10)（夜間）兩個代碼，**沒有晚間**
 * ——振動參考日本振動規制法，只分日夜兩段。但報告上的「測定時間」是用噪音那套
 * 三段規則（日／晚／夜）判的，所以營建工程若在 18:00~21:59 之間量測，
 * 振動列會落在一個對振動不存在的時段。
 *
 * 程式預設把它併入夜間（Lvn(10)），但**不替使用者默默決定**：匯入時跳出來問一次。
 * 營建工程本來就很少這麼晚量測，所以這個提醒不會常出現；真的出現時，
 * 那一筆多半值得看一眼。
 */
function findEveningVibrationRows(rows, cat) {
  return rows
    .map((r, idx) => ({ r, idx }))
    .filter(({ r }) =>
      r['檢測類別'] === '振動'
      && r['監測時段'] === '晚間'
      && (r[cat.itemField] === VIB_LV10_DAY || r[cat.itemField] === VIB_LV10_NIGHT));
}

/**
 * 問使用者晚間那幾筆振動要算 Lvn(10) 還是 Lvd(10)，並就地改好 rows。
 * 回傳 false 代表使用者按了取消，整個匯入應該中止。
 */
function askEveningVibrationItem(rows, cat) {
  const hits = findEveningVibrationRows(rows, cat);
  if (!hits.length) return true;
  const times = [...new Set(hits.map(({ r }) => `${r['日期(起)'] || ''} ${r['時間(起)'] || ''}`.trim()))];
  const locs = [...new Set(hits.map(({ r }) => r[cat.locationField] || '（未填地點）'))];
  const answer = confirm(
    '【晚間量測的振動要算日間還是夜間？】\n\n'
    + '官方的振動音源發聲特性只有兩個代碼：Lvd(10)（日間）與 Lvn(10)（夜間），沒有「晚間」。\n'
    + '振動參考日本振動規制法，本來就只分日夜兩段；「晚間」是噪音那套三段分法。\n\n'
    + `這次匯入有 ${hits.length} 筆振動落在 18:00~21:59：\n`
    + `　時間：${times.slice(0, 4).join('、')}${times.length > 4 ? ` 等 ${times.length} 個時段` : ''}\n`
    + `　測站：${locs.slice(0, 4).join('、')}${locs.length > 4 ? ` 等 ${locs.length} 個` : ''}\n\n`
    + '要用哪一個？\n\n'
    + '　【確定】＝ Lvn(10) 夜間（預設，日本振動規制法多數地區 19:00 後即為夜間）\n'
    + '　【取消】＝ Lvd(10) 日間',
  );
  const item = answer ? VIB_LV10_NIGHT : VIB_LV10_DAY;
  hits.forEach(({ idx }) => { rows[idx][cat.itemField] = item; });
  showToast(`晚間的 ${hits.length} 筆振動已設為 ${item}。之後可在表格裡逐列調整。`);
  return true;
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

    /*
     * ⚠️ 一個「測站」可能橫跨好幾次採樣。
     *
     * 這裡的分組鍵值是「測點編號或原始地點名稱」，**沒有帶採樣日期**，而一次匯入
     * 好幾份月報是系統明確支援的用法。表格裡那一列的值是拿第一份檔案的資料當種子，
     * 舊寫法再無條件 Object.assign 到整組——於是 3 月、4 月、5 月被當成同一組，
     * 4 月與 5 月各自報告裡的座標會被 3 月的蓋掉。
     * （系統其他地方的座標同步都要求「同地點 ＋ 同日期」，註解也寫明
     *   同一個測站不同次採訪的座標可能真的不一樣。）
     *
     * 改成分兩種情況：
     *   ・使用者**真的在表格裡改過**那一格（值 ≠ data-seeded）→ 是明確指令，整組套用。
     *   ・沒改過 → 只補「那一列本來就空白」的欄位，已經有值的一律不動。
     * 單月匯入（最常見）行為完全不變，因為那時整組本來就只有一次採樣。
     */
    const editedFields = new Set();
    tr.querySelectorAll('[data-site-field]').forEach(el => {
      if (el.value !== (el.dataset.seeded ?? '')) editedFields.add(el.dataset.siteField);
    });
    site.rowIndices.forEach(idx => {
      const row = result.rows[idx];
      Object.entries(overrides).forEach(([key, value]) => {
        if (editedFields.has(key)) { row[key] = value; return; }
        if (String(row[key] ?? '').trim() === '') row[key] = value;
      });
    });
  });
  try {
    DataStore.saveSiteAliases(project.id, catKey, savedAliases);
  } catch (err) {
    console.error(err);
    alert('瀏覽器的儲存空間已滿，這批資料沒有存進去。\n\n請先用右上角「匯出備份」把現有資料存成檔案，再刪除不需要的計畫，然後重新匯入。');
    closeImportModal();
    return;
  }

  let selectedRows = filterRowsBySelection(result.rows, cat.itemField);
  const suggestedRows = buildSuggestedRows(project, catKey, cat, result);
  selectedRows = selectedRows.concat(suggestedRows);
  applyLimitFixMode(selectedRows, cat); // 檢測極限 must be a bare number — see renderLimitWarning
  if (selectedRows.length === 0) { alert('目前沒有勾選任何監測項目，請至少勾選一項再匯入。'); return; }
  // 晚間量測的振動：官方沒有晚間代碼，問過使用者再寫進去（見 askEveningVibrationItem）
  if (!askEveningVibrationItem(selectedRows, cat)) return;

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
    runImportCommit(() => finalizeImportCommit(project, catKey, cat, brandNew, updates, assignBatchId, 'smart'));
  };

  if (conflicts.length === 0) {
    if (brandNew.length === 0) { alert('這批資料跟現有資料完全相同，沒有新增或需要處理的內容。'); return; }
    proceed(null);
  } else {
    openConflictResolutionModal(conflicts, ({ useNew }) => proceed(useNew));
  }
}

/** Wraps a commit so a full localStorage reports itself instead of throwing into the
 *  void. Without this, a QuotaExceededError escaped mid-save: nothing was written,
 *  the modal just sat there, and the person walked away believing the import worked. */
function runImportCommit(fn) {
  try {
    fn();
  } catch (err) {
    console.error(err);
    const quota = err && (err.name === 'QuotaExceededError' || /quota/i.test(err.message || ''));
    alert(quota
      ? '瀏覽器的儲存空間已滿，這批資料沒有存進去。\n\n請先用右上角「匯出備份」把現有資料存成檔案，再刪除不需要的計畫，然後重新匯入。'
      : '匯入時發生錯誤，這批資料沒有存進去：' + (err && err.message ? err.message : err));
    closeImportModal();
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
  // 匯入已是官方格式的檔案（上一季的完成版、別家交來的填好檔）走的是這條路徑。
  // 那些檔案一樣可能只有噪音那一半有座標，所以這裡也要帶一次。
  const filledFromNoise = fillVibrationSharedFromNoise(newRows, catKey);
  state.filledFromNoiseNote = filledFromNoise.rows > 0
    ? `ℹ️ 報告的振動那一半通常不寫座標與檢測機構，已依「同一測站、同一次採樣」從噪音那一半`
      + `帶入 ${filledFromNoise.rows} 筆振動資料（`
      + `${Object.entries(filledFromNoise.byField).map(([f, n]) => `${f} ${n} 筆`).join('、')}）。`
      + `\n只填原本空白的格子，不會覆蓋檔案本身已經有的值——請核對一下。`
    : '';
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
  applyLimitFixMode(selectedRows, cat); // 檢測極限 must be a bare number — see renderLimitWarning
  if (selectedRows.length === 0) { alert('目前沒有勾選任何監測項目，請至少勾選一項再匯入。'); return; }
  // 對應欄位這條路徑同樣要問——完成版申報檔重新匯入時也可能帶著晚間的振動列
  if (!askEveningVibrationItem(selectedRows, cat)) return;

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
    runImportCommit(() => finalizeImportCommit(project, catKey, cat, brandNew, updates, assignBatchId, 'generic'));
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
/**
 * Wiring first, rendering last.
 *
 * init() used to render before attaching a single listener and had no try/catch, so
 * any exception thrown while drawing (a malformed value in localStorage was enough)
 * aborted it and left a page that LOOKS normal but where every button is inert —
 * including 匯出備份, so the data couldn't even be rescued. Rendering last, inside a
 * guard, means the worst case is a blank content area on a fully working page.
 */
/*
 * 儲存空間滿的時候，不管是哪一條路徑寫入失敗，都要看得到訊息。
 * 用節流避免一次操作寫好幾個 key 時彈出一整排視窗。
 */
let lastStorageAlertAt = 0;
DataStore.onStorageError = (err) => {
  console.error(err);
  const now = Date.now();
  if (now - lastStorageAlertAt < 3000) return;
  lastStorageAlertAt = now;
  alert(
    '⚠️ 這次的修改沒有存進去——瀏覽器的儲存空間已滿。\n\n'
    + '畫面上顯示的是您剛才輸入的內容，但實際存下來的還是舊的。\n\n'
    + '請先用右上角「匯出備份」把現有資料存成檔案，再刪除不需要的計畫或季度，'
    + '然後重新整理頁面確認。'
  );
};

function init() {
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
  document.getElementById('btnCommonItemsCancel').addEventListener('click', () => {
    document.getElementById('commonItemsModal').classList.add('hidden');
    const popup = document.getElementById('colFilterPopup');
    if (popup) popup.classList.add('hidden');
  });
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
        // 期別要跟著（理由見 currentPeriodForNewRow）——沒有的話這些列會變成
        // 「未標示期別」，依季度篩選或匯出時整批被漏掉。
        const blank = { _period: currentPeriodForNewRow(project, catKey) };
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
  // #colFilterPopup is a single shared body-level dropdown. Closing a modal that had
  // one open used to leave it floating over the data grid, covering controls, with no
  // way to dismiss it — so every modal-close path hides it too.
  const hideSharedPopup = () => {
    const popup = document.getElementById('colFilterPopup');
    if (popup) popup.classList.add('hidden');
  };
  document.querySelectorAll('.modal-overlay').forEach(overlay => {
    if (overlay.id === 'importModal') return;
    overlay.addEventListener('click', (e) => {
      if (e.target === overlay) { overlay.classList.add('hidden'); hideSharedPopup(); }
    });
  });
  document.addEventListener('keydown', (e) => {
    if (e.key !== 'Escape') return;
    hideSharedPopup();
    document.querySelectorAll('.modal-overlay').forEach(overlay => {
      if (overlay.id === 'importModal') return;
      if (!overlay.classList.contains('hidden')) overlay.classList.add('hidden');
    });
  });

  // Render LAST, and never let a drawing error take the listeners down with it.
  try {
    renderVersionBadge();
    renderProjectList();
    renderContent();
  } catch (err) {
    console.error('初始畫面繪製失敗', err);
    const content = document.getElementById('content');
    if (content) {
      content.innerHTML = '<div class="empty-state"><p>⚠️ 讀取本機資料時發生問題，畫面無法完整顯示。</p>'
        + '<p>其他功能仍可使用——建議先點右上角「匯出備份」保存目前資料。</p></div>';
    }
  }
}

document.addEventListener('DOMContentLoaded', init);
