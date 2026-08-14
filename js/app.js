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
  const project = getCurrentProject();
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

function escapeHtml(s) {
  return String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}
function escapeAttr(s) { return escapeHtml(s); }

function toDateInputValue(v) {
  if (!v) return '';
  const s = String(v);
  const m = s.match(/(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (m) return `${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}`;
  return '';
}
/** ISO "YYYY-MM-DD" (internal storage/export format, unchanged) -> "YYYY/MM/DD" for display. */
function toDateDisplayValue(v) {
  const iso = toDateInputValue(v);
  return iso ? iso.replace(/-/g, '/') : '';
}
/** Accepts "2026/5/12", "2026-5-12", "20260512" typed free-hand and normalizes to the
 *  canonical ISO "YYYY-MM-DD" used for storage/export/date-math. Returns '' if it
 *  doesn't look like a date at all, rather than silently keeping garbage. */
function normalizeDateString(raw) {
  const s = String(raw ?? '').trim();
  if (s === '') return '';
  let m = s.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/);
  if (m) return `${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}`;
  m = s.match(/^(\d{4})(\d{2})(\d{2})$/); // 20260512
  if (m) return `${m[1]}-${m[2]}-${m[3]}`;
  return s; // unrecognized — leave as typed so the person can see and fix it
}
function toTimeInputValue(v) {
  if (!v) return '';
  const s = String(v);
  const m = s.match(/(\d{1,2}):(\d{2})(:(\d{2}))?/);
  if (m) return `${m[1].padStart(2, '0')}:${m[2]}:${m[4] || '00'}`;
  return '';
}
/** HH:MM:SS (internal storage/export format, unchanged) -> HH:MM for display — the
 *  official template's own time format is "h:mm" with no seconds shown anyway. */
function toTimeDisplayValue(v) {
  const full = toTimeInputValue(v);
  return full ? full.slice(0, 5) : '';
}
/** Accepts "1430", "14:30", "14:30:00", "143000" typed free-hand and normalizes to HH:MM:SS.
 *  Returns the original trimmed string unchanged if it doesn't look like a time at all,
 *  so an in-progress or unrecognized entry isn't silently discarded. */
function normalizeTimeString(raw) {
  const s = String(raw ?? '').trim();
  if (s === '') return '';
  let m = s.match(/^(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?$/);
  if (m) return `${m[1].padStart(2, '0')}:${m[2].padStart(2, '0')}:${(m[3] || '00').padStart(2, '0')}`;
  m = s.match(/^(\d{2})(\d{2})(\d{2})$/); // 143000
  if (m) return `${m[1]}:${m[2]}:${m[3]}`;
  m = s.match(/^(\d{1,2})(\d{2})$/); // 1430 or 930
  if (m) return `${m[1].padStart(2, '0')}:${m[2]}:00`;
  return s;
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

// ---------- project list ----------
function renderProjectList() {
  const list = document.getElementById('projectList');
  const projects = DataStore.getProjects();
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

function selectProject(id) {
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
    <div class="toolbar">
      <div class="toolbar-left">
        <button class="btn btn-primary btn-sm" id="btnImport">📥 匯入資料（Excel/PDF）</button>
        <button class="btn btn-ghost btn-sm" id="btnAddRow">＋ 新增一筆</button>
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
        <button class="btn btn-danger btn-sm" id="btnClearCat" ${allRows.length === 0 ? 'disabled' : ''}>🗑 清空此類別</button>
      </div>
      <div class="row-count" id="rowCountDisplay">共 ${displayRows.length} 筆資料${activePeriod ? `（篩選中，全部共 ${allRows.length} 筆）` : ''}</div>
    </div>
    <div class="toolbar bulk-toolbar hidden" id="bulkToolbar">
      <span id="bulkSelCount">已選取 0 筆</span>
      <button class="btn btn-danger btn-sm" id="btnBulkDelete">🗑 刪除已選取</button>
      <button class="btn btn-ghost btn-sm" id="btnBulkClear">取消選取</button>
    </div>
    <div class="table-wrap">
      <table class="data-grid">
        <thead><tr>
          <th class="col-check"><input type="checkbox" id="checkAllRows" ${displayRows.length === 0 ? 'disabled' : ''}></th>
          <th>操作</th>
          <th>#</th>
          ${cat.fields.map(f => {
            const activeFilter = state.columnFilters[catKey]?.[f.key];
            const filterActive = activeFilter && activeFilter.size > 0;
            const sortState = state.columnSort[catKey];
            const isSorted = sortState && sortState.fieldKey === f.key;
            const sortIcon = isSorted ? (sortState.direction === 'asc' ? ' ▲' : ' ▼') : '';
            return `<th${f.key === cat.itemField ? ' class="col-item"' : f.key === cat.locationField ? ' class="col-loc"' : ''}${f.help ? ` title="${escapeAttr(f.help)}"` : ''}>
              <span class="th-label th-sortable" data-sort-field="${escapeAttr(f.key)}" title="點擊依此欄位排序">${escapeHtml(f.label)}${f.required ? '<span class="req">＊</span>' : ''}${f.help ? ' ℹ️' : ''}${sortIcon}</span>
              <button class="col-filter-btn${filterActive ? ' col-filter-active' : ''}" data-field-key="${escapeAttr(f.key)}" title="篩選「${escapeAttr(f.label)}」">▾</button>
            </th>`;
          }).join('')}
        </tr></thead>
        <tbody id="gridBody">${displayEntries.map(({ row, idx }) => rowHtml(cat, row, idx)).join('')}</tbody>
      </table>
    </div>
    ${allRows.length === 0 ? '<p class="hint" style="margin-top:10px">尚無資料。可點「匯入資料」上傳該類別的檢測結果檔案，或「新增一筆」手動輸入。</p>' : ''}
    ${allRows.length > 0 && displayRows.length === 0 ? '<p class="hint" style="margin-top:10px">此期別目前沒有資料。</p>' : ''}
  `;

  document.getElementById('btnImport').addEventListener('click', () => openImportModal(catKey));
  document.getElementById('btnAddRow').addEventListener('click', () => addEmptyRow(project, catKey));
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
    if (!confirm(`確定要清空「${cat.label}」的全部 ${allRows.length} 筆資料嗎？（含所有期別）此操作無法復原。`)) return;
    DataStore.clearData(project.id, catKey);
    state.periodFilter[catKey] = '';
    renderContent();
  });

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
      renderContent();
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

  document.querySelectorAll('.col-filter-btn').forEach(btn => {
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
    renderContent();
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
      if (!confirm(`確定要刪除已選取的 ${indices.size} 筆資料嗎？此操作無法復原。`)) return;
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
        if (!confirm(`確定要刪除這批匯入資料嗎？（來源：${row.sourceLabel}，共 ${row.rowCount} 筆）此操作無法復原。`)) return;
        DataStore.deleteImportBatch(project.id, catKey, batchId);
        openBatchHistoryModal(project, catKey); // refresh list in place
        renderContent();
      });
    });
  }
  document.getElementById('batchHistoryModal').classList.remove('hidden');
}

function rowHtml(cat, row, idx) {
  const cells = cat.fields.map(f => `<td${f.key === cat.itemField ? ' class="col-item"' : f.key === cat.locationField ? ' class="col-loc"' : ''}>${fieldControlHTML(f, row[f.key], `data-row="${idx}"`)}</td>`).join('');
  return `<tr data-row="${idx}"><td class="col-check"><input type="checkbox" class="row-check" data-row="${idx}"></td><td class="col-actions"><button class="row-del-btn" data-row="${idx}" title="刪除此列">🗑</button></td><td>${idx + 1}</td>${cells}</tr>`;
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
    DataStore.saveData(project.id, catKey, rows);
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
    if (requireDate && r['日期(起)'] !== source['日期(起)']) return false;
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
      `偵測到同一份檔案、同一天（${source['日期(起)']}）、同一個測站「${source[locField]}」${source['檢測類別'] ? `、同為「${source['檢測類別']}」` : ''}還有 ${matches.length} 筆其他資料。\n` +
      `是否要將這些資料的座標一併同步更新為與這一筆相同？\n\n` +
      `（選擇「取消」則只修改目前這一筆，其他資料維持原狀。不同採樣日期或不同檢測類別的資料不會被同步。）`
    );
    if (!ok) return false;
    matches.forEach(({ r }) => { COORD_FIELDS.forEach(f => { r[f] = source[f]; }); });
    DataStore.saveData(project.id, catKey, rows);
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
      `偵測到同一份檔案、同一天（${source['日期(起)']}）、同一個測站「${source[locField]}」還有 ${matches.length} 筆其他資料。\n` +
      `是否要將這些資料的檢測類別一併同步更新為「${source['檢測類別']}」？\n\n` +
      `（選擇「取消」則只修改目前這一筆，其他資料維持原狀。）`
    );
    if (!ok) return false;
    matches.forEach(({ r }) => { r['檢測類別'] = source['檢測類別']; });
    DataStore.saveData(project.id, catKey, rows);
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
      `偵測到同一份檔案、同一天（${source['日期(起)']}）、同一個測站「${source[locField]}」${source['檢測類別'] ? `、同為「${source['檢測類別']}」` : ''}還有 ${matches.length} 筆其他資料。\n` +
      `是否要將這些資料的「${fieldLabel}」一併同步更新為「${source[fieldKey] || '（空白）'}」？\n\n` +
      `（選擇「取消」則只修改目前這一筆，其他資料維持原狀。）`
    );
    if (!ok) return false;
    matches.forEach(({ r }) => { r[fieldKey] = source[fieldKey]; });
    DataStore.saveData(project.id, catKey, rows);
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
    if (!confirm('確定要刪除這一列資料嗎？')) return;
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
  const rows = DataStore.getData(projectId, catKey);
  const locField = CATEGORIES[catKey].locationField;

  document.querySelectorAll('#coordSitesBody tr').forEach(tr => {
    const [loc, date] = tr.dataset.groupKey.split('\u0001');
    const values = {};
    tr.querySelectorAll('[data-coord-field]').forEach(el => { values[el.dataset.coordField] = el.value; });
    rows.forEach(row => {
      const rowLoc = (row[locField] || '').trim() || '（未命名測站）';
      const rowDate = row['日期(起)'] || '（未填日期）';
      if (rowLoc === loc && rowDate === date) Object.assign(row, values);
    });
  });

  DataStore.saveData(projectId, catKey, rows);
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

  document.querySelectorAll('#methodItemsBody tr').forEach(tr => {
    const item = tr.dataset.item;
    const values = {};
    tr.querySelectorAll('[data-method-field]').forEach(el => {
      const targetField = el.dataset.methodField === 'method' ? cat.methodField : cat.unitField;
      if (targetField) values[targetField] = el.value;
    });
    rows.forEach(row => {
      const rowItem = (row[itemField] || '').trim() || '（未命名項目）';
      if (rowItem === item) Object.assign(row, values);
    });
    const memFields = {};
    if (values[cat.methodField]) memFields.method = values[cat.methodField];
    if (cat.unitField && values[cat.unitField]) memFields.unitCode = values[cat.unitField];
    if (Object.keys(memFields).length) memoryUpdates[item] = memFields;
  });

  DataStore.saveData(projectId, catKey, rows);
  if (Object.keys(memoryUpdates).length) DataStore.updateItemMemory(projectId, catKey, memoryUpdates);
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
    state.currentProjectId = p.id;
  }
  closeProjectModal();
  renderProjectList();
  renderContent();
}
function deleteProjectFlow(project) {
  if (!confirm(`確定要刪除計畫「${project.code} ${project.name}」嗎？此操作將刪除該計畫所有已輸入的監測資料，且無法復原。`)) return;
  DataStore.deleteProject(project.id);
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
  rows.forEach(r => {
    const v = (r[itemField] || '').trim() || '（未標示）';
    counts[v] = (counts[v] || 0) + 1;
  });
  const items = Object.keys(counts);
  if (items.length === 0) { containerEl.innerHTML = ''; return; }
  if (!state.itemSelection) state.itemSelection = new Set(items);

  containerEl.innerHTML = `
    <p class="hint">此份報告偵測到以下監測項目，請勾選要匯入的項目（預設全選，可依實際需求取消勾選）：</p>
    <div class="item-checklist">
      ${items.map(item => `
        <label class="item-check">
          <input type="checkbox" data-item-check value="${escapeAttr(item)}" ${state.itemSelection.has(item) ? 'checked' : ''}>
          ${escapeHtml(item)} <span class="hint">(${counts[item]})</span>
        </label>
      `).join('')}
      <button type="button" class="btn btn-ghost btn-sm" id="btnItemSelectAll">全選</button>
      <button type="button" class="btn btn-ghost btn-sm" id="btnItemSelectNone">全不選</button>
    </div>
  `;
  containerEl.querySelectorAll('[data-item-check]').forEach(cb => {
    cb.addEventListener('change', () => {
      if (cb.checked) state.itemSelection.add(cb.value); else state.itemSelection.delete(cb.value);
      if (onChange) onChange();
    });
  });
  containerEl.querySelector('#btnItemSelectAll').addEventListener('click', () => {
    items.forEach(i => state.itemSelection.add(i));
    renderItemChecklist(containerEl, rows, itemField, onChange);
    if (onChange) onChange();
  });
  containerEl.querySelector('#btnItemSelectNone').addEventListener('click', () => {
    state.itemSelection.clear();
    renderItemChecklist(containerEl, rows, itemField, onChange);
    if (onChange) onChange();
  });
}

function filterRowsBySelection(rows, itemField) {
  let filtered = rows;
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
  if (rows.length === 0) { containerEl.innerHTML = ''; return; }
  const locField = cat.locationField;
  const itemField = cat.itemField;
  const valueField = cat.fields.find(f => ['檢測數值', '監測數值'].includes(f.key))?.key || '';
  const valueLabel = cat.fields.find(f => f.key === valueField)?.label || '數值';

  const wasOpen = containerEl.querySelector('details')?.open;
  containerEl.innerHTML = `
    <details class="row-detail-toggle" ${wasOpen ? 'open' : ''}>
      <summary>📋 詳細資料列表（可個別排除不匯入的資料，共 ${rows.length} 筆）</summary>
      <div class="mapping-table-wrap" style="max-height:320px;margin-top:8px">
        <table class="mapping-table">
          <thead><tr>
            <th><input type="checkbox" id="rowDetailCheckAll" checked></th>
            <th>地點</th><th>測項</th><th>日期</th><th>時間</th><th>${escapeHtml(valueLabel)}</th>
          </tr></thead>
          <tbody>
            ${rows.map(r => `
              <tr>
                <td><input type="checkbox" class="row-detail-check" data-row-uid="${r._rowUid}" ${state.excludedRowIndices.has(r._rowUid) ? '' : 'checked'}></td>
                <td>${escapeHtml(r[locField] || '')}</td>
                <td>${escapeHtml(r[itemField] || '')}</td>
                <td>${escapeHtml(r['日期(起)'] || '')}</td>
                <td>${escapeHtml(r['時間(起)'] || '')}</td>
                <td>${escapeHtml(r[valueField] || '')}</td>
              </tr>
            `).join('')}
          </tbody>
        </table>
      </div>
    </details>
  `;
  const checkAll = containerEl.querySelector('#rowDetailCheckAll');
  const rowChecks = containerEl.querySelectorAll('.row-detail-check');
  const syncCheckAllState = () => {
    checkAll.checked = [...rowChecks].every(cb => cb.checked);
  };
  rowChecks.forEach(cb => {
    cb.addEventListener('change', () => {
      const uid = Number(cb.dataset.rowUid);
      if (cb.checked) state.excludedRowIndices.delete(uid); else state.excludedRowIndices.add(uid);
      syncCheckAllState();
    });
  });
  checkAll.addEventListener('change', () => {
    rowChecks.forEach(cb => {
      cb.checked = checkAll.checked;
      const uid = Number(cb.dataset.rowUid);
      if (checkAll.checked) state.excludedRowIndices.delete(uid); else state.excludedRowIndices.add(uid);
    });
  });
}

// ---------- batch import (multi-file, auto-detect category) ----------
const AUTO_DETECT_CATEGORIES = ['noise', 'water', 'air']; // categories with smart-parsers capable of self-identifying

function openBatchImportModal() {
  document.getElementById('batchFileInput').value = '';
  document.getElementById('batchDetectStatus').textContent = '';
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
        const rows = SmartParse.parseSheet(catKey, sheetName, grid);
        if (rows && rows.length) {
          rows.forEach(r => { r._sourceFile = file.name; });
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
  state.importParsed = null;
  state.importMode = null;
  state.smartResult = null;
  state.itemSelection = null;
  state.excludedRowIndices = null;
  state.importPeriod = '';
  document.getElementById('importModalTitle').textContent = `匯入${CATEGORIES[catKey].label}監測資料`;
  document.getElementById('importFileInput').value = '';
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

const SMART_PARSE_CATEGORIES = ['noise', 'water', 'air'];

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

  try {
    if (SMART_PARSE_CATEGORIES.includes(catKey)) {
      const aggregate = { rows: [], matchedSheets: [], skippedSheets: [], sourceFiles: new Set() };
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
        for (const [sheetName, grid] of Object.entries(grids)) {
          const rows = SmartParse.parseSheet(catKey, sheetName, grid);
          if (rows && rows.length) {
            rows.forEach(r => { r._sourceFile = file.name; });
            aggregate.rows.push(...rows);
            aggregate.matchedSheets.push(`${file.name} / ${sheetName}`);
            aggregate.sourceFiles.add(file.name);
            if (isPdf) aggregate.hasPdfSource = true;
          } else {
            aggregate.skippedSheets.push(`${file.name} / ${sheetName}`);
          }
        }
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
      alert('這個類別目前無法自動判讀多個檔案的原始報告格式，因此只能一次匯入一個檔案（一般欄位比對）。請改為逐一匯入，或確認檔案是否為噪音／水質／空氣品質可自動判讀的報告格式。');
      return;
    }

    const file = files[0];
    const parsed = await ImportEngine.readFile(file);
    if (!parsed.rows || parsed.rows.length === 0) {
      alert('無法從此檔案讀取到任何資料列，請確認檔案內容或改用 Excel 格式。');
      return;
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
  const project = getCurrentProject();
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
  if (!result._memoryApplied && (cat.methodField || cat.unitField)) {
    const memory = DataStore.getItemMemory(project.id, catKey);
    const methodDiffsByItem = {};
    result.rows.forEach(row => {
      const mem = memory[row[cat.itemField]];
      if (!mem) return;
      if (cat.methodField && !row[cat.methodField] && mem.method) {
        row[cat.methodField] = mem.method; row._methodFromMemory = true;
      } else if (cat.methodField && row[cat.methodField] && mem.method && row[cat.methodField] !== mem.method) {
        methodDiffsByItem[row[cat.itemField]] = { reportMethod: row[cat.methodField], memoryMethod: mem.method };
      }
      if (cat.unitField && !row[cat.unitField] && mem.unitCode) {
        row[cat.unitField] = mem.unitCode; row._unitFromMemory = true;
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
    // rows filtered only by item type (not by row exclusion), otherwise a row the
    // person unchecked would vanish from the list and they could never re-check it.
    const itemFilteredRows = state.itemSelection
      ? result.rows.filter(r => state.itemSelection.has((r[cat.itemField] || '').trim() || '（未標示）'))
      : result.rows;
    renderRowDetailTable(document.getElementById('smartImportRowDetailWrap'), itemFilteredRows, cat, updateCounts);
  };

  renderItemChecklist(document.getElementById('smartImportItemsWrap'), result.rows, cat.itemField, updateCountsAndRowDetail);
  updateCountsAndRowDetail();

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
  // history has items this batch doesn't, offer to add them as blank rows
  // (method/unit pre-filled from item memory, value left for the person to type,
  // since the source is often a PDF this app can't extract numbers from at all).
  //
  // Grouped by CONFIRMED location name, not raw site key: a category like noise
  // can have several distinct raw "sites" (e.g. a 環境噪音 report and a separate
  // 振動 report) that all resolve to the same physical location once the person
  // confirms the official site name — comparing history against just one of those
  // site keys' items would wrongly "miss" every item that actually belongs to the
  // other sub-report at the same location.
  const existingMissingWrap = document.getElementById('smartImportMissingItemsWrap');
  const siteHistory = DataStore.getSiteItemHistory(project.id, catKey);
  const itemMemoryForSuggestion = DataStore.getItemMemory(project.id, catKey);

  const itemsByLoc = {}; // confirmedLoc -> Set(itemName)
  const siteKeysByLoc = {}; // confirmedLoc -> [siteKey, ...] (may span multiple raw sites)
  siteEntries.forEach(([key, site]) => {
    const saved = savedAliases[siteAliasKey(site, result, cat)] || {};
    const confirmedLoc = saved[locField] || site.rawLocation;
    if (!itemsByLoc[confirmedLoc]) { itemsByLoc[confirmedLoc] = new Set(); siteKeysByLoc[confirmedLoc] = []; }
    site.rowIndices.forEach(i => itemsByLoc[confirmedLoc].add(result.rows[i][cat.itemField]));
    siteKeysByLoc[confirmedLoc].push(key);
  });

  const suggestions = [];
  Object.entries(itemsByLoc).forEach(([confirmedLoc, currentItemsSet]) => {
    const historyForLoc = siteHistory[confirmedLoc];
    if (!historyForLoc) return;
    // tolerate the older array-only history format (no per-item category)
    const historicalEntries = Array.isArray(historyForLoc)
      ? historyForLoc.map(item => [item, ''])
      : Object.entries(historyForLoc);
    const missing = historicalEntries.filter(([item]) => !currentItemsSet.has(item));
    if (missing.length > 0) {
      suggestions.push({
        siteKeys: siteKeysByLoc[confirmedLoc], location: confirmedLoc,
        missingItems: missing.map(([item, category]) => ({ item, category })),
      });
    }
  });
  // A whole location can be entirely absent from this import — e.g. this quarter's
  // report simply doesn't cover site A or B at all, even though they were sampled
  // last quarter. Those locations never show up in itemsByLoc above (there's no
  // current row for them at all), so they need a separate pass over the full
  // history to be offered — with no row from this import to copy shared fields
  // (date/coordinates/etc) from, buildSuggestedRows falls back to the remembered
  // site profile and leaves the date blank for the person to fill in themselves.
  Object.entries(siteHistory).forEach(([histLoc, historyForLoc]) => {
    if (itemsByLoc[histLoc]) return; // already handled above — this location DID appear
    const historicalEntries = Array.isArray(historyForLoc)
      ? historyForLoc.map(item => [item, ''])
      : Object.entries(historyForLoc);
    if (historicalEntries.length === 0) return;
    suggestions.push({
      siteKeys: [], location: histLoc, entirelyAbsent: true,
      missingItems: historicalEntries.map(([item, category]) => ({ item, category })),
    });
  });
  state.missingItemSuggestions = suggestions;

  if (suggestions.length === 0) {
    existingMissingWrap.innerHTML = '';
  } else {
    existingMissingWrap.innerHTML = `
      <div class="warning" style="background:#e8f0fe;border-color:#a8c7fa;">
        📋 系統比對過去記錄，以下測站過去曾出現、但本次報告未出現的測項。若要一併新增（檢測方法／單位／項目名稱及對應的檢測類別會依過去記錄先幫您填好，檢測數值請自行輸入），請勾選：
        <div id="missingItemsList" style="margin-top:8px">
          ${suggestions.map((s, i) => `
            <div style="margin-top:6px">
              <label style="font-weight:700">
                <input type="checkbox" class="missing-item-group" data-group-idx="${i}" checked> ${escapeHtml(s.location)}${s.entirelyAbsent ? ' <span class="hint">（本次報告完全沒有這個測站，建議新增的資料日期需要您自行填寫）</span>' : ''}
              </label>
              <div style="margin-left:22px">
                ${s.missingItems.map(({ item, category }) => {
                  const mem = itemMemoryForSuggestion[item];
                  const memParts = [category ? `檢測類別：${category}` : '', mem?.method, mem?.unitCode ? `單位代碼${mem.unitCode}` : ''].filter(Boolean);
                  const memNote = memParts.length ? `（已記憶：${memParts.join('，')}）` : '（無先前記憶的方法/單位，需另外補上）';
                  return `<label style="display:block">
                    <input type="checkbox" class="missing-item-single" data-group-idx="${i}" data-item="${escapeAttr(item)}" checked>
                    ${escapeHtml(item)} <span class="hint">${memNote}</span>
                  </label>`;
                }).join('')}
              </div>
            </div>
          `).join('')}
        </div>
      </div>
    `;
    document.querySelectorAll('.missing-item-group').forEach(cb => {
      cb.addEventListener('change', () => {
        const idx = cb.dataset.groupIdx;
        document.querySelectorAll(`.missing-item-single[data-group-idx="${idx}"]`).forEach(sub => { sub.checked = cb.checked; });
      });
    });
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
        return `<tr data-site-key="${escapeAttr(key)}">
          <td>${escapeHtml(site.siteCode || site.rawLocation || key)}${site.siteCode && site.rawLocation ? `<br><span class="hint">${escapeHtml(site.rawLocation)}</span>` : ''}</td>
          ${profileFields.map(f => {
            const val = saved[f.key] !== undefined ? saved[f.key] : (defaults[f.key] || '');
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
function learnSiteItemHistory(projectId, catKey, cat, rows) {
  DataStore.learnSiteItems(projectId, catKey, cat.locationField, cat.itemField, '檢測類別', rows);
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
      if (oldVal !== newVal) diffFields.push({ key: f.key, label: f.label, oldVal, newVal });
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
          <span><strong>${escapeHtml(c.location)}</strong>／${escapeHtml(c.item)}／${escapeHtml(c.date)}：套用本次匯入的新版本（取消勾選＝保留原有資料，不更動）</span>
        </label>
      </div>
      <table class="conflict-diff-table">
        <thead><tr><th>欄位</th><th>原有資料</th><th>本次匯入</th></tr></thead>
        <tbody>
          ${c.diffFields.map(d => `<tr>
            <td>${escapeHtml(d.label)}</td>
            <td class="diff-old">${escapeHtml(d.oldVal || '（空白）')}</td>
            <td class="diff-new">${escapeHtml(d.newVal || '（空白）')}</td>
          </tr>`).join('')}
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
  const checkedBoxes = [...document.querySelectorAll('.missing-item-single:checked')];
  if (checkedBoxes.length === 0) return [];
  const itemMemory = DataStore.getItemMemory(project.id, catKey);
  const savedAliases = DataStore.getSiteAliases(project.id, catKey);
  const newRows = [];
  checkedBoxes.forEach(cb => {
    const suggestion = state.missingItemSuggestions?.[Number(cb.dataset.groupIdx)];
    if (!suggestion) return;
    const itemName = cb.dataset.item;
    const missingEntry = suggestion.missingItems.find(m => m.item === itemName);
    const historicalCategory = missingEntry ? missingEntry.category : '';

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
      const template = candidateRows.find(r => !historicalCategory || r['檢測類別'] === historicalCategory) || candidateRows[0];
      newRow = { ...template };
    } else {
      // Entirely-absent location (suggestion.entirelyAbsent): this quarter's report
      // has no rows for this site at all, so there's nothing to copy shared fields
      // from — build from scratch using whatever site profile was remembered, and
      // leave date/time blank since there's no genuine reading to attribute a date
      // to; the person fills that in themselves.
      newRow = {};
      cat.fields.forEach(f => { newRow[f.key] = ''; });
      newRow[cat.locationField] = suggestion.location;
      const aliasKeyGuess = historicalCategory ? `${suggestion.location}::${historicalCategory}` : suggestion.location;
      const savedProfile = savedAliases[aliasKeyGuess] || {};
      Object.entries(savedProfile).forEach(([k, v]) => { if (v) newRow[k] = v; });
    }
    newRow[cat.itemField] = itemName;
    if (historicalCategory && '檢測類別' in newRow) newRow['檢測類別'] = historicalCategory;
    ['檢測數值', '監測數值', '比較關係', '檢測極限'].forEach(k => { if (k in newRow) newRow[k] = ''; });
    const mem = itemMemory[itemName];
    if (mem) {
      if (cat.methodField && mem.method) newRow[cat.methodField] = mem.method;
      if (cat.unitField && mem.unitCode) newRow[cat.unitField] = mem.unitCode;
    }
    newRows.push(newRow);
  });
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
  const project = getCurrentProject();
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

  const project = getCurrentProject();
  const catKey = state.importCatKey;
  const cat = CATEGORIES[catKey];
  const mapping = {};
  document.querySelectorAll('#mappingTableBody select').forEach(sel => {
    mapping[sel.dataset.targetField] = sel.value || null;
  });

  const newRows = ImportEngine.applyMapping(state.importParsed.rows, mapping, cat.fields);
  newRows.forEach((r, i) => { r._rowUid = i; });
  if (cat.methodField || cat.unitField) {
    const memory = DataStore.getItemMemory(project.id, catKey);
    newRows.forEach(row => {
      const mem = memory[row[cat.itemField]];
      if (!mem) return;
      if (cat.methodField && !row[cat.methodField] && mem.method) row[cat.methodField] = mem.method;
      if (cat.unitField && !row[cat.unitField] && mem.unitCode) row[cat.unitField] = mem.unitCode;
    });
  }
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

  document.getElementById('importFileInput').addEventListener('change', (e) => handleImportFile(e.target.files));
  document.getElementById('btnImportCancel').addEventListener('click', closeImportModal);
  document.getElementById('btnImportConfirm').addEventListener('click', confirmImport);

  document.getElementById('btnCoordCancel').addEventListener('click', closeCoordModal);
  document.getElementById('btnCoordSave').addEventListener('click', saveCoordModal);

  document.getElementById('btnMethodCancel').addEventListener('click', closeMethodModal);
  document.getElementById('btnMethodSave').addEventListener('click', saveMethodModal);

  document.getElementById('btnUnitCodeRef').addEventListener('click', openUnitRefModal);
  document.getElementById('btnUnitRefClose').addEventListener('click', () => document.getElementById('unitRefModal').classList.add('hidden'));
  document.getElementById('unitRefSearch').addEventListener('input', (e) => renderUnitRefTable(e.target.value));

  document.getElementById('batchFileInput').addEventListener('change', (e) => handleBatchFiles(e.target.files));
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
