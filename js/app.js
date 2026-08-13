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
};

// ---------- helpers ----------
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
function toTimeInputValue(v) {
  if (!v) return '';
  const s = String(v);
  const m = s.match(/(\d{1,2}):(\d{2})(:(\d{2}))?/);
  if (m) return `${m[1].padStart(2, '0')}:${m[2]}:${m[4] || '00'}`;
  return '';
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
      control = `<input type="date" data-field="${f.key}" value="${escapeAttr(toDateInputValue(val))}">`;
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
  });
}

// ---------- category data tab ----------
function renderCategoryTab(project, catKey) {
  const cat = CATEGORIES[catKey];
  const rows = DataStore.getData(project.id, catKey);
  const batches = DataStore.getImportBatches(project.id, catKey);
  const body = document.getElementById('tabBody');

  body.innerHTML = `
    <div class="toolbar">
      <div class="toolbar-left">
        <button class="btn btn-primary btn-sm" id="btnImport">📥 匯入資料（Excel/PDF）</button>
        <button class="btn btn-ghost btn-sm" id="btnAddRow">＋ 新增一筆</button>
        <button class="btn btn-ghost btn-sm" id="btnCoordManager">📍 測站座標管理</button>
        <button class="btn btn-ghost btn-sm" id="btnBatchHistory">📜 匯入紀錄${batches.length ? ` (${batches.length})` : ''}</button>
        <button class="btn btn-ghost btn-sm" id="btnExportCat">匯出此類別（${cat.sourceFile}）</button>
        <button class="btn btn-danger btn-sm" id="btnClearCat" ${rows.length === 0 ? 'disabled' : ''}>🗑 清空此類別</button>
      </div>
      <div class="row-count">共 ${rows.length} 筆資料</div>
    </div>
    <div class="toolbar bulk-toolbar hidden" id="bulkToolbar">
      <span id="bulkSelCount">已選取 0 筆</span>
      <button class="btn btn-danger btn-sm" id="btnBulkDelete">🗑 刪除已選取</button>
      <button class="btn btn-ghost btn-sm" id="btnBulkClear">取消選取</button>
    </div>
    <div class="table-wrap">
      <table class="data-grid">
        <thead><tr>
          <th class="col-check"><input type="checkbox" id="checkAllRows" ${rows.length === 0 ? 'disabled' : ''}></th>
          <th>操作</th>
          <th>#</th>
          ${cat.fields.map(f => `<th${f.help ? ` title="${escapeAttr(f.help)}"` : ''}>${escapeHtml(f.label)}${f.required ? '<span class="req">＊</span>' : ''}${f.help ? ' ℹ️' : ''}</th>`).join('')}
        </tr></thead>
        <tbody id="gridBody">${rows.map((r, idx) => rowHtml(cat, r, idx)).join('')}</tbody>
      </table>
    </div>
    ${rows.length === 0 ? '<p class="hint" style="margin-top:10px">尚無資料。可點「匯入資料」上傳該類別的檢測結果檔案，或「新增一筆」手動輸入。</p>' : ''}
  `;

  document.getElementById('btnImport').addEventListener('click', () => openImportModal(catKey));
  document.getElementById('btnAddRow').addEventListener('click', () => addEmptyRow(project, catKey));
  document.getElementById('btnCoordManager').addEventListener('click', () => openCoordModal(project, catKey));
  document.getElementById('btnBatchHistory').addEventListener('click', () => openBatchHistoryModal(project, catKey));
  document.getElementById('btnExportCat').addEventListener('click', () => {
    if (rows.length === 0) { alert('此類別尚無資料可匯出。'); return; }
    ExportEngine.downloadCategory(project, DataStore.getBasicInfo(project.id), catKey);
  });
  document.getElementById('btnClearCat').addEventListener('click', () => {
    if (rows.length === 0) return;
    if (!confirm(`確定要清空「${cat.label}」的全部 ${rows.length} 筆資料嗎？此操作無法復原。`)) return;
    DataStore.clearData(project.id, catKey);
    renderContent();
  });

  wireGridEvents(project, catKey, cat);
  wireBulkSelection(project, catKey, cat);
}

function wireBulkSelection(project, catKey, cat) {
  const tbody = document.getElementById('gridBody');
  const checkAll = document.getElementById('checkAllRows');
  const bulkToolbar = document.getElementById('bulkToolbar');
  const bulkCount = document.getElementById('bulkSelCount');

  const getChecked = () => [...tbody.querySelectorAll('.row-check:checked')].map(cb => Number(cb.dataset.row));
  const updateBulkUI = () => {
    const n = getChecked().length;
    bulkToolbar.classList.toggle('hidden', n === 0);
    bulkCount.textContent = `已選取 ${n} 筆`;
    if (checkAll) checkAll.checked = n > 0 && n === tbody.querySelectorAll('.row-check').length;
  };

  tbody.addEventListener('change', (e) => {
    if (e.target.classList.contains('row-check')) updateBulkUI();
  });
  if (checkAll) {
    checkAll.addEventListener('change', () => {
      tbody.querySelectorAll('.row-check').forEach(cb => { cb.checked = checkAll.checked; });
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
      <thead><tr><th>匯入時間</th><th>來源檔案</th><th>方式</th><th>筆數</th><th>操作</th></tr></thead>
      <tbody>
        ${batches.slice().reverse().map(b => `
          <tr data-batch-id="${escapeAttr(b.id)}">
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
  const cells = cat.fields.map(f => `<td>${fieldControlHTML(f, row[f.key], `data-row="${idx}"`)}</td>`).join('');
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
      return `<input type="date" ${base} value="${escapeAttr(toDateInputValue(value))}">`;
    case 'time':
      // Plain text rather than a native <input type=time>: native time pickers on many
      // devices show a scroll-wheel that's fiddly to land on an exact second, and on
      // some mobile browsers don't reliably fire change events at all. Typing "1430",
      // "14:30", or "14:30:00" all work — normalized to HH:MM:SS on blur.
      return `<input type="text" ${base} value="${escapeAttr(toTimeInputValue(value))}" class="time-input" placeholder="HH:MM:SS" inputmode="numeric" maxlength="8">`;
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

  // If the person fills in coordinates for one row, carry them forward to any other
  // row sharing the same 日期(起) that doesn't have that field filled in yet — so a
  // multi-item report (e.g. water quality's several test items from one sampling
  // event) only needs coordinates entered once. Never overwrites an existing value.
  const propagateCoordsForDate = (rowIdx, fieldKey) => {
    if (!COORD_FIELDS.includes(fieldKey)) return false;
    const rows = DataStore.getData(project.id, catKey);
    const source = rows[rowIdx];
    if (!source || !source[fieldKey] || !source['日期(起)']) return false;
    let changed = false;
    rows.forEach((r, idx) => {
      if (idx === rowIdx || r['日期(起)'] !== source['日期(起)']) return;
      if (!r[fieldKey]) { r[fieldKey] = source[fieldKey]; changed = true; }
    });
    if (changed) DataStore.saveData(project.id, catKey, rows);
    return changed;
  };

  // If the person corrects a date/time on one row, offer to sync all four date/time
  // fields to every other row that (a) came from the same import batch and (b) shares
  // the same sampling location — e.g. correcting one test item's date in a multi-item
  // water report should usually update the whole site's rows from that report, since
  // they're really one sampling event. This always asks first rather than silently
  // overwriting, both to avoid surprising edits and because some browsers/devices
  // don't reliably fire events for native time pickers, which made a silent version
  // of this hard to trust.
  const DATE_TIME_FIELDS = ['日期(起)', '時間(起)', '日期(迄)', '時間(迄)'];
  const offerDateTimeSync = (rowIdx) => {
    const rows = DataStore.getData(project.id, catKey);
    const source = rows[rowIdx];
    const locField = cat.locationField;
    if (!source || !source._batchId || !source[locField]) return false;
    const matches = rows
      .map((r, idx) => ({ r, idx }))
      .filter(({ r, idx }) => idx !== rowIdx && r._batchId === source._batchId && r[locField] === source[locField]);
    if (matches.length === 0) return false;
    const anyDiff = matches.some(({ r }) => DATE_TIME_FIELDS.some(f => r[f] !== source[f]));
    if (!anyDiff) return false;
    const ok = confirm(
      `偵測到同一份檔案、同一個測站「${source[locField]}」還有 ${matches.length} 筆其他資料。\n` +
      `是否要將這些資料的採樣日期／時間一併同步更新為與這一筆相同？\n\n` +
      `（選擇「取消」則只修改目前這一筆，其他資料維持原狀。）`
    );
    if (!ok) return false;
    matches.forEach(({ r }) => { DATE_TIME_FIELDS.forEach(f => { r[f] = source[f]; }); });
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
      if (propagateCoordsForDate(Number(t.dataset.row), t.dataset.field)) { renderContent(); return; }
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
      const normalized = normalizeTimeString(t.value);
      t.value = normalized;
      commit(rowIdx, fieldKey, normalized);
    }

    if (propagateCoordsForDate(rowIdx, fieldKey)) { renderContent(); return; }
    if (DATE_TIME_FIELDS.includes(fieldKey) && offerDateTimeSync(rowIdx)) renderContent();
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

  // group row indices by site name text
  const groups = {}; // locationText -> { indices: [...], coordSystem, x, y }
  rows.forEach((row, idx) => {
    const loc = (row[locField] || '').trim() || '（未命名測站）';
    if (!groups[loc]) {
      groups[loc] = {
        indices: [],
        coordSystem: row['座標系統'] || '',
        x: row['採樣座標-經度 X'] || '',
        y: row['採樣座標-緯度 Y'] || '',
      };
    }
    groups[loc].indices.push(idx);
    // prefer an already-filled value if this group doesn't have one yet
    if (!groups[loc].x && row['採樣座標-經度 X']) groups[loc].x = row['採樣座標-經度 X'];
    if (!groups[loc].y && row['採樣座標-緯度 Y']) groups[loc].y = row['採樣座標-緯度 Y'];
    if (!groups[loc].coordSystem && row['座標系統']) groups[loc].coordSystem = row['座標系統'];
  });

  const entries = Object.entries(groups);
  const wrap = document.getElementById('coordSitesWrap');
  if (entries.length === 0) {
    wrap.innerHTML = '<p class="hint" style="padding:14px">目前此類別尚無資料，請先新增或匯入資料後再使用測站座標管理。</p>';
  } else {
    wrap.innerHTML = `<table class="mapping-table">
      <thead><tr><th>測站名稱</th><th title="2：WGS84（全球座標，例如經度 120.681，緯度 24.147）／3：TWD97-TM2（投影座標系，例如 X=193150, Y=2670900）">座標系統 ℹ️</th><th>經度 X</th><th>緯度 Y</th><th>筆數</th></tr></thead>
      <tbody id="coordSitesBody">
        ${entries.map(([loc, g]) => `<tr data-loc="${escapeAttr(loc)}">
          <td>${escapeHtml(loc)}</td>
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
    const loc = tr.dataset.loc;
    const values = {};
    tr.querySelectorAll('[data-coord-field]').forEach(el => { values[el.dataset.coordField] = el.value; });
    rows.forEach(row => {
      const rowLoc = (row[locField] || '').trim() || '（未命名測站）';
      if (rowLoc === loc) Object.assign(row, values);
    });
  });

  DataStore.saveData(projectId, catKey, rows);
  closeCoordModal();
  renderContent();
  alert('已套用座標到符合的資料列。');
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
  const basicInfo = DataStore.getBasicInfo(project.id);
  checked.forEach(catKey => ExportEngine.downloadCategory(project, basicInfo, catKey));
  closeExportSelectModal();
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
  if (!state.itemSelection) return rows;
  return rows.filter(r => state.itemSelection.has((r[itemField] || '').trim() || '（未標示）'));
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
    let grids;
    try {
      grids = await ImportEngine.readWorkbookGrids(file);
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
          perCategory[catKey].rows.push(...rows);
          perCategory[catKey].matchedSheets.push(`${file.name} / ${sheetName}`);
          perCategory[catKey].sourceFiles.add(file.name);
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

async function handleImportFile(file) {
  if (!file) return;
  const catKey = state.importCatKey;
  const isSpreadsheet = /\.(xlsx|xls)$/i.test(file.name);
  state.currentImportSourceLabel = file.name;

  try {
    // Try the smart form-parser first for categories/report types it understands.
    if (isSpreadsheet && SMART_PARSE_CATEGORIES.includes(catKey)) {
      const grids = await ImportEngine.readWorkbookGrids(file);
      const result = SmartParse.parseWorkbook(catKey, grids);
      if (result.rows.length > 0) {
        state.importMode = 'smart';
        state.smartResult = result;
        renderSmartImportPreview();
        return;
      }
      // fall through to generic import if nothing recognized
    }

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
    if (!mappedHeader) { wrap.innerHTML = ''; return; }
    const shimRows = rows.map(r => ({ [cat.itemField]: r[mappedHeader] }));
    renderItemChecklist(wrap, shimRows, cat.itemField, null);
  };
  refreshGenericItemChecklist();
  const itemSel = tbody.querySelector(`select[data-target-field="${cat.itemField}"]`);
  if (itemSel) itemSel.addEventListener('change', refreshGenericItemChecklist);
}

// ---------- smart import (report-form parser) preview ----------
function renderSmartImportPreview() {
  const project = getCurrentProject();
  const catKey = state.importCatKey;
  const cat = CATEGORIES[catKey];
  const result = state.smartResult;
  const profileFields = SmartParse.SITE_PROFILE_FIELDS[catKey] || [];
  const savedAliases = DataStore.getSiteAliases(project.id, catKey);

  const siteEntries = Object.entries(result.sites); // [key, {siteCode, rawLocation, rowIndices}]
  state.itemSelection = null; // reset so renderItemChecklist re-seeds with "all checked"

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

  renderItemChecklist(document.getElementById('smartImportItemsWrap'), result.rows, cat.itemField, updateCounts);

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
        const saved = savedAliases[key] || {};
        const firstRow = result.rows[site.rowIndices[0]] || {};
        const defaults = { [locField]: site.rawLocation, ...firstRow };
        return `<tr data-site-key="${escapeAttr(key)}">
          <td>${escapeHtml(site.siteCode || site.rawLocation || key)}${site.siteCode && site.rawLocation ? `<br><span class="hint">${escapeHtml(site.rawLocation)}</span>` : ''}</td>
          ${profileFields.map(f => {
            const val = saved[f.key] !== undefined ? saved[f.key] : (defaults[f.key] || '');
            if (f.type === 'select') {
              return `<td><select data-site-field="${f.key}">${f.options.map(o => `<option value="${escapeAttr(o)}" ${o === val ? 'selected' : ''}>${o || '（未選擇）'}</option>`).join('')}</select></td>`;
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
    const overrides = {};
    tr.querySelectorAll('[data-site-field]').forEach(el => {
      overrides[el.dataset.siteField] = el.value;
    });
    const remember = tr.querySelector('[data-site-remember]').checked;
    if (remember) savedAliases[siteKey] = overrides;
    else delete savedAliases[siteKey];

    // apply overrides to every row belonging to this site
    const site = result.sites[siteKey];
    site.rowIndices.forEach(idx => {
      Object.assign(result.rows[idx], overrides);
    });
  });
  DataStore.saveSiteAliases(project.id, catKey, savedAliases);

  const selectedRows = filterRowsBySelection(result.rows, cat.itemField);
  if (selectedRows.length === 0) { alert('目前沒有勾選任何監測項目，請至少勾選一項再匯入。'); return; }

  // tag rows with a shared batch id so the grid can later offer "same file, same
  // site" sync when the person corrects a date/time on one row (see wireGridEvents)
  const batchId = 'b_' + Date.now() + '_' + Math.random().toString(36).slice(2, 8);

  // strip internal metadata fields before saving into the schema data (keep _batchId)
  const cleanRows = selectedRows.map(r => {
    const out = { _batchId: batchId };
    cat.fields.forEach(f => { out[f.key] = r[f.key] || ''; });
    return out;
  });

  const existing = DataStore.getData(project.id, catKey);
  DataStore.saveData(project.id, catKey, existing.concat(cleanRows));
  DataStore.addImportBatch(project.id, catKey, {
    id: batchId,
    timestamp: new Date().toISOString(),
    sourceLabel: state.currentImportSourceLabel || '（未知來源）',
    mode: 'smart',
    rowCount: cleanRows.length,
  });

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
  alert(`已匯入 ${cleanRows.length} 筆資料到「${cat.label}」。請於表格中核對內容是否正確，特別是尚未有測站設定的欄位。`);
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
  const selectedRows = filterRowsBySelection(newRows, cat.itemField);
  if (selectedRows.length === 0) { alert('目前沒有勾選任何監測項目，請至少勾選一項再匯入。'); return; }
  const batchId = 'b_' + Date.now() + '_' + Math.random().toString(36).slice(2, 8);
  selectedRows.forEach(r => { r._batchId = batchId; });
  const existing = DataStore.getData(project.id, catKey);
  const merged = existing.concat(selectedRows);
  DataStore.saveData(project.id, catKey, merged);
  DataStore.addImportBatch(project.id, catKey, {
    id: batchId,
    timestamp: new Date().toISOString(),
    sourceLabel: state.currentImportSourceLabel || '（未知來源）',
    mode: 'generic',
    rowCount: selectedRows.length,
  });

  closeImportModal();
  renderContent();
  alert(`已匯入 ${selectedRows.length} 筆資料到「${cat.label}」。請於表格中核對內容是否正確。`);
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

  document.getElementById('importFileInput').addEventListener('change', (e) => handleImportFile(e.target.files[0]));
  document.getElementById('btnImportCancel').addEventListener('click', closeImportModal);
  document.getElementById('btnImportConfirm').addEventListener('click', confirmImport);

  document.getElementById('btnCoordCancel').addEventListener('click', closeCoordModal);
  document.getElementById('btnCoordSave').addEventListener('click', saveCoordModal);

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
  document.querySelectorAll('.modal-overlay').forEach(overlay => {
    overlay.addEventListener('click', (e) => {
      if (e.target === overlay) overlay.classList.add('hidden');
    });
  });
  document.addEventListener('keydown', (e) => {
    if (e.key !== 'Escape') return;
    document.querySelectorAll('.modal-overlay').forEach(overlay => {
      if (!overlay.classList.contains('hidden')) overlay.classList.add('hidden');
    });
  });
}

document.addEventListener('DOMContentLoaded', init);
