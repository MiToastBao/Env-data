// app.js — main UI logic

const state = {
  currentProjectId: null,
  currentTab: 'basic', // 'basic' | 'air' | 'water' | 'geo' | 'noise' | 'eco'
  editingProjectId: null, // set when project modal is in "edit" mode
  importCatKey: null,
  importParsed: null, // { headers, rows, ... } (generic import)
  importMode: null, // 'generic' | 'smart'
  smartResult: null, // { rows, matchedSheets, skippedSheets, sites } from SmartParse.parseWorkbook
  smartOverrides: null, // { siteKey: { fieldKey: value, _remember: bool } }
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
      <div class="p-code">${escapeHtml(p.code)}</div>
      <div class="p-name">${escapeHtml(p.name)}</div>
    </li>
  `).join('');
  list.querySelectorAll('.project-item').forEach(el => {
    el.addEventListener('click', () => selectProject(el.dataset.id));
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
  document.getElementById('btnExportAll').addEventListener('click', () => {
    const anyData = CATEGORY_ORDER.some(c => DataStore.getData(project.id, c).length > 0);
    if (!anyData) { alert('目前尚無任何監測資料可匯出，請先匯入或新增資料。'); return; }
    ExportEngine.downloadAll(project);
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
  const body = document.getElementById('tabBody');

  body.innerHTML = `
    <div class="toolbar">
      <div class="toolbar-left">
        <button class="btn btn-primary btn-sm" id="btnImport">📥 匯入資料（Excel/PDF）</button>
        <button class="btn btn-ghost btn-sm" id="btnAddRow">＋ 新增一筆</button>
        <button class="btn btn-ghost btn-sm" id="btnExportCat">匯出此類別（${cat.sourceFile}）</button>
      </div>
      <div class="row-count">共 ${rows.length} 筆資料</div>
    </div>
    <div class="table-wrap">
      <table class="data-grid">
        <thead><tr>
          <th>#</th>
          ${cat.fields.map(f => `<th>${escapeHtml(f.label)}${f.required ? '<span class="req">＊</span>' : ''}</th>`).join('')}
          <th>操作</th>
        </tr></thead>
        <tbody id="gridBody">${rows.map((r, idx) => rowHtml(cat, r, idx)).join('')}</tbody>
      </table>
    </div>
    ${rows.length === 0 ? '<p class="hint" style="margin-top:10px">尚無資料。可點「匯入資料」上傳該類別的檢測結果檔案，或「新增一筆」手動輸入。</p>' : ''}
  `;

  document.getElementById('btnImport').addEventListener('click', () => openImportModal(catKey));
  document.getElementById('btnAddRow').addEventListener('click', () => addEmptyRow(project, catKey));
  document.getElementById('btnExportCat').addEventListener('click', () => {
    if (rows.length === 0) { alert('此類別尚無資料可匯出。'); return; }
    ExportEngine.downloadCategory(project, DataStore.getBasicInfo(project.id), catKey);
  });

  wireGridEvents(project, catKey, cat);
}

function rowHtml(cat, row, idx) {
  const cells = cat.fields.map(f => `<td>${fieldControlHTML(f, row[f.key], `data-row="${idx}"`)}</td>`).join('');
  return `<tr data-row="${idx}"><td>${idx + 1}</td>${cells}<td class="col-actions"><button class="row-del-btn" data-row="${idx}" title="刪除此列">🗑</button></td></tr>`;
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
      return `<select ${base}>${opts}</select>`;
    }
    case 'date':
      return `<input type="date" ${base} value="${escapeAttr(toDateInputValue(value))}">`;
    case 'time':
      return `<input type="time" step="1" ${base} value="${escapeAttr(toTimeInputValue(value))}">`;
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

  const commit = (rowIdx, fieldKey, value) => {
    const rows = DataStore.getData(project.id, catKey);
    if (!rows[rowIdx]) return;
    rows[rowIdx][fieldKey] = value;
    DataStore.saveData(project.id, catKey, rows);
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
    if (t.dataset.field && t.tagName === 'SELECT') {
      commit(Number(t.dataset.row), t.dataset.field, t.value);
    }
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

// ---------- import modal ----------
function openImportModal(catKey) {
  state.importCatKey = catKey;
  state.importParsed = null;
  state.importMode = null;
  state.smartResult = null;
  state.smartOverrides = null;
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
}

const SMART_PARSE_CATEGORIES = ['noise', 'water', 'air'];

async function handleImportFile(file) {
  if (!file) return;
  const catKey = state.importCatKey;
  const isSpreadsheet = /\.(xlsx|xls)$/i.test(file.name);

  try {
    // Try the smart form-parser first for categories/report types it understands.
    if (isSpreadsheet && SMART_PARSE_CATEGORIES.includes(catKey)) {
      const grids = await ImportEngine.readWorkbookGrids(file);
      const result = SmartParse.parseWorkbook(catKey, grids);
      if (result.rows.length > 0) {
        state.importMode = 'smart';
        state.smartResult = result;
        state.smartOverrides = {};
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
}

// ---------- smart import (report-form parser) preview ----------
function renderSmartImportPreview() {
  const project = getCurrentProject();
  const catKey = state.importCatKey;
  const result = state.smartResult;
  const profileFields = SmartParse.SITE_PROFILE_FIELDS[catKey] || [];
  const savedAliases = DataStore.getSiteAliases(project.id, catKey);

  const siteEntries = Object.entries(result.sites); // [key, {siteCode, rawLocation, rowIndices}]

  document.getElementById('smartImportSummary').textContent =
    `系統辨識出 ${siteEntries.length} 個測站、共 ${result.rows.length} 筆資料（來自 ${result.matchedSheets.length} 個工作表）。`
    + ` 部分欄位無法從報告本身取得，請在下方為每個測站補充一次，之後匯入同一測站會自動套用。`;

  const locField = catKey === 'noise' ? '監測地點' : '採樣地點';

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
          <td>${site.rowIndices.length}</td>
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
  confirmBtn.textContent = `確認匯入 ${result.rows.length} 筆資料`;
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

  // strip internal metadata fields before saving into the schema data
  const cleanRows = result.rows.map(r => {
    const out = {};
    cat.fields.forEach(f => { out[f.key] = r[f.key] || ''; });
    return out;
  });

  const existing = DataStore.getData(project.id, catKey);
  DataStore.saveData(project.id, catKey, existing.concat(cleanRows));

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
  const existing = DataStore.getData(project.id, catKey);
  const merged = existing.concat(newRows);
  DataStore.saveData(project.id, catKey, merged);

  closeImportModal();
  renderContent();
  alert(`已匯入 ${newRows.length} 筆資料到「${cat.label}」。請於表格中核對內容是否正確。`);
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
  renderProjectList();
  renderContent();

  document.getElementById('btnNewProject').addEventListener('click', () => openProjectModal(null));
  document.getElementById('btnProjectCancel').addEventListener('click', closeProjectModal);
  document.getElementById('btnProjectSave').addEventListener('click', saveProjectModal);

  document.getElementById('btnBackupExport').addEventListener('click', backupExport);
  document.getElementById('btnBackupImport').addEventListener('click', () => document.getElementById('backupFileInput').click());
  document.getElementById('backupFileInput').addEventListener('change', (e) => backupImport(e.target.files[0]));

  document.getElementById('importFileInput').addEventListener('change', (e) => handleImportFile(e.target.files[0]));
  document.getElementById('btnImportCancel').addEventListener('click', closeImportModal);
  document.getElementById('btnImportConfirm').addEventListener('click', confirmImport);
}

document.addEventListener('DOMContentLoaded', init);
