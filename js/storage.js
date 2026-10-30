// storage.js
// All data lives in the browser's localStorage. This is a static site
// (designed for GitHub Pages) with no backend, so data is per-browser.
// Use "匯出備份" / "匯入備份" in the app to move data between machines.

const DataStore = {
  PROJECTS_KEY: 'envapp_projects_v1',

  /*
   * 每一次寫入都要走這裡。
   *
   * localStorage 滿了會丟 QuotaExceededError，而 setItem 是原子的——舊值原封不動
   * 留著。以前只有 saveData 接住並通知，其他七個 save* 都是裸的 setItem：例外
   * 一路丟進 focusout／click 處理常式被瀏覽器默默吃掉，畫面顯示新值、存下去的
   * 還是舊值，使用者以為改好了。專案名稱、基本資料、地點別名、匯入紀錄、
   * 檢測方法記憶全都在這條路上。
   *
   * 先通知（由 app.js 掛上 onStorageError 顯示訊息），再照樣往上丟，
   * 原本已經有 try/catch 的呼叫端行為完全不變。
   */
  onStorageError: null,
  _setItem(key, value) {
    try {
      localStorage.setItem(key, value);
    } catch (err) {
      if (typeof this.onStorageError === 'function') this.onStorageError(err);
      throw err;
    }
  },

  getProjects() {
    // Guard the SHAPE, not just the parse. A malformed value here used to throw out
    // of renderProjectList during init(), which aborted init before any event
    // listener was attached — leaving a page that looks normal but where every
    // button, including 匯出備份, is dead and the data can't even be rescued.
    try {
      const v = JSON.parse(localStorage.getItem(this.PROJECTS_KEY));
      return Array.isArray(v) ? v.filter(p => p && typeof p === 'object' && p.id) : [];
    } catch (e) {
      return [];
    }
  },

  saveProjects(projects) {
    this._setItem(this.PROJECTS_KEY, JSON.stringify(projects));
  },

  createProject(code, name) {
    const projects = this.getProjects();
    const id = 'p_' + Date.now() + '_' + Math.random().toString(36).slice(2, 8);
    const project = { id, code, name, createdAt: new Date().toISOString() };
    projects.push(project);
    this.saveProjects(projects);
    return project;
  },

  updateProject(id, patch) {
    const projects = this.getProjects();
    const idx = projects.findIndex(p => p.id === id);
    if (idx === -1) return null;
    projects[idx] = { ...projects[idx], ...patch };
    this.saveProjects(projects);
    return projects[idx];
  },

  deleteProject(id) {
    const projects = this.getProjects().filter(p => p.id !== id);
    this.saveProjects(projects);
    this._removeProjectKeys(id); // one shared key list — see the helper near importAll
  },

  _basicKey(projectId) {
    return `envapp_basic_${projectId}`;
  },
  _dataKey(projectId, category) {
    return `envapp_data_${projectId}_${category}`;
  },
  _siteAliasKey(projectId, category) {
    return `envapp_sitealias_${projectId}_${category}`;
  },
  _batchKey(projectId, category) {
    return `envapp_batches_${projectId}_${category}`;
  },
  _methodMemoryKey(projectId, category) {
    return `envapp_methodmem_${projectId}_${category}`;
  },
  _siteItemHistoryKey(projectId, category) {
    return `envapp_siteitems_${projectId}_${category}`;
  },

  // v4.46（#29～#32）：「跨季測站記憶」「測項方法記憶」不再另外存。
  // 以前這兩本帳只增不減，資料刪掉了帳還在，下次匯入會拿已經刪掉的資料來比對。
  // 現在都在 app.js 從目前表格裡的資料現算（liveSiteHistory／liveItemMemory）。
  // 舊的鍵（envapp_siteitems_*、envapp_methodmem_*）由 migrateV446() 清掉。

  // v4.46（#33、#34）：每個計畫各自一份的設定
  _presetKey(projectId, category) {
    return `envapp_itemPresets_${projectId}_${category}`;
  },
  _decimalKey(projectId) {
    return `envapp_decimal_settings_v1_${projectId}`;
  },

  /**
   * v4.46 一次性轉換（冪等，每次開啟都跑也安全）：
   *  1. 刪掉所有計畫的舊帳本（跨季測站記憶、測項方法記憶）。
   *  2. 以前整台電腦共用的「常用測項記憶標籤」「小數位數設定」複製給每一個現有計畫
   *     （以前每個計畫都看得到同一份，複製後各計畫看到的和以前一樣，之後各改各的），
   *     然後刪掉共用的那一份。
   */
  /**
   * v4.48（#53，使用者裁示）一次性清理（冪等）：v4.46 以前存的「記住此設定」測站對照沒有記是哪一季存的，
   * 無法判斷是不是來自已經刪掉的匯入，一律清掉；下次匯入時在測站表重新勾一次即可。
   * v4.47 起存的對照帶 _period，不受影響。
   */
  migrateV448() {
    this.getProjects().forEach(p => {
      CATEGORY_ORDER.forEach(c => {
        const all = this.getSiteAliases(p.id, c);
        const keep = {};
        let dropped = 0;
        Object.entries(all || {}).forEach(([k, v]) => {
          if (v && typeof v === 'object' && Object.prototype.hasOwnProperty.call(v, '_period')) keep[k] = v; else dropped++;
        });
        if (dropped) this.saveSiteAliases(p.id, c, keep);
      });
    });
  },

  migrateV446() {
    const projects = this.getProjects();
    const legacyDecimal = localStorage.getItem(DECIMAL_SETTINGS_KEY);
    const legacyPresets = {};
    CATEGORY_ORDER.forEach(c => { legacyPresets[c] = localStorage.getItem(`envapp_itemPresets_${c}`); });
    projects.forEach(p => {
      if (legacyDecimal !== null && localStorage.getItem(this._decimalKey(p.id)) === null) this._setItem(this._decimalKey(p.id), legacyDecimal);
      CATEGORY_ORDER.forEach(c => {
        if (legacyPresets[c] !== null && localStorage.getItem(this._presetKey(p.id, c)) === null) this._setItem(this._presetKey(p.id, c), legacyPresets[c]);
      });
    });
    localStorage.removeItem(DECIMAL_SETTINGS_KEY);
    CATEGORY_ORDER.forEach(c => localStorage.removeItem(`envapp_itemPresets_${c}`));
    const stale = [];
    for (let i = 0; i < localStorage.length; i++) {
      const k = localStorage.key(i);
      if (k && (k.startsWith('envapp_siteitems_') || k.startsWith('envapp_methodmem_'))) stale.push(k);
    }
    stale.forEach(k => localStorage.removeItem(k));
  },

  // Import-batch registry: one entry per "import" action (a single-file import,
  // or one category's slice of a multi-file batch import), so the person can see
  // what was imported and remove a whole import in one step instead of hunting
  // for individual rows. Manually-added/edited rows aren't tracked here.
  getImportBatches(projectId, category) {
    try {
      return JSON.parse(localStorage.getItem(this._batchKey(projectId, category))) || [];
    } catch (e) {
      return [];
    }
  },
  saveImportBatches(projectId, category, batches) {
    this._setItem(this._batchKey(projectId, category), JSON.stringify(batches));
  },
  addImportBatch(projectId, category, batchMeta) {
    const batches = this.getImportBatches(projectId, category);
    batches.push(batchMeta);
    this.saveImportBatches(projectId, category, batches);
  },
  /** Remove a batch's rows from the data and remove it from the registry. */
  deleteImportBatch(projectId, category, batchId) {
    const rows = this.getData(projectId, category).filter(r => r._batchId !== batchId);
    this.saveData(projectId, category, rows);
    const batches = this.getImportBatches(projectId, category).filter(b => b.id !== batchId);
    this.saveImportBatches(projectId, category, batches);
  },

  getSiteAliases(projectId, category) {
    try {
      return JSON.parse(localStorage.getItem(this._siteAliasKey(projectId, category))) || {};
    } catch (e) {
      return {};
    }
  },
  saveSiteAliases(projectId, category, aliases) {
    this._setItem(this._siteAliasKey(projectId, category), JSON.stringify(aliases));
  },

  getBasicInfo(projectId) {
    try {
      return JSON.parse(localStorage.getItem(this._basicKey(projectId))) || {};
    } catch (e) {
      return {};
    }
  },
  saveBasicInfo(projectId, info) {
    this._setItem(this._basicKey(projectId), JSON.stringify(info));
  },

  getData(projectId, category) {
    try {
      return JSON.parse(localStorage.getItem(this._dataKey(projectId, category))) || [];
    } catch (e) {
      return [];
    }
  },
  /** ⚠️ 存不進去一定要讓人知道——見檔案最上面 _setItem 的說明。 */
  saveData(projectId, category, rows) {
    this._setItem(this._dataKey(projectId, category), JSON.stringify(rows));
  },
  clearData(projectId, category) {
    this.saveData(projectId, category, []);
    this.saveImportBatches(projectId, category, []);
  },

  /*
   * 不屬於任何一個專案、而是整個瀏覽器共用的設定。
   *
   * 這兩樣以前**沒有進備份檔**：換一台電腦、還原備份之後，小數位數設定會退回
   * 出廠預設（只有噪音監測數值補兩位），檢測項目的「最近使用的組合」整個消失。
   * 專案、資料、別名、匯入紀錄都搬過去了，只有這兩個沒有，而且畫面上不會報錯——
   * 使用者要等到匯出的檔案位數不對才發現。
   *
   * 鍵名要和 schema.js 的 DECIMAL_SETTINGS_KEY、app.js 的 presetStorageKey 一致。
   */
  _browserSettingKeys() {
    // v4.46 起這兩樣改成每個計畫各一份（#33、#34），備份檔裡放在各計畫的 settings。
    // 這份清單只用來讀**舊版備份檔**裡整台電腦共用的那一份。
    return [DECIMAL_SETTINGS_KEY, ...CATEGORY_ORDER.map(c => `envapp_itemPresets_${c}`)];
  },
  _projectSettings(projectId) {
    const out = { decimal: localStorage.getItem(this._decimalKey(projectId)), presets: {} };
    CATEGORY_ORDER.forEach(c => { out.presets[c] = localStorage.getItem(this._presetKey(projectId, c)); });
    return out;
  },

  // Full export/import for backup/transfer between browsers
  exportAll() {
    const projects = this.getProjects();
    const out = { projects: [], browserSettings: {} };
    projects.forEach(p => {
      // v4.46：不再有跨季測站記憶、測項方法記憶（都從資料現算）；設定改成每個計畫一份
      const entry = { ...p, basicInfo: this.getBasicInfo(p.id), data: {}, siteAliases: {}, importBatches: {}, settings: this._projectSettings(p.id) };
      CATEGORY_ORDER.forEach(cat => {
        entry.data[cat] = this.getData(p.id, cat);
        entry.siteAliases[cat] = this.getSiteAliases(p.id, cat);
        // Without this the 匯入紀錄 list came back empty after a backup restore, so
        // "🗑 刪除此批次" could no longer undo an import on the new machine.
        entry.importBatches[cat] = this.getImportBatches(p.id, cat);
      });
      out.projects.push(entry);
    });
    return out;
  },

  /**
   * Restore a backup.
   *
   * ORDER MATTERS. The old version deleted every existing project FIRST and then
   * wrote the new ones — so a single failing setItem part-way through (quota is
   * realistic here, since a restore briefly holds both copies) left the person with
   * no projects at all and nothing restored, behind an error message that only
   * talked about the file format. Now: write all the new data first, publish the new
   * project list last, and only then remove the old projects' keys. If anything
   * throws before the list is published, the previous state is still intact.
   */
  importAll(payload, mode = 'merge') {
    if (!payload || !Array.isArray(payload.projects)) throw new Error('備份檔格式不正確');
    const previousProjects = this.getProjects();
    const written = [];
    let newProjects;
    try {
      newProjects = mode === 'replace' ? [] : previousProjects.slice();
      payload.projects.forEach((entry, i) => {
        const newId = 'p_' + Date.now() + '_' + i + '_' + Math.random().toString(36).slice(2, 8);
        written.push(newId);
        this.saveBasicInfo(newId, entry.basicInfo || {});
        CATEGORY_ORDER.forEach(cat => {
          this.saveData(newId, cat, (entry.data && entry.data[cat]) || []);
          // 舊版備份檔裡的 itemMemory／siteItemHistory 不還原（v4.46 起都從資料現算，#29～#32）
          if (entry.siteAliases && entry.siteAliases[cat]) this.saveSiteAliases(newId, cat, entry.siteAliases[cat]);
          if (entry.importBatches && entry.importBatches[cat]) this.saveImportBatches(newId, cat, entry.importBatches[cat]);
        });
        // 每個計畫自己的設定（#33、#34）。舊版備份檔沒有 settings，就用備份檔裡整台電腦共用的那一份
        const st = entry.settings || {};
        const bs = (payload.browserSettings && typeof payload.browserSettings === 'object') ? payload.browserSettings : {};
        const dec = typeof st.decimal === 'string' ? st.decimal : bs[DECIMAL_SETTINGS_KEY];
        if (typeof dec === 'string') this._setItem(this._decimalKey(newId), dec);
        CATEGORY_ORDER.forEach(c => {
          const pr = (st.presets && typeof st.presets[c] === 'string') ? st.presets[c] : bs[`envapp_itemPresets_${c}`];
          if (typeof pr === 'string') this._setItem(this._presetKey(newId, c), pr);
        });
        newProjects.push({ id: newId, code: entry.code, name: entry.name, createdAt: entry.createdAt || new Date().toISOString() });
      });
      this.saveProjects(newProjects);
      // v4.46：小數位數、常用測項記憶標籤已在上面依計畫還原，不再寫入整台電腦共用的設定。
    } catch (err) {
      // roll back everything this call wrote, leave the old projects untouched
      written.forEach(id => { try { this._removeProjectKeys(id); } catch (e) { /* best effort */ } });
      try { this.saveProjects(previousProjects); } catch (e) { /* best effort */ }
      throw err;
    }
    // safe to drop the old data only now that the new list is committed
    if (mode === 'replace') previousProjects.forEach(p => { try { this._removeProjectKeys(p.id); } catch (e) { /* best effort */ } });
  },

  /** Delete every localStorage key belonging to one project (without touching the
   *  project list itself — callers decide when the list changes). */
  _removeProjectKeys(projectId) {
    localStorage.removeItem(this._basicKey(projectId));
    CATEGORY_ORDER.forEach(cat => {
      localStorage.removeItem(this._dataKey(projectId, cat));
      localStorage.removeItem(this._methodMemoryKey(projectId, cat));
      localStorage.removeItem(this._siteItemHistoryKey(projectId, cat));
      localStorage.removeItem(this._siteAliasKey(projectId, cat));
      localStorage.removeItem(this._batchKey(projectId, cat));
      localStorage.removeItem(this._presetKey(projectId, cat));
    });
    localStorage.removeItem(this._decimalKey(projectId));
    // v4.30 的一次性轉換旗標。它不是每個類別各一把，而是整份 JSON 一個物件，
    // 所以要挑掉這個專案的那一項，不能整把刪掉（會影響其他專案）。
    try {
      const key = 'envapp_vibnight_migrated_v1';
      const done = JSON.parse(localStorage.getItem(key)) || {};
      if (done[projectId]) {
        delete done[projectId];
        if (Object.keys(done).length) localStorage.setItem(key, JSON.stringify(done));
        else localStorage.removeItem(key);
      }
    } catch (e) { /* best effort：旗標留著最多就是重問一次，不影響資料 */ }
  },
};
