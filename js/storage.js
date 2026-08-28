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

  // Site-item history: remembers, per (project, category, location name), a full
  // field snapshot of every item ever confirmed for that site — across all
  // quarters. This is what lets the app notice "上次這個測站還有 def 三項，這次的
  // 報告沒出現" and offer to add those back with all their real detail (座標/管制
  // 標準/檢測方法/單位/備註 etc, not just method/unit) intact, leaving only date/
  // time/value blank for the person to fill in.
  getSiteItemHistory(projectId, category) {
    try {
      return JSON.parse(localStorage.getItem(this._siteItemHistoryKey(projectId, category))) || {};
    } catch (e) {
      return {};
    }
  },
  saveSiteItemHistory(projectId, category, history) {
    this._setItem(this._siteItemHistoryKey(projectId, category), JSON.stringify(history));
  },
  /** Adds (location, item) pairs to the history — additive/idempotent, never removes.
   *  Also remembers which 檢測類別 each item was last seen with, per location — e.g.
   *  a noise report site and a vibration report site can share the same location
   *  name but need different 檢測類別 values, so when the app later suggests adding
   *  a missing item back, it must know which category that specific item actually
   *  belongs to rather than guessing from an arbitrary other row at that location.
   *  Stored as { location: { itemName: lastSeenCategory } } — transparently upgrades
   *  the older { location: [itemName, ...] } array format if found. */
  /**
   * Records a full field snapshot for each (location, item-identity) pair — additive
   * and always overwritten with the most recent confirmed import, never removes.
   * Stored as { location: { identityKey: { itemName, timeSegment, category, snapshot, count } } }
   * where `snapshot` holds every OTHER field's value (座標/管制標準/檢測方法/單位/
   * 備註 etc — everything except the location/item identity fields and 日期/時間/
   * 檢測數值), and `count` is how many rows shared that exact identity — e.g. a site
   * sampled both 平日 and 假日 in the same quarter legitimately has TWO rows with
   * the identical "音源發聲特性::監測時段" identity (same item, same time-of-day,
   * different sampling date). Without tracking count, the second row learned would
   * simply overwrite the first and the app would only ever remember/rebuild one of
   * the two. `entries` should reflect the FULL current picture for each touched
   * location (see learnSiteItemHistory in app.js, which re-derives this from
   * complete DataStore data rather than whatever partial row list triggered the
   * update) — that's what lets `count` be computed by just counting how many times
   * an identityKey appears in `entries`, no separate bookkeeping needed.
   * This is what lets the app rebuild a historically-known-but-currently-missing
   * row (or set of rows) with all its real detail intact — not just item name and
   * category — when the person asks to add it back, leaving only date/time/value
   * blank for them to fill in. `identityKey` folds in 監測時段 (day/evening/night)
   * when the category has that field, so noise's three time-of-day rows per item
   * don't collapse into one indistinguishable entry.
   * Transparently discards the older, narrower history formats if found (both the
   * original { location: [itemName,...] } array and the interim
   * { location: { itemName: category } } object) — those didn't carry enough
   * detail to reconstruct a full row anyway, so there's nothing worth preserving
   * from them; the next confirmed import repopulates correctly.
   */
  learnSiteItemSnapshots(projectId, category, entries) {
    const history = this.getSiteItemHistory(projectId, category);
    // Group first so repeated (location, identityKey) pairs within `entries`
    // (the 平日/假日 case) accumulate into a count rather than the last one
    // silently overwriting the rest.
    const grouped = {}; // location -> identityKey -> { itemName, timeSegment, category, snapshot, count, dates }
    entries.forEach(({ location, identityKey, itemName, timeSegment, itemCategory, snapshot, date }) => {
      if (!location || !identityKey) return;
      if (!grouped[location]) grouped[location] = {};
      if (!grouped[location][identityKey]) {
        grouped[location][identityKey] = { itemName, timeSegment: timeSegment || '', category: itemCategory || '', snapshot, count: 0, dates: [] };
      }
      grouped[location][identityKey].count += 1;
      // The DISTINCT sampling dates behind that count, not just the count itself.
      // Three rows can mean "three monthly visits" or "one visit measured three
      // ways", and only the dates tell them apart — which is what stops a later
      // single-month import from being told it is "missing 2 readings" purely
      // because the quarter it is compared against covered three months.
      if (date && !grouped[location][identityKey].dates.includes(date)) {
        grouped[location][identityKey].dates.push(date);
      }
      grouped[location][identityKey].snapshot = snapshot; // keep the most recently seen snapshot
    });
    Object.entries(grouped).forEach(([location, identities]) => {
      if (!history[location] || Array.isArray(history[location]) || typeof history[location] !== 'object') {
        history[location] = {};
      }
      Object.entries(identities).forEach(([identityKey, entry]) => {
        history[location][identityKey] = entry;
      });
    });
    this.saveSiteItemHistory(projectId, category, history);
  },

  // Item memory: remembers, per (project, category, item name), the last-confirmed
  // unit code and test method — independent of which quarter/file it came from. This
  // is what lets next season's import default to what was already established this
  // season, without needing the person to track which period supplied which value.
  getItemMemory(projectId, category) {
    try {
      return JSON.parse(localStorage.getItem(this._methodMemoryKey(projectId, category))) || {};
    } catch (e) {
      return {};
    }
  },
  saveItemMemory(projectId, category, memory) {
    this._setItem(this._methodMemoryKey(projectId, category), JSON.stringify(memory));
  },
  /** Merge in newly-confirmed per-item values (only overwrites a field if the new value is non-empty). */
  updateItemMemory(projectId, category, itemUpdates) {
    const memory = this.getItemMemory(projectId, category);
    Object.entries(itemUpdates).forEach(([itemName, fields]) => {
      memory[itemName] = memory[itemName] || {};
      Object.entries(fields).forEach(([key, val]) => {
        if (val) memory[itemName][key] = val;
      });
      memory[itemName]._updatedAt = new Date().toISOString();
    });
    this.saveItemMemory(projectId, category, memory);
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
    return [DECIMAL_SETTINGS_KEY, ...CATEGORY_ORDER.map(c => `envapp_itemPresets_${c}`)];
  },

  // Full export/import for backup/transfer between browsers
  exportAll() {
    const projects = this.getProjects();
    const out = { projects: [], browserSettings: {} };
    this._browserSettingKeys().forEach((k) => {
      const v = localStorage.getItem(k);
      if (v !== null) out.browserSettings[k] = v;
    });
    projects.forEach(p => {
      const entry = { ...p, basicInfo: this.getBasicInfo(p.id), data: {}, itemMemory: {}, siteItemHistory: {}, siteAliases: {}, importBatches: {} };
      CATEGORY_ORDER.forEach(cat => {
        entry.data[cat] = this.getData(p.id, cat);
        entry.itemMemory[cat] = this.getItemMemory(p.id, cat);
        entry.siteItemHistory[cat] = this.getSiteItemHistory(p.id, cat);
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
          if (entry.itemMemory && entry.itemMemory[cat]) this.saveItemMemory(newId, cat, entry.itemMemory[cat]);
          if (entry.siteItemHistory && entry.siteItemHistory[cat]) this.saveSiteItemHistory(newId, cat, entry.siteItemHistory[cat]);
          if (entry.siteAliases && entry.siteAliases[cat]) this.saveSiteAliases(newId, cat, entry.siteAliases[cat]);
          if (entry.importBatches && entry.importBatches[cat]) this.saveImportBatches(newId, cat, entry.importBatches[cat]);
        });
        newProjects.push({ id: newId, code: entry.code, name: entry.name, createdAt: entry.createdAt || new Date().toISOString() });
      });
      this.saveProjects(newProjects);
      /*
       * 瀏覽器共用設定。
       *
       * 「取代」＝這台電腦要變成備份那台的樣子，所以照抄。
       * 「合併」＝這台電腦本來就有自己的專案在用，**只補上這裡還沒有的**，
       *   不去動已經存在的設定——否則把同事的備份併進來，會順手改掉自己
       *   原有專案的小數位數，而那會直接改變匯出的檔案內容。
       * 舊版備份檔沒有這一段，跳過就是了。
       */
      const bs = payload.browserSettings;
      if (bs && typeof bs === 'object') {
        this._browserSettingKeys().forEach((k) => {
          if (typeof bs[k] !== 'string') return;
          if (mode !== 'replace' && localStorage.getItem(k) !== null) return;
          // 盡力就好。設定寫不進去（例如容量滿了）會由 _setItem 通知使用者，
          // 但**不能因此把已經還原好的專案資料整批退掉**——資料是本體，
          // 設定是附帶的。
          try { this._setItem(k, bs[k]); } catch (e) { /* 已通知，繼續 */ }
        });
      }
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
    });
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
