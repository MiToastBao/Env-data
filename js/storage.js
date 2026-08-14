// storage.js
// All data lives in the browser's localStorage. This is a static site
// (designed for GitHub Pages) with no backend, so data is per-browser.
// Use "匯出備份" / "匯入備份" in the app to move data between machines.

const DataStore = {
  PROJECTS_KEY: 'envapp_projects_v1',

  getProjects() {
    try {
      return JSON.parse(localStorage.getItem(this.PROJECTS_KEY)) || [];
    } catch (e) {
      return [];
    }
  },

  saveProjects(projects) {
    localStorage.setItem(this.PROJECTS_KEY, JSON.stringify(projects));
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
    // clean up associated data
    localStorage.removeItem(this._basicKey(id));
    CATEGORY_ORDER.forEach(cat => {
      localStorage.removeItem(this._dataKey(id, cat));
      localStorage.removeItem(this._siteAliasKey(id, cat));
      localStorage.removeItem(this._batchKey(id, cat));
      localStorage.removeItem(this._methodMemoryKey(id, cat));
      localStorage.removeItem(this._siteItemHistoryKey(id, cat));
    });
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

  // Site-item history: remembers, per (project, category, location name), every
  // item name ever confirmed for that site — across all quarters. This is what lets
  // the app notice "上次這個測站還有 def 三項，這次的報告沒出現" and offer to add
  // those as blank rows (with method/unit pre-filled from item memory) instead of the
  // person having to remember and manually recreate columns a PDF-only report can't
  // supply data for.
  getSiteItemHistory(projectId, category) {
    try {
      return JSON.parse(localStorage.getItem(this._siteItemHistoryKey(projectId, category))) || {};
    } catch (e) {
      return {};
    }
  },
  saveSiteItemHistory(projectId, category, history) {
    localStorage.setItem(this._siteItemHistoryKey(projectId, category), JSON.stringify(history));
  },
  /** Adds (location, item) pairs to the history — additive/idempotent, never removes.
   *  Also remembers which 檢測類別 each item was last seen with, per location — e.g.
   *  a noise report site and a vibration report site can share the same location
   *  name but need different 檢測類別 values, so when the app later suggests adding
   *  a missing item back, it must know which category that specific item actually
   *  belongs to rather than guessing from an arbitrary other row at that location.
   *  Stored as { location: { itemName: lastSeenCategory } } — transparently upgrades
   *  the older { location: [itemName, ...] } array format if found. */
  learnSiteItems(projectId, category, locationField, itemField, categoryField, rows) {
    const history = this.getSiteItemHistory(projectId, category);
    rows.forEach(r => {
      const loc = r[locationField];
      const item = r[itemField];
      if (!loc || !item) return;
      if (Array.isArray(history[loc])) {
        const converted = {};
        history[loc].forEach(it => { converted[it] = ''; });
        history[loc] = converted;
      }
      if (!history[loc]) history[loc] = {};
      const catVal = categoryField ? (r[categoryField] || '') : '';
      if (catVal || !(item in history[loc])) history[loc][item] = catVal;
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
    localStorage.setItem(this._methodMemoryKey(projectId, category), JSON.stringify(memory));
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
    localStorage.setItem(this._batchKey(projectId, category), JSON.stringify(batches));
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
    localStorage.setItem(this._siteAliasKey(projectId, category), JSON.stringify(aliases));
  },

  getBasicInfo(projectId) {
    try {
      return JSON.parse(localStorage.getItem(this._basicKey(projectId))) || {};
    } catch (e) {
      return {};
    }
  },
  saveBasicInfo(projectId, info) {
    localStorage.setItem(this._basicKey(projectId), JSON.stringify(info));
  },

  getData(projectId, category) {
    try {
      return JSON.parse(localStorage.getItem(this._dataKey(projectId, category))) || [];
    } catch (e) {
      return [];
    }
  },
  saveData(projectId, category, rows) {
    localStorage.setItem(this._dataKey(projectId, category), JSON.stringify(rows));
  },
  clearData(projectId, category) {
    this.saveData(projectId, category, []);
    this.saveImportBatches(projectId, category, []);
  },

  // Full export/import for backup/transfer between browsers
  exportAll() {
    const projects = this.getProjects();
    const out = { projects: [] };
    projects.forEach(p => {
      const entry = { ...p, basicInfo: this.getBasicInfo(p.id), data: {}, itemMemory: {}, siteItemHistory: {}, siteAliases: {} };
      CATEGORY_ORDER.forEach(cat => {
        entry.data[cat] = this.getData(p.id, cat);
        entry.itemMemory[cat] = this.getItemMemory(p.id, cat);
        entry.siteItemHistory[cat] = this.getSiteItemHistory(p.id, cat);
        entry.siteAliases[cat] = this.getSiteAliases(p.id, cat);
      });
      out.projects.push(entry);
    });
    return out;
  },

  importAll(payload, mode = 'merge') {
    if (!payload || !Array.isArray(payload.projects)) throw new Error('備份檔格式不正確');
    if (mode === 'replace') {
      const existing = this.getProjects();
      existing.forEach(p => this.deleteProject(p.id));
    }
    const projects = this.getProjects();
    payload.projects.forEach(entry => {
      const newId = 'p_' + Date.now() + '_' + Math.random().toString(36).slice(2, 8);
      projects.push({ id: newId, code: entry.code, name: entry.name, createdAt: entry.createdAt || new Date().toISOString() });
      this.saveBasicInfo(newId, entry.basicInfo || {});
      CATEGORY_ORDER.forEach(cat => {
        this.saveData(newId, cat, (entry.data && entry.data[cat]) || []);
        if (entry.itemMemory && entry.itemMemory[cat]) this.saveItemMemory(newId, cat, entry.itemMemory[cat]);
        if (entry.siteItemHistory && entry.siteItemHistory[cat]) this.saveSiteItemHistory(newId, cat, entry.siteItemHistory[cat]);
        if (entry.siteAliases && entry.siteAliases[cat]) this.saveSiteAliases(newId, cat, entry.siteAliases[cat]);
      });
    });
    this.saveProjects(projects);
  },
};
