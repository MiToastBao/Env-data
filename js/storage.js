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

  // Full export/import for backup/transfer between browsers
  exportAll() {
    const projects = this.getProjects();
    const out = { projects: [] };
    projects.forEach(p => {
      const entry = { ...p, basicInfo: this.getBasicInfo(p.id), data: {} };
      CATEGORY_ORDER.forEach(cat => {
        entry.data[cat] = this.getData(p.id, cat);
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
      });
    });
    this.saveProjects(projects);
  },
};
