// export.js
// Builds an .xlsx that matches the government blank-template structure
// exactly: sheet 1 "監測點基本資料" (one header row + one data row),
// sheet 2 "<category>檢測項目" (header row + all monitoring data rows).
//
// Date/time fields are written as REAL Excel date/time cells (not text), matching
// how the official completed template itself stores them — confirmed by inspecting
// a real filed report: date cells are true date-typed with format "mm-dd-yy", time
// cells are true time-typed with format "h:mm" (no seconds). Writing plain strings
// instead (which is what earlier versions of this tool did) could cause a downstream
// system to treat the date as ordinary text rather than a real date value.

const ExportEngine = {
  /** "YYYY-MM-DD" -> a SheetJS date cell object, or '' passthrough if not a valid date. */
  _dateCell(isoStr) {
    const m = String(isoStr || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!m) return isoStr || '';
    const d = new Date(Date.UTC(+m[1], +m[2] - 1, +m[3]));
    return { t: 'd', v: d, z: 'yyyy/mm/dd' };
  },
  /** "HH:MM:SS" -> a SheetJS time cell object (stored as a fraction-of-day number,
   *  which is how Excel represents time-only values internally), or '' passthrough. */
  _timeCell(hms) {
    const m = String(hms || '').match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
    if (!m) return hms || '';
    const h = +m[1], mi = +m[2], s = +(m[3] || 0);
    const frac = (h * 3600 + mi * 60 + s) / 86400;
    return { t: 'n', v: frac, z: 'h:mm' };
  },

  buildWorkbook(project, basicInfo, categoryKey, rowsOverride) {
    const cat = CATEGORIES[categoryKey];
    const rows = rowsOverride || DataStore.getData(project.id, categoryKey);

    const wb = XLSX.utils.book_new();

    // Sheet 1: 監測點基本資料
    const basicHeaders = BASIC_INFO_FIELDS.map(f => f.key);
    const basicRow = BASIC_INFO_FIELDS.map(f => {
      const val = basicInfo[f.key] || '';
      return f.type === 'date' ? this._dateCell(val) : val;
    });
    const wsBasic = XLSX.utils.aoa_to_sheet([basicHeaders, basicRow]);
    XLSX.utils.book_append_sheet(wb, wsBasic, cat.basicSheetName);

    // Sheet 2: <category>檢測項目
    const dataHeaders = cat.fields.map(f => f.key);
    const dataRows = rows.map(r => cat.fields.map(f => {
      const val = r[f.key] || '';
      if (f.type === 'date') return this._dateCell(val);
      if (f.type === 'time') return this._timeCell(val);
      return val;
    }));
    const wsData = XLSX.utils.aoa_to_sheet([dataHeaders, ...dataRows]);
    XLSX.utils.book_append_sheet(wb, wsData, cat.dataSheetName);

    return wb;
  },

  downloadCategory(project, basicInfo, categoryKey, rowsOverride) {
    const cat = CATEGORIES[categoryKey];
    const wb = this.buildWorkbook(project, basicInfo, categoryKey, rowsOverride);
    const fname = `${project.code}_${cat.sourceFile}`;
    XLSX.writeFile(wb, fname);
  },

  downloadAll(project) {
    const basicInfo = DataStore.getBasicInfo(project.id);
    CATEGORY_ORDER.forEach(catKey => {
      const rows = DataStore.getData(project.id, catKey);
      if (rows.length > 0) {
        this.downloadCategory(project, basicInfo, catKey);
      }
    });
  },
};
