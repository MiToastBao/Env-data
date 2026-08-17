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
    // Compute the Excel date SERIAL NUMBER directly (days since the 1900 date
    // system's epoch) instead of handing SheetJS a JS Date object. Confirmed as a
    // real bug: constructing `new Date(Date.UTC(y, m-1, d))` and letting SheetJS
    // convert it produced a cell whose underlying value showed a non-zero time
    // (e.g. "08:00:00 AM" in the formula bar for a plain "2026-05-13" date) —
    // SheetJS's Date->serial conversion reads the JS Date through local-timezone
    // getters internally, so a UTC-midnight Date silently picked up the browser's
    // UTC+8 offset as a fractional day. Computing the serial via pure arithmetic
    // on the calendar numbers actually typed in sidesteps that path entirely: no
    // Date object is ever handed to SheetJS, so no timezone conversion can happen.
    const y = +m[1], mo = +m[2], d = +m[3];
    // Days between this date and 1899-12-31 (Excel's day-0 reference under the
    // 1900 date system), computed via UTC millisecond arithmetic — both sides are
    // UTC timestamps, so the subtraction itself is timezone-agnostic.
    const utcMs = Date.UTC(y, mo - 1, d);
    const epochMs = Date.UTC(1899, 11, 31);
    let serial = Math.round((utcMs - epochMs) / 86400000);
    // Excel (for backward compatibility with a Lotus 1-2-3 bug) treats 1900 as a
    // leap year even though it isn't one — every real date on or after 1900-03-01
    // needs its serial bumped by 1 to match Excel's actual numbering.
    if (serial > 59) serial += 1;
    return { t: 'n', v: serial, z: 'yyyy/mm/dd' };
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
