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
    // An impossible date (2026-02-30, a mis-mapped column) must never be turned
    // into a number: Date.UTC silently rolls it over to a DIFFERENT real date
    // (2026-03-02) and the filing then carries a date nobody ever typed. Write the
    // original text through instead, so it stays visibly wrong and fixable.
    const check = new Date(utcMs);
    if (check.getUTCFullYear() !== y || check.getUTCMonth() + 1 !== mo || check.getUTCDate() !== d) {
      return isoStr || '';
    }
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
    // Out-of-range values must stay text. "24:00" written as the fraction 1.0 is
    // read back by Excel as 1900-01-01 00:00 — the screen said 24:00 and the file
    // says midnight — and "99:99" became 4:39 on 1900-01-03. Neither is a time the
    // person can spot as wrong once it is a number.
    if (mi > 59 || s > 59) return hms || '';
    if (h > 24 || (h === 24 && (mi > 0 || s > 0))) return hms || '';
    const frac = (h * 3600 + mi * 60 + s) / 86400;
    // 24:00:00 is midnight at the END of the day; Excel's h:mm shows 1.0 as 0:00,
    // which is the same clock reading, so store it as 0 rather than as a date.
    return { t: 'n', v: frac >= 1 ? 0 : frac, z: 'h:mm' };
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
      // 噪音（含振動）的監測數值固定兩位小數。這裡再做一次是保險：
      // v4.29 以前存下來、而且使用者從來沒點過那一格的舊資料，交出去的檔案
      // 也必須符合官方「小數點2位數」的規定。補零而已，數值不變。
      if (f.key === NOISE_VALUE_FIELD) return formatNoiseValue(val);
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
