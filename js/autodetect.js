// autodetect.js
// "I've never seen this report layout before" fallback reader.
//
// WHY THIS EXISTS
// ---------------
// smartparse.js contains parsers written against specific, known report templates
// (華光-style noise/water/air/底泥 forms). They are precise, but they only fire on
// layouts they were written for — hand an unfamiliar lab's report to the app and it
// used to find nothing at all, leaving the person to map columns by hand on a file
// that has no header row to map.
//
// This module takes the opposite approach: it doesn't know any template. It knows
// what the SEVEN fields the person actually needs LOOK like, and hunts for them:
//
//     日期(起) 日期(迄) 時間(起) 時間(迄) 檢測項目 監測數值 檢測單位
//
// plus, when they happen to be present, 採樣地點 / 檢測方法 / 檢測極限 / 檢測機構 /
// 檢測類別, because those are cheap to pick up once the table has been located.
//
// It works in two layers, and uses both together:
//   1. FORM LAYER — "採樣日期：115年06月25日" style label/value pairs scattered around
//      the top of a report sheet. These give the sampling date/time/location/agency
//      that apply to the whole sheet.
//   2. TABLE LAYER — a band of column headings (which may be spread over two or three
//      merged rows) followed by one row per test item. These give item/value/unit/
//      method/limit per row.
// Whatever the table layer doesn't supply, the form layer fills in.
//
// Anything produced here is flagged `_autoDetected` so the import preview can say
// loudly that these values were GUESSED from an unrecognized layout and need to be
// checked by a human, rather than presenting them with the same confidence as a
// value read from a known template.

const AutoDetect = {

  // ---------- what each of the wanted fields looks like as a column heading ----------
  // Order matters: the first rule that matches wins, so the more specific
  // "日期(起)" rules must sit above the generic "日期" one.
  HEADER_RULES: [
    { key: 'dateStart', re: /^(?:採樣|監測|檢測|調查|檢驗|取樣)?日期.*(?:起|開始|start)|^(?:起始|開始)日期|^start\s*date/i },
    { key: 'dateEnd', re: /^(?:採樣|監測|檢測|調查|檢驗|取樣)?日期.*(?:迄|訖|結束|end)|^(?:結束|終止)日期|^end\s*date/i },
    { key: 'timeStart', re: /^(?:採樣|監測|檢測|調查|檢驗|取樣)?時間.*(?:起|開始|start)|^(?:起始|開始)時間|^start\s*time/i },
    { key: 'timeEnd', re: /^(?:採樣|監測|檢測|調查|檢驗|取樣)?時間.*(?:迄|訖|結束|end)|^(?:結束|終止)時間|^end\s*time/i },
    { key: 'dateRange', re: /^(?:採樣|監測|檢測|調查)?(?:日期|期間)$|^監測日期$|^採樣期間$/ },
    { key: 'timeRange', re: /^(?:採樣|監測|檢測|調查|測定)?時間$|^測定時間$|^監測時間$/ },
    { key: 'item', re: /^(?:檢驗|檢測|監測|調查|分析|測定|試驗)?項\s*目$|^測項$|^項目名稱$|^parameter$|^item$|^analyte$/i },
    { key: 'value', re: /檢測值|測定值|監測值|檢測數值|監測數值|分析值|檢驗值|^測值$|^數值$|^結果$|檢測結果|分析結果|^result$|^value$|^conc/i },
    { key: 'limit', re: /偵測極限|檢測極限|定量極限|方法偵測極限|^m\.?d\.?l\.?$|^r\.?l\.?$|^loq$/i },
    { key: 'unit', re: /^單\s*位$|濃度單位|質量單位|檢測單位|監測單位|單位代碼|^unit$/i },
    { key: 'method', re: /檢測方法|分析方法|檢驗方法|監測方法|採樣方法|^方\s*法$|^method$/i },
    { key: 'location', re: /採樣地點|監測地點|調查地點|測點名稱|測站名稱|^地\s*點$|^測\s*站$|^測\s*點$|^站\s*名$|^location$|^site$/i },
    { key: 'agency', re: /檢測機構|檢驗機構|受檢單位|檢驗室名稱|許可證號/ },
    { key: 'category', re: /檢測類別|監測類別|樣品特性|樣品性質|^類\s*別$/ },
    { key: 'compare', re: /比較關係|比較符號/ },
    { key: 'remark', re: /^備\s*註$|^說\s*明$|^remark$|^note$/i },
  ],

  /** Header text -> one of the keys above, or null.
   *
   *  Report tables merge cells freely, so one physical cell can hold a heading split
   *  across lines ("偵測\n極限") or two headings stacked together ("樣品編號\n檢驗項目").
   *  Each newline-separated piece is therefore tested on its own as well as the whole
   *  collapsed string — otherwise a perfectly ordinary 檢驗項目 column goes unseen
   *  purely because the lab merged it with the sample-id caption above it. */
  matchHeader(text) {
    const raw = String(text ?? '');
    const collapse = (t) => t.replace(/\s+/g, '').replace(/（/g, '(').replace(/）/g, ')').trim();
    const candidates = [collapse(raw), ...raw.split(/\r?\n/).map(collapse)];
    for (const s of candidates) {
      if (s === '' || s.length > 24) continue;
      if (/^-?[\d.,%]+$/.test(s)) continue; // a bare number is data, never a heading
      // strip a trailing unit annotation a heading may carry, e.g. "檢測值(mg/L)"
      const bare = s.replace(/\([^)]*\)\s*$/, '') || s;
      for (const rule of this.HEADER_RULES) {
        if (rule.re.test(s) || rule.re.test(bare)) return rule.key;
      }
    }
    return null;
  },

  // ---------- form-layer labels ("採樣日期：...") ----------
  META_RULES: [
    { key: 'dateRange', re: /^(採樣|監測|檢測|調查|檢驗|取樣)?(日期|期間)$/ },
    { key: 'timeRange', re: /^(採樣|監測|檢測|調查|測定|檢驗)?時間$/ },
    { key: 'location', re: /^(採樣|監測|調查|檢測)?地點$|^測點(編號|名稱)?$|^測站(編號|名稱)?$/ },
    { key: 'siteCode', re: /^測點編號$|^採樣點編號$/ },
    { key: 'agency', re: /^(採樣|檢測|檢驗|受檢)單位$|^檢驗室名稱$|^公司名稱$|^檢測機構$/ },
    { key: 'method', re: /^(採樣|檢測|分析|檢驗)方法$/ },
    { key: 'category', re: /^樣品(特性|性質|種類)$|^檢測類別$|^監測類別$/ },
  ],

  cellStr(v) { return v === undefined || v === null ? '' : String(v).trim(); },

  /**
   * Collects every "label：value" pair on the sheet. Handles both shapes seen in real
   * reports: the value in the SAME cell after the colon, and the label in one cell
   * with the value in the next non-empty cell to its right.
   */
  extractFormMeta(grid) {
    const meta = {};
    const put = (key, val) => {
      const v = this.cellStr(val);
      if (v !== '' && !meta[key]) meta[key] = v;
    };
    for (let r = 0; r < grid.length; r++) {
      const row = grid[r] || [];
      for (let c = 0; c < row.length; c++) {
        const cell = this.cellStr(row[c]);
        if (cell === '' || cell.length > 60) continue;
        const colonIdx = cell.search(/[:：]/);
        let label = cell, inlineValue = '';
        if (colonIdx >= 0) {
          label = cell.slice(0, colonIdx);
          inlineValue = cell.slice(colonIdx + 1).trim();
        }
        label = label.replace(/\s+/g, '');
        if (label === '' || label.length > 12) continue;
        const rule = this.META_RULES.find(m => m.re.test(label));
        if (!rule) continue;
        if (inlineValue !== '') { put(rule.key, inlineValue); continue; }
        if (colonIdx < 0) continue; // a bare word with no colon is too weak a signal
        for (let cc = c + 1; cc < row.length; cc++) {
          const v = this.cellStr(row[cc]);
          if (v !== '') { put(rule.key, v); break; }
        }
      }
    }
    return meta;
  },

  /** "115.01.30~115.02.01" / "2/12 ~ 2/13" / a single date -> {start, end} ISO. */
  splitDateRange(text, fallbackYear) {
    const s = String(text ?? '');
    if (s.trim() === '') return { start: '', end: '' };
    const parts = s.split(/\s*[~～至－—]\s*|\s*--\s*/);
    if (parts.length >= 2) {
      const a = DateTimeUtil.toISODate(parts[0]);
      let b = DateTimeUtil.toISODate(parts[1]);
      if (/^\d{4}-\d{2}-\d{2}$/.test(a)) {
        if (!/^\d{4}-\d{2}-\d{2}$/.test(b)) {
          // "2/12 ~ 2/13" — the second half has no year of its own
          const md = String(parts[1]).match(/(\d{1,2})\s*[/月-]\s*(\d{1,2})/);
          if (md) {
            const y = parseInt(a.slice(0, 4), 10);
            const pad = n => String(n).padStart(2, '0');
            const y2 = parseInt(md[1], 10) < parseInt(a.slice(5, 7), 10) ? y + 1 : y;
            b = `${y2}-${pad(md[1])}-${pad(md[2])}`;
          } else b = '';
        }
        return { start: a, end: b || a };
      }
      // neither half carried a year: "2/12 ~ 2/13" with the year only in a 報告日期
      if (fallbackYear) {
        const pad = n => String(n).padStart(2, '0');
        const m1 = String(parts[0]).match(/(\d{1,2})\s*[/月-]\s*(\d{1,2})/);
        const m2 = String(parts[1]).match(/(\d{1,2})\s*[/月-]\s*(\d{1,2})/);
        if (m1 && m2) {
          const y2 = parseInt(m2[1], 10) < parseInt(m1[1], 10) ? fallbackYear + 1 : fallbackYear;
          return { start: `${fallbackYear}-${pad(m1[1])}-${pad(m1[2])}`, end: `${y2}-${pad(m2[1])}-${pad(m2[2])}` };
        }
      }
    }
    const only = DateTimeUtil.toISODate(s);
    return /^\d{4}-\d{2}-\d{2}$/.test(only) ? { start: only, end: only } : { start: '', end: '' };
  },

  /** "10:05~10:07" / "12時00分" / "6~20" -> {start, end} as HH:MM:00. */
  splitTimeRange(text) {
    const s = String(text ?? '');
    if (s.trim() === '') return { start: '', end: '' };
    const m = s.match(/(\d{1,2}[:：]\d{2})\s*[~～至\-—]\s*(\d{1,2}[:：]\d{2})/);
    if (m) return { start: DateTimeUtil.toHMS(m[1].replace('：', ':')), end: DateTimeUtil.toHMS(m[2].replace('：', ':')) };
    const one = DateTimeUtil.toHMS(s);
    return /^\d{2}:\d{2}:\d{2}$/.test(one) ? { start: one, end: one } : { start: '', end: '' };
  },

  // ---------- table layer ----------

  /**
   * Finds the band of rows that carries the column headings and returns
   * { headerRows:[...], dataStart, colMap:{colIndex:key}, keys:Set }.
   *
   * A "band" rather than a single row because report tables frequently split a
   * heading across two merged rows ("偵測" / "極限") or put the value column's own
   * heading one row below the rest (e.g. the sample id sits in the heading row and
   * the literal word 檢測值 sits underneath it). Scanning a 3-row window and keeping
   * the first key found per column handles both without special-casing either.
   */
  findHeaderBand(grid) {
    const limit = Math.min(grid.length, 60);
    let best = null;
    for (let r = 0; r < limit; r++) {
      const colMap = {};
      const keys = new Set();
      let lastLabelRow = r;
      for (let rr = r; rr < Math.min(r + 3, grid.length); rr++) {
        const row = grid[rr] || [];
        // A row already carrying measurements is data, not part of the heading band —
        // reading labels out of it would both mis-map columns and swallow the first
        // row of real results.
        const numericCells = row.filter(c => /^-?[\d.,]+$/.test(this.cellStr(c)) && this.cellStr(c) !== '').length;
        if (rr > r && numericCells >= 2) break;
        for (let c = 0; c < row.length; c++) {
          if (colMap[c] !== undefined) continue;
          const key = this.matchHeader(row[c]);
          if (!key) continue;
          if (keys.has(key)) continue; // first column wins a given role
          colMap[c] = key;
          keys.add(key);
          if (rr > lastLabelRow) lastLabelRow = rr;
        }
      }
      if (keys.size < 2) continue;
      // A usable table needs, at minimum, something to name the measurement and
      // something to hold it. Without both, this is a stray label row, not a table.
      if (!(keys.has('item') || keys.has('location')) || !keys.has('value')) {
        // a wide "one row per sample, columns are dates/times" table is still useful
        // if it carries date/time columns plus a value column
        if (!(keys.has('value') && (keys.has('dateStart') || keys.has('dateRange')))) continue;
      }
      const score = keys.size - r * 0.05;
      if (!best || score > best.score) {
        best = { headerRow: r, dataStart: lastLabelRow + 1, colMap, keys, score };
      }
    }
    return best;
  },

  /** Rows at/after which a report stops being data and starts being boilerplate. */
  STOP_ROW_RE: /以下空白|以下\s*空白|^備\s*註|聲明書|本報告|檢驗室主管|報告簽署人|負責人|^公司名稱|第\s*\d+\s*頁|簽章/,

  /**
   * The main entry point: try to pull rows for `category` out of one unknown sheet.
   * Returns an array of schema-shaped rows, or null when nothing usable was found.
   */
  parseSheet(category, sheetName, grid) {
    if (!grid || grid.length === 0) return null;
    const cat = CATEGORIES[category];
    if (!cat) return null;

    const meta = this.extractFormMeta(grid);
    const band = this.findHeaderBand(grid);
    if (!band) return null;

    // A year to fall back on when a date range in the table omits it (very common:
    // "2/12 ~ 2/13" with the year only stated once in 報告日期 at the top).
    let fallbackYear = null;
    for (const key of ['dateRange', 'reportDate']) {
      const iso = DateTimeUtil.toISODate(meta[key] || '');
      if (/^\d{4}-/.test(iso)) { fallbackYear = parseInt(iso.slice(0, 4), 10); break; }
    }
    if (!fallbackYear) {
      const anyDate = this.cellStr(grid.flat().find(c => /(\d{2,4})\s*年\s*\d{1,2}\s*月/.test(this.cellStr(c))) || '');
      const iso = DateTimeUtil.toISODate(anyDate);
      if (/^\d{4}-/.test(iso)) fallbackYear = parseInt(iso.slice(0, 4), 10);
    }

    // sheet-level defaults from the form layer. A "採樣時間：115年06月25日12時00分"
    // label carries BOTH the date and the clock time, so it feeds both readers.
    const metaDates = this.splitDateRange(meta.dateRange || meta.timeRange || '', fallbackYear);
    const metaTimes = this.splitTimeRange(meta.timeRange || meta.dateRange || '');
    const metaAgency = meta.agency ? SmartParse.reverseAgencyLookup(meta.agency) : '';
    const metaMethod = meta.method ? SmartParse.extractMethodCode(meta.method) : '';
    const metaCategory = meta.category ? this.guessCategoryValue(category, meta.category) : '';
    const metaLocation = meta.location || '';

    const valueFieldKey = cat.fields.some(f => f.key === '檢測數值') ? '檢測數值'
      : cat.fields.some(f => f.key === '監測數值') ? '監測數值' : null;
    const hasField = (k) => cat.fields.some(f => f.key === k);

    const byKey = {};
    Object.entries(band.colMap).forEach(([c, key]) => { byKey[key] = Number(c); });

    // A unit printed once in the heading — "檢測值(mg/L)" or a dedicated 單位 row —
    // applies to every row in that column when there's no per-row 單位 column.
    let headerUnitText = '';
    if (byKey.value !== undefined && byKey.unit === undefined) {
      for (let rr = band.headerRow; rr < band.dataStart; rr++) {
        const h = this.cellStr((grid[rr] || [])[byKey.value]);
        const m = h.match(/[(（]([^)）]+)[)）]\s*$/);
        if (m) { headerUnitText = m[1].trim(); break; }
      }
    }

    const rows = [];
    for (let r = band.dataStart; r < grid.length; r++) {
      const row = grid[r] || [];
      const joined = row.map(c => this.cellStr(c)).join('');
      if (joined === '') continue;
      if (this.STOP_ROW_RE.test(joined) && rows.length > 0) break;

      const get = (key) => (byKey[key] === undefined ? '' : this.cellStr(row[byKey[key]]));
      const itemName = SmartParse.normalizeItemName(get('item'));
      const valueRaw = get('value');
      if (byKey.item !== undefined && itemName === '') continue;
      if (byKey.value !== undefined && valueRaw === '') continue;
      if (itemName === '' && valueRaw === '') continue;

      const rowDates = byKey.dateStart !== undefined || byKey.dateRange !== undefined
        ? this.splitDateRange(get('dateStart') || get('dateRange'), fallbackYear) : { start: '', end: '' };
      const explicitEnd = DateTimeUtil.toISODate(get('dateEnd'));
      const rowTimes = byKey.timeStart !== undefined || byKey.timeRange !== undefined
        ? this.splitTimeRange(get('timeStart') || get('timeRange')) : { start: '', end: '' };
      const explicitTimeEnd = DateTimeUtil.toHMS(get('timeEnd'));

      const dateStart = rowDates.start || metaDates.start || '';
      const dateEnd = (/^\d{4}-\d{2}-\d{2}$/.test(explicitEnd) ? explicitEnd : '')
        || rowDates.end || metaDates.end || dateStart;
      const timeStart = rowTimes.start || metaTimes.start || '';
      const timeEnd = (/^\d{2}:\d{2}/.test(explicitTimeEnd) ? explicitTimeEnd : '')
        || rowTimes.end || metaTimes.end || timeStart;

      const { cmp, val, note } = SmartParse.parseValueCell(valueRaw);
      const unitText = get('unit') || headerUnitText;
      const unitLookup = unitText ? SmartParse.reverseUnitLookup(unitText, itemName) : { code: '', confident: false };
      const methodText = SmartParse.extractMethodCode(get('method')) || metaMethod;
      const limitRaw = get('limit').replace(/[-—─–]{1,2}/g, '').trim();
      const limitApplies = cmp === 'ND' || cmp === '<';
      const location = get('location') || metaLocation;
      const agencyText = get('agency');
      const agencyCode = agencyText
        ? (AGENCY_CODES[agencyText] ? agencyText : SmartParse.reverseAgencyLookup(agencyText))
        : metaAgency;

      const out = {};
      cat.fields.forEach(f => { out[f.key] = ''; });
      out['日期(起)'] = dateStart;
      out['日期(迄)'] = dateEnd;
      out['時間(起)'] = timeStart;
      out['時間(迄)'] = timeEnd;
      out[cat.locationField] = location;
      out[cat.itemField] = itemName || (valueFieldKey ? '（未標示）' : '');
      if (valueFieldKey) out[valueFieldKey] = /^[\d.]+$/.test(val) ? SmartParse.formatNumber(val, 3) : val;
      if (hasField('比較關係')) out['比較關係'] = cmp;
      if (hasField('檢測極限')) out['檢測極限'] = limitApplies && /^[\d.]+$/.test(limitRaw) ? SmartParse.formatNumber(limitRaw, 3) : '';
      if (cat.unitField) out[cat.unitField] = unitLookup.code;
      if (cat.methodField) out[cat.methodField] = methodText;
      // 地質 deliberately leaves 檢測類別 for the person to choose (底泥品質 vs 土壤品質
      // vs 廢棄物 is a filing decision, and a wrong auto-filled value is harder to
      // spot than an obviously empty one) — same rule as parseLabItemTableSheet.
      if (hasField('檢測類別') && category !== 'geo') {
        out['檢測類別'] = this.guessCategoryValue(category, get('category')) || metaCategory;
      }
      out['檢測機構許可證號'] = agencyCode;
      out['備註'] = note || '';

      out._siteCode = meta.siteCode || '';
      out._rawLocation = location;
      out._autoDetected = true;
      out._uncertainUnit = !!unitText && !unitLookup.confident;
      rows.push(out);
    }

    if (rows.length === 0) return null;
    const detectedFields = this.describeDetection(rows, cat);
    rows.forEach(r => { r._autoDetectInfo = detectedFields; });
    return rows;
  },

  /** Maps free text ("底泥", "河川水") onto one of the category's 檢測類別 options. */
  guessCategoryValue(category, text) {
    const s = String(text ?? '').trim();
    if (s === '') return '';
    const options = CATEGORY_TYPE_OPTIONS[category] || [];
    const exact = options.find(o => o === s);
    if (exact) return exact;
    const contains = options.find(o => s.includes(o) || o.includes(s));
    if (contains) return contains;
    if (category === 'water') return SmartParse.sampleTypeToCategory(s);
    if (category === 'geo') return SmartParse.sedimentTypeToCategory(s);
    return '';
  },

  /** Reports what the detector ACTUALLY managed to fill in, judged from the finished
   *  rows rather than from which headings it spotted — a date can arrive from a
   *  "採樣時間：115年06月25日12時00分" label with no 日期 column anywhere, and telling
   *  the person that field is "missing" when it's sitting right there is worse than
   *  useless. A field counts as found when at least one row carries a value. */
  describeDetection(rows, cat) {
    const valueField = cat.fields.some(f => f.key === '檢測數值') ? '檢測數值'
      : cat.fields.some(f => f.key === '監測數值') ? '監測數值' : null;
    const required = [
      ['日期(起)', '日期(起)'], ['日期(迄)', '日期(迄)'],
      ['時間(起)', '時間(起)'], ['時間(迄)', '時間(迄)'],
      [cat.itemField, '檢測項目'], [valueField, '監測數值'], [cat.unitField, '檢測單位'],
    ];
    const optional = [
      [cat.locationField, '採樣地點'], [cat.methodField, '檢測方法'],
      ['檢測極限', '檢測極限'], ['檢測機構許可證號', '檢測機構'], ['檢測類別', '檢測類別'],
    ];
    const filled = (key) => !!key && rows.some(r => String(r[key] ?? '').trim() !== '');
    return {
      found: required.filter(([k]) => filled(k)).map(([, label]) => label),
      missing: required.filter(([k]) => !filled(k)).map(([, label]) => label),
      extra: optional.filter(([k]) => filled(k)).map(([, label]) => label),
    };
  },

  /** Runs the detector over every sheet of a workbook, same shape as
   *  SmartParse.parseWorkbook's per-sheet loop. */
  parseWorkbook(category, sheetGrids) {
    const rows = [];
    const matchedSheets = [];
    const skippedSheets = [];
    for (const [sheetName, grid] of Object.entries(sheetGrids)) {
      const parsed = this.parseSheet(category, sheetName, grid);
      if (parsed && parsed.length) { rows.push(...parsed); matchedSheets.push(sheetName); }
      else skippedSheets.push(sheetName);
    }
    return { rows, matchedSheets, skippedSheets };
  },
};
