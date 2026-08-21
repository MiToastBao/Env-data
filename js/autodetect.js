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
    { key: 'dateRange', re: /^(採樣|監測|檢測|調查|檢驗|取樣|測定|量測)?(日期|期間)$/ },
    { key: 'timeRange', re: /^(採樣|監測|檢測|調查|測定|檢驗|量測)?時間$/ },
    // 測點編號 must be tested BEFORE the location rule, or a site CODE ends up in the
    // 採樣地點 field and the real place name is never looked for.
    { key: 'siteCode', re: /^測點編號$|^採樣點編號$|^測站編號$/ },
    { key: 'location', re: /^(採樣|監測|調查|檢測)?地點$|^名稱或地點$|^測點名稱$|^測站名稱$/ },
    { key: 'agency', re: /^(採樣|檢測|檢驗|受檢)單位$|^檢驗室名稱$|^公司名稱$|^檢測機構$/ },
    { key: 'method', re: /^(採樣|檢測|分析|檢驗)方法$/ },
    { key: 'category', re: /^樣品(特性|性質|種類)$|^檢測類別$|^監測類別$/ },
  ],

  cellStr(v) { return v === undefined || v === null ? '' : String(v).trim(); },

  /** Last column that actually holds content — shares SmartParse's cached scan so a
   *  sheet whose declared range is 16,000 columns wide isn't swept end to end by
   *  every heading probe. See SmartParse.lastCol for why these sheets are that wide. */
  lastCol(grid) { return SmartParse.lastCol(grid); },
  rowEnd(grid, row) { return Math.min((row || []).length, this.lastCol(grid) + 1); },

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
      const rowEnd = this.rowEnd(grid, row);
      for (let c = 0; c < rowEnd; c++) {
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
        const rowEnd = this.rowEnd(grid, row);
        let numericCells = 0;
        for (let c = 0; c < rowEnd; c++) {
          const v = this.cellStr(row[c]);
          if (v !== '' && /^-?[\d.,]+$/.test(v)) numericCells++;
        }
        if (rr > r && numericCells >= 2) break;
        for (let c = 0; c < rowEnd; c++) {
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

  /**
   * A row stating a REGULATORY LIMIT rather than a measurement.
   *
   * These sit directly under the readings and line up with exactly the same columns —
   * a 24-hour noise report puts 環境音量標準 76/75/72 right below the measured
   * 58.3/58.1/54.7 — so a reader that just keeps consuming numeric rows files the
   * legal limits as if they were this quarter's results.
   */
  // Anchored, so it matches a row LABELLED as a limit and not a monitoring point
  // whose name merely contains one of these words — 標準檢驗局旁民宅 and 基準點測站
  // are real station names, and an unanchored test cut the block short there,
  // discarding that row and every row below it.
  STANDARD_ROW_RE: /^(環境音量|噪音|振動|空氣品質|水質|放流水)?(管制)?(標準|基準|管制值|限|法規值?|建議值)(值)?$/,

  /** A row that summarises the rows above it rather than reporting a measurement. */
  SUMMARY_ROW_RE: /^(日?平均值?|均值|算術平均|合計|總計|小計|最大值?|最小值?|中位數)$/,

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
    const byItemColumn = this._parseItemColumnSheet(category, cat, grid);
    if (byItemColumn) return byItemColumn;
    // No "one row per test item" table here — try the layouts where the measurement
    // names run across the sheet as column headings instead (noise/vibration/air).
    return this.parseByColumnHeadings(category, grid);
  },

  _parseItemColumnSheet(category, cat, grid) {
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
    let metaDates = this.splitDateRange(meta.dateRange || meta.timeRange || '', fallbackYear);
    if (!metaDates.start) metaDates = this.recoverStitchedDate(grid);
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
      // Not a bare number? Pass it through unchanged and let the import preview ask
      // what to do with it — see renderLimitWarning in app.js.
      if (hasField('檢測極限')) {
        out['檢測極限'] = !limitApplies ? ''
          : (/^[\d.]+$/.test(limitRaw) ? SmartParse.formatNumber(limitRaw, 3) : limitRaw);
      }
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

  // ---------- layouts where the MEASUREMENT NAME is a column heading ----------

  /**
   * Third and fourth reading modes, used when no "one row per test item" table was
   * found. Between them they cover the way noise, vibration and air reports are
   * actually laid out — the measurement names run ACROSS the sheet as column
   * headings, with the numbers underneath:
   *
   *   MODE C (wide block)          Leq | Lmax | L5 | L10 | L50        <- names
   *                        14~15 | 67.2 | 80.9 | 73.2 | 71.1 | 63.2   <- values
   *                                   L日(7~20) | L晚(20~23) | L夜(23~翌日7)
   *                                       66.7 |       60.2 |    55.1
   *
   *   MODE D (label=value)   Lv日(Lv10)= | 39.19      Lv夜(Lv10)= | 30.0
   *
   * Each wide block found becomes its own selectable group in the import preview
   * (the same picker the air report's 日平均值／最大小時平均值 choice uses), because a
   * single sheet routinely holds several — an hour-by-hour block AND a summary block.
   * A block whose rows begin with a time range ("14~15") produces one row per hour,
   * with 時間(起)/(迄) filled from that range; otherwise the value row is taken as-is.
   */
  parseByColumnHeadings(category, grid) {
    const cat = CATEGORIES[category];
    if (!cat) return null;
    const meta = this.extractFormMeta(grid);
    let metaDates = this.splitDateRange(meta.dateRange || meta.timeRange || '', null);
    if (!metaDates.start) metaDates = this.recoverStitchedDate(grid);
    const metaTimes = this.splitTimeRange(meta.timeRange || meta.dateRange || '');
    const metaLocation = meta.location || '';
    const metaAgency = meta.agency ? SmartParse.reverseAgencyLookup(meta.agency) : '';
    const metaMethod = meta.method ? SmartParse.extractMethodCode(meta.method) : '';

    const blocks = this.findWideBlocks(grid).concat(this.findLabelValuePairs(grid));
    if (blocks.length === 0) return null;
    // Some sheets state the same figures twice — once across a row and once down a
    // column. Identical blocks are the same measurements, not two sets of them.
    const seenBlocks = new Set();
    const uniqueBlocks = blocks.filter(b => {
      const sig = b.entries.map(e => `${e.item}=${e.value}@${e.timeStart || ''}`).join('|');
      if (seenBlocks.has(sig)) return false;
      seenBlocks.add(sig);
      return true;
    });

    const valueFieldKey = cat.fields.some(f => f.key === '檢測數值') ? '檢測數值'
      : cat.fields.some(f => f.key === '監測數值') ? '監測數值' : null;
    const hasField = (k) => cat.fields.some(f => f.key === k);

    // WHICH BLOCK IS THE DEFAULT
    // A report states its summary figures (L日 / L晚 / L夜, Lv日(Lv10), a daily
    // average) alongside the hour-by-hour readings they were computed from. The
    // summary is what gets filed; the hourly block is 260-odd rows nobody normally
    // submits. So summary blocks are ticked by default and hourly blocks are not —
    // they are still read, and still there to tick, just folded away. When a sheet
    // has nothing but hourly blocks, the first one is the default so the import
    // never arrives completely empty.
    const hasSummary = uniqueBlocks.some(b => !b.hourly);
    const isPreferred = (block, idx) => (hasSummary ? !block.hourly : idx === 0);

    const out = [];
    uniqueBlocks.forEach((block, blockIdx) => {
      const preferred = isPreferred(block, blockIdx);
      block.entries.forEach(entry => {
        const { cmp, val, note } = SmartParse.parseValueCell(entry.value);
        if (val === '' && cmp === '' && note === '') return;
        const row = {};
        cat.fields.forEach(f => { row[f.key] = ''; });
        // A row's own date beats the sheet-level one; an hour-by-hour block that ran
        // past midnight advances the day (dayOffset) so the 01:00 reading is filed
        // under the day it was actually taken, not the day the run started. And an
        // hourly slot ends one hour later, NOT at the end of the whole 24-hour window
        // — every hourly row used to claim a 25-hour span.
        const baseDate = entry.rowDate || entry.dateStart || metaDates.start || '';
        const startDate = baseDate && entry.dayOffset
          ? SmartParse.addDaysISO(baseDate, entry.dayOffset) : baseDate;
        row['日期(起)'] = startDate;
        row['日期(迄)'] = entry.dateEnd
          || (entry.timeStart ? (entry.rollsOver && startDate ? SmartParse.addDaysISO(startDate, 1) : startDate)
            : (metaDates.end || startDate));
        row['時間(起)'] = entry.timeStart || metaTimes.start || '';
        row['時間(迄)'] = entry.timeEnd || metaTimes.end || row['時間(起)'];
        row[cat.locationField] = entry.rowLocation || metaLocation;
        row[cat.itemField] = entry.item;
        if (valueFieldKey) row[valueFieldKey] = /^[\d.]+$/.test(val) ? SmartParse.formatNumber(val, 3) : val;
        if (hasField('比較關係')) row['比較關係'] = cmp;
        if (cat.methodField) row[cat.methodField] = metaMethod;
        if (cat.unitField && entry.unitText) {
          const u = SmartParse.reverseUnitLookup(entry.unitText, entry.item);
          row[cat.unitField] = u.code;
          row._uncertainUnit = !!u.code && !u.confident;
        }
        row['檢測機構許可證號'] = metaAgency;
        row['備註'] = note || '';
        row._siteCode = meta.siteCode || '';
        row._rawLocation = entry.rowLocation || metaLocation;
        row._autoDetected = true;
        row._secondaryItem = !preferred;
        row._blockLabel = block.label;
        out.push(row);
      });
    });
    if (out.length === 0) return null;
    // A grid of numbers with NO sampling date and NO clock time anywhere is not a
    // monitoring result — it is a lookup table, a scratch sheet, or an instrument
    // list. Real workbooks are full of those (參照值, 工作表2, 氣象…), and this mode is
    // permissive enough to read numbers out of all of them. Requiring a date or a
    // time is what separates "measurements" from "numbers that happen to be here".
    const anchored = out.some(r => r['日期(起)'] || r['時間(起)']);
    if (!anchored) return null;
    const info = this.describeDetection(out, cat);
    out.forEach(r => { r._autoDetectInfo = info; });
    return out;
  },

  /** A cell that could be a measurement name: short, has content, isn't a number,
   *  isn't a "label：value" form field, and isn't one of the standard field headings
   *  (those are the item-table mode's job, not this one). */
  _looksLikeMeasurementName(v) {
    const s = this.cellStr(v).replace(/\s+/g, ' ');
    if (s === '' || s.length > 24) return false;
    if (/[:：]/.test(s)) return false;
    if (/^-?[\d.,%]+$/.test(s)) return false;
    // A bare time window is a WHEN, not a WHAT: "14~15", "(7～20)", "(23～翌日7)".
    // Reports print these as sub-headings under a real measurement name (日間 L日,
    // 晚間 L晚…), and treating them as items produces meaningless "(7～20)" rows that
    // duplicate readings already read properly from the report sheets themselves.
    if (/^[(（]?\s*\d{1,2}\s*[~～至\-—]\s*(?:翌日)?\s*\d{1,2}\s*[)）]?$/.test(s)) return false;
    if (!/[A-Za-z0-9\u4e00-\u9fff]/.test(s)) return false;      // "*", "—", "※" are placeholders
    if (this._isNumericCell(s)) return false;                    // "< 0.002" is a reading, not a name
    // A standard field heading (單位, 監測項目, 備註…) is a column ABOUT the
    // measurements, never a measurement in its own right.
    if (this.matchHeader(s)) return false;
    return true;
  },

  /** A row that could be a band of column headings. Any colon anywhere in the row
   *  means it is part of the report's "label：value" form area, not a table header —
   *  that single test removes essentially every false positive from the top of a
   *  report (客戶名稱, 業別, 測點編號 and friends all sit on such rows). */
  _isLabelRow(row) {
    if (!row) return false;
    let names = 0;
    for (const cell of row) {
      const s = this.cellStr(cell);
      if (/[:：]/.test(s)) return false;
      if (this._looksLikeMeasurementName(cell)) names++;
    }
    return names >= 2;
  },

  _isNumericCell(v) {
    const s = this.cellStr(v);
    if (s === '') return false;
    return /^[<>]?\s*-?[\d.,]+$/.test(s) || /^ND$/i.test(s);
  },

  /**
   * Finds every "row of names, numbers underneath" block on a sheet.
   *
   * Candidates are collected for EVERY row first and only then chosen, largest
   * first, discarding any that overlaps one already taken. That ordering matters: a
   * report's table头 is usually two or three stacked rows (a merged banner such as
   * 「時間 / 噪 / 音 dB(A)」 sitting above the real 「Leq | Lmax | L5 | L10」 row), and
   * taking the topmost candidate would lock onto the banner and read three garbled
   * columns instead of seven real ones. Whichever row explains more of the numbers
   * below it is the real heading.
   */
  findWideBlocks(grid) {
    const candidates = [];
    const limit = Math.min(grid.length, 200);
    for (let r = 0; r < limit; r++) {
      const labelRow = grid[r] || [];
      if (!this._isLabelRow(labelRow)) continue;
      const labels = [];
      labelRow.forEach((cell, c) => {
        if (this._looksLikeMeasurementName(cell)) labels.push({ col: c, name: this.cellStr(cell).replace(/\s+/g, ' ') });
      });
      if (labels.length < 2) continue;
      // A row of nine identical "ppm" cells is the UNIT row sitting under the real
      // pollutant headings, not a set of measurement names. Distinct names are what
      // make a heading row a heading row.
      const distinct = new Set(labels.map(l => l.name)).size;
      if (distinct < 2) continue;

      // the value row: within the next 3 rows, one where most of those columns are
      // numbers (3 rather than 1, because a unit row and a sub-heading row commonly
      // sit between the headings and the first line of data)
      let valueRowIdx = -1;
      for (let rr = r + 1; rr <= Math.min(r + 3, grid.length - 1); rr++) {
        const hits = labels.filter(l => this._isNumericCell((grid[rr] || [])[l.col])).length;
        if (hits >= 2 && hits >= Math.ceil(labels.length * 0.5)) { valueRowIdx = rr; break; }
      }
      if (valueRowIdx < 0) continue;

      // Does the block continue downward as one row per time slot? ("14~15", "15~16")
      const timeOf = (rowIdx) => {
        const row = grid[rowIdx] || [];
        for (let c = 0; c <= Math.max(0, labels[0].col); c++) {
          const m = this.cellStr(row[c]).match(/^(\d{1,2})\s*[~～-]\s*(\d{1,2})$/);
          if (m) return { h1: parseInt(m[1], 10), h2: parseInt(m[2], 10) };
        }
        return null;
      };

      // Anything to the LEFT of the first measurement column that looks like a place
      // name or a date belongs to that row, not to the sheet as a whole — a
      // "one row per site" summary table carries both.
      const rowContext = (rowIdx) => {
        const row = grid[rowIdx] || [];
        let loc = '', date = '', isStandard = false;
        for (let c = 0; c < labels[0].col; c++) {
          const v = this.cellStr(row[c]);
          if (v === '' || this._isNumericCell(v)) continue;
          if (this.STANDARD_ROW_RE.test(v) || this.SUMMARY_ROW_RE.test(v)) { isStandard = true; continue; }
          const iso = DateTimeUtil.toISODate(v);
          if (/^\d{4}-\d{2}-\d{2}$/.test(iso)) { if (!date) date = iso; continue; }
          if (!loc && !/^\d{1,2}\s*[~～-]\s*\d{1,2}$/.test(v)) loc = v;
        }
        return { loc, date, isStandard };
      };

      const entries = [];
      const pad = n => String(n).padStart(2, '0');
      let lastRow = valueRowIdx;
      let anyTime = false;
      let dayOffset = 0;
      let prevH1 = null;
      let carriedLoc = '', carriedDate = '';
      for (let rr = valueRowIdx; rr < grid.length; rr++) {
        const hits = labels.filter(l => this._isNumericCell((grid[rr] || [])[l.col])).length;
        // Stop only when the row carries no measurement at all. Breaking on "this row
        // has no time range" meant a summary table with one row per station gave up
        // after its FIRST station — 25 of 30 real readings silently discarded.
        if (hits === 0) break;
        const t = timeOf(rr);
        if (t) {
          anyTime = true;
          // an hour-by-hour block that wraps past midnight continues on the next day
          if (prevH1 !== null && t.h1 < prevH1) dayOffset++;
          prevH1 = t.h1;
        }
        const ctx = rowContext(rr);
        // "環境音量標準 76 | 75 | 72" sits immediately under the measured values and
        // occupies the same columns — reading on would file the legal limits as this
        // quarter's readings.
        if (ctx.isStandard) break;
        // A merged cell reports its text on the FIRST row of the merge only, so a
        // station sampled on two dates shows its name once and blank underneath.
        // Carry the last seen values down, which is what the merge means visually.
        if (ctx.loc) carriedLoc = ctx.loc; else ctx.loc = carriedLoc;
        if (ctx.date) carriedDate = ctx.date; else ctx.date = carriedDate;
        labels.forEach(l => {
          const v = this.cellStr((grid[rr] || [])[l.col]);
          // Only actual readings become rows. A caption sitting above the numbers
          // ("振動測值" over "(dB)", "測定項目" over "合成噪音測值(L1)") looks like a
          // heading with a value underneath but carries no measurement at all.
          if (!this._isNumericCell(v)) return;
          entries.push({
            item: l.name, value: v,
            timeStart: t ? `${pad(t.h1)}:00:00` : '',
            timeEnd: t ? `${pad(t.h2 % 24)}:00:00` : '',
            dayOffset,
            // an hourly slot written "23~24" or "23~00" ends on the following day
            rollsOver: !!t && (t.h2 <= t.h1 || t.h2 >= 24),
            rowLocation: ctx.loc,
            rowDate: ctx.date,
          });
        });
        lastRow = rr;
      }
      if (entries.length === 0) continue;
      const preview = labels.slice(0, 4).map(l => l.name).join('、') + (labels.length > 4 ? '…' : '');
      candidates.push({
        label: `${anyTime ? '逐時' : '彙整'}數值：${preview}`,
        hourly: anyTime, entries, startRow: r, endRow: lastRow, distinct,
      });
    }

    // Largest wins; anything overlapping an accepted block is a stacked-header
    // duplicate of it and is dropped.
    // Ranked by how many DISTINCT measurement names the row explains, before sheer
    // row count: a unit row ("ppm ppm ppm …") lines up with just as many numbers as
    // the pollutant-name row above it, but names nothing.
    candidates.sort((a, b) => b.distinct - a.distinct
      || b.entries.length - a.entries.length || a.startRow - b.startRow);
    const taken = [];
    for (const cand of candidates) {
      if (taken.some(t => cand.startRow <= t.endRow && cand.endRow >= t.startRow)) continue;
      taken.push(cand);
      if (taken.length >= 8) break; // a sheet with more blocks than this isn't a report
    }
    return taken.sort((a, b) => a.startRow - b.startRow);
  },

  /** Finds scattered "NAME= value" measurement pairs, e.g. "Lv日(Lv10)=" then 39.19. */
  findLabelValuePairs(grid) {
    const entries = [];
    for (let r = 0; r < Math.min(grid.length, 120); r++) {
      const row = grid[r] || [];
      const rowEnd = this.rowEnd(grid, row);
      for (let c = 0; c < rowEnd; c++) {
        const s = this.cellStr(row[c]).replace(/\s+/g, '');
        const m = s.match(/^(.{1,20}?)[=＝]$/);
        if (!m) continue;
        const name = m[1];
        if (!/[A-Za-z一-鿿]/.test(name)) continue;
        for (let cc = c + 1; cc < row.length; cc++) {
          const v = this.cellStr(row[cc]);
          if (v === '') continue;
          if (this._isNumericCell(v)) entries.push({ item: name, value: v });
          break;
        }
      }
    }
    if (entries.length < 2) return [];
    const preview = entries.slice(0, 4).map(e => e.item).join('、') + (entries.length > 4 ? '…' : '');
    return [{ label: `標示式數值：${preview}`, hourly: false, entries }];
  },

  /**
   * Last resort for a sampling date: some reports never write it as a label at all,
   * they run it DOWN a column one merged cell per character —
   * "115 / 年 / 6 / 月 / 25 / 日 / 至 / 115 / 年 / 6 / 月 / 26 / 日". Stitching the
   * left-hand columns back together recovers it. Only called when nothing else found
   * a date, so a stray number elsewhere can't hijack a date that was read properly.
   */
  recoverStitchedDate(grid) {
    for (const col of [0, 1]) {
      const all = grid.map(r => this.cellStr((r || [])[col])).filter(v => v !== '');
      if (all.length === 0) continue;
      // A column that already holds WHOLE dates must never be joined: concatenating a
      // list of "115/07/03" values invents digit runs that were never there and
      // produced the year 3115. If any single cell parses to a complete date on its
      // own, this is a list, not fragments — leave it to the per-row readers.
      if (all.some(v => /^\d{4}-\d{2}-\d{2}$/.test(DateTimeUtil.toISODate(v)))) continue;
      // Join only the SHORT cells: a stitched date is one fragment per merged cell
      // ("115" / "年" / "6" / "月" / "25" / "日" / "至" …). Longer cells in the same
      // column are row captions and prose, and only add noise.
      const cells = all.filter(v => v.length <= 6);
      if (cells.length === 0) continue;
      const joined = cells.join('');
      if (!/\d/.test(joined)) continue;
      const parts = joined.split(/至|~|～|－|—/);
      const a = DateTimeUtil.toISODate(parts[0] || '');
      if (!/^\d{4}-\d{2}-\d{2}$/.test(a)) continue;
      const b = parts.length > 1 ? DateTimeUtil.toISODate(parts[1]) : '';
      const y = parseInt(a.slice(0, 4), 10);
      if (y < 1990 || y > 2100) continue;
      return { start: a, end: /^\d{4}-\d{2}-\d{2}$/.test(b) ? b : a };
    }
    return { start: '', end: '' };
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
