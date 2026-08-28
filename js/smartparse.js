// smartparse.js
// Extracts monitoring rows directly from real lab-report Excel forms
// (as opposed to clean tabular exports). These reports are semi-structured
// forms with labels scattered around values, not header+rows tables, so
// ordinary column-mapping doesn't apply. This module was built and
// verified against an actual quarterly report package (labels/positions
// come from 華光工程顧問股份有限公司-style report templates, a common
// consultant format for Taiwan environmental monitoring). Other labs'
// report layouts will likely need their own patterns added here — when a
// sheet isn't recognized, the app falls back to the generic column-mapping
// importer so nothing is silently lost.

const SmartParse = {

  CJK_NUM: { 一: 1, 二: 2, 三: 3, 四: 4, 五: 5, 六: 6, 七: 7, 八: 8, 九: 9, 十: 10 },

  // Fields that a "site profile" (per report site-code) can override/fill,
  // shown as editable columns in the smart-import preview UI.
  SITE_PROFILE_FIELDS: {
    noise: [
      { key: '監測地點', label: '正式監測站名', type: 'text' },
      { key: '檢測類別', label: '檢測類別（系統判讀，可調整）', type: 'select', options: ['', ...CATEGORY_TYPE_OPTIONS.noise] },
      { key: '座標系統', label: '座標系統', type: 'select', options: ['', '2', '3'],
        optionLabels: { '2': '2：WGS84（全球座標）', '3': '3：TWD97-TM2（投影座標系）' } },
      { key: '採樣座標-經度 X', label: '座標X', type: 'text' },
      { key: '採樣座標-緯度 Y', label: '座標Y', type: 'text' },
      { key: '管制區', label: '管制區', type: 'text' },
      { key: '環境音量標準', label: '環境音量標準(0/1/2)', type: 'text' },
    ],
    water: [
      { key: '採樣地點', label: '正式採樣地點', type: 'text' },
      { key: '檢測類別', label: '檢測類別（系統判讀，可調整）', type: 'select', options: ['', ...CATEGORY_TYPE_OPTIONS.water] },
      { key: '座標系統', label: '座標系統', type: 'select', options: ['', '2', '3'],
        optionLabels: { '2': '2：WGS84（全球座標）', '3': '3：TWD97-TM2（投影座標系）' } },
      { key: '採樣座標-經度 X', label: '座標X', type: 'text' },
      { key: '採樣座標-緯度 Y', label: '座標Y', type: 'text' },
      { key: '管制編號', label: '管制編號', type: 'text' },
    ],
    air: [
      { key: '採樣地點', label: '正式採樣地點', type: 'text' },
      { key: '檢測類別', label: '檢測類別（系統判讀，可調整）', type: 'select', options: ['', ...CATEGORY_TYPE_OPTIONS.air] },
      { key: '座標系統', label: '座標系統', type: 'select', options: ['', '2', '3'],
        optionLabels: { '2': '2：WGS84（全球座標）', '3': '3：TWD97-TM2（投影座標系）' } },
      { key: '採樣座標-經度 X', label: '座標X', type: 'text' },
      { key: '採樣座標-緯度 Y', label: '座標Y', type: 'text' },
      { key: '管制編號', label: '管制編號', type: 'text' },
    ],
    // 地質 reports (底泥/土壤 lab sheets) never print coordinates or a 檢測類別 code,
    // so both are asked for here once per site. 檢測類別 deliberately starts BLANK
    // rather than being guessed from 樣品特性 — 底泥品質 vs 土壤品質 vs 廢棄物 is a
    // filing decision, and a wrong auto-filled category is harder to notice than an
    // obviously empty one.
    geo: [
      { key: '採樣地點', label: '正式採樣地點', type: 'text' },
      { key: '檢測類別', label: '檢測類別（請自行選擇）', type: 'select', options: ['', ...CATEGORY_TYPE_OPTIONS.geo] },
      { key: '座標系統', label: '座標系統', type: 'select', options: ['', '2', '3'],
        optionLabels: { '2': '2：WGS84（全球座標）', '3': '3：TWD97-TM2（投影座標系）' } },
      { key: '採樣座標-經度 X', label: '座標X', type: 'text' },
      { key: '採樣座標-緯度 Y', label: '座標Y', type: 'text' },
      { key: '採樣深度(公尺)', label: '採樣深度(公尺)', type: 'text' },
      { key: '管制編號', label: '管制編號', type: 'text' },
    ],
  },

  // ---------- grid utilities ----------
  // grid: array of rows, each row an array of cell values (strings/numbers/''),
  // as produced by XLSX.utils.sheet_to_json(ws, {header:1, defval:''}).

  cellStr(v) { return v === undefined || v === null ? '' : String(v).trim(); },

  /**
   * The last column index that actually holds anything, cached on the grid.
   *
   * SheetJS reports these report sheets with an enormous used range — 06525NV25's
   * N-01 comes back 97 rows x 16,163 columns (1.57M cells), almost all empty
   * padding. Every findCell was sweeping all of it, which made one workbook take
   * 7.2 seconds of blocking work and froze the browser tab during import. Every
   * scan below stops at this column instead.
   */
  lastCol(grid) {
    if (grid.__lastCol !== undefined) return grid.__lastCol;
    let last = -1;
    for (let r = 0; r < grid.length; r++) {
      const row = grid[r] || [];
      for (let c = row.length - 1; c > last; c--) {
        if (this.cellStr(row[c]) !== '') { last = c; break; }
      }
    }
    try { Object.defineProperty(grid, '__lastCol', { value: last, enumerable: false }); } catch (e) { /* frozen grid */ }
    return last;
  },

  /** Scan the whole grid for the first cell matching regex; return {r,c} or null. */
  findCell(grid, regex) {
    const maxCol = this.lastCol(grid);
    for (let r = 0; r < grid.length; r++) {
      const row = grid[r];
      if (!row) continue;
      const end = Math.min(row.length, maxCol + 1);
      for (let c = 0; c < end; c++) {
        if (regex.test(this.cellStr(row[c]))) return { r, c };
      }
    }
    return null;
  },

  /** From a found label cell, return the next N non-empty cell values to its right (same row). */
  valuesRightOf(grid, r, c, n = 1) {
    const row = grid[r] || [];
    const out = [];
    for (let cc = c + 1; cc < row.length && out.length < n; cc++) {
      const v = this.cellStr(row[cc]);
      if (v !== '') out.push(v);
    }
    return out;
  },

  /**
   * Find a label anywhere in the grid and return its value.
   *
   * Two layouts have to be handled, and BOTH occur inside the same workbook:
   *   "採樣時間：" in one cell, "115年06月25日12時00分" in the next cell to the right, and
   *   "3.量測方法依據：NIEA P204.90C" all inside ONE cell.
   * The same-cell form is checked FIRST. Looking rightwards first was a real bug: on a
   * 振動 report the note cell "3.量測方法依據：NIEA P204.90C" matched, the value after
   * the colon was ignored, and the next non-empty cell far to the right — a leftover
   * project title — was returned as the 檢測方法 instead.
   */
  labelValue(grid, regex, n = 1) {
    const hit = this.findCell(grid, regex);
    if (!hit) return n === 1 ? null : [];
    if (n === 1) {
      const cell = this.cellStr(grid[hit.r][hit.c]);
      const colonIdx = cell.search(/[:：]/);
      if (colonIdx >= 0) {
        const inline = cell.slice(colonIdx + 1).trim();
        if (inline !== '') return inline;
      }
    }
    const vals = this.valuesRightOf(grid, hit.r, hit.c, n);
    return n === 1 ? (vals[0] ?? null) : vals;
  },

  // ---------- small converters ----------

  /** "115.01.30" / "115年02月05日" / "115.03.09(平日)" -> "YYYY-MM-DD" (ROC year + 1911) */
  rocDateToISO(text) {
    if (!text) return '';
    const s = String(text);
    let m = s.match(/(\d{2,3})[.\-年](\d{1,2})[.\-月](\d{1,2})/);
    if (!m) return '';
    const y = parseInt(m[1], 10) + 1911;
    const mo = String(m[2]).padStart(2, '0');
    const d = String(m[3]).padStart(2, '0');
    return `${y}-${mo}-${d}`;
  },
  addDaysISO(iso, days) {
    if (!iso) return '';
    const d = new Date(iso + 'T00:00:00');
    d.setDate(d.getDate() + days);
    const pad = n => String(n).padStart(2, '0');
    return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  },
  /** "10:05~10:07" -> ["10:05:00","10:07:00"] */
  splitTimeRange(text) {
    if (!text) return ['', ''];
    const m = String(text).match(/(\d{1,2}:\d{2})\s*[~\-至]\s*(\d{1,2}:\d{2})/);
    if (!m) return ['', ''];
    return [m[1] + ':00', m[2] + ':00'];
  },
  hourToTod(hh) {
    const h = parseInt(hh, 10);
    if (h >= 6 && h < 18) return '日間';
    if (h >= 18 && h < 22) return '晚間';
    return '夜間';
  },
  /** Extract a bare NIEA-style method code, dropping the version suffix:
   *  "NIEA W217.51A" -> "NIEA W217". Also recognizes CNS/EPA/ASTM/APHA/ISO codes.
   *
   *  When the text is clearly NOT a method — a sentence, a company name, a project
   *  title — this returns '' rather than the text itself. A blank 檢測方法 is visible,
   *  gets flagged by the import preview and can be filled from memory; a plausible-
   *  looking wrong one goes straight into an official filing unnoticed. Short
   *  non-coded method names ("電極法", "碘定量法") are still kept. */
  extractMethodCode(text) {
    if (!text) return '';
    const s = String(text).trim();
    const m = s.match(/NIEA\s*([A-Z]?\d+)/i);
    if (m) return `NIEA ${m[1]}`;
    const other = s.match(/\b(?:CNS|EPA|ASTM|APHA|ISO)\s*[A-Z]?\d[\w.-]*/i);
    if (other) return other[0].replace(/\s+/g, ' ').trim();
    return s.length > 10 ? '' : s;
  },
  /** "噪音管制區第二類" / "第2類" -> "第2類" */
  extractZone(text) {
    if (!text) return '';
    const m = String(text).match(/第([一二三四五六七八九十\d]+)類/);
    if (!m) return '';
    const raw = m[1];
    const num = /^\d+$/.test(raw) ? raw : String(this.CJK_NUM[raw] || '');
    return num ? `第${num}類` : '';
  },

  /** Reverse-lookup a testing-agency code from a free-text company name. */
  reverseAgencyLookup(text) {
    if (!text) return '';
    const clean = String(text).split('(')[0].split('（')[0].trim();
    if (!clean) return '';
    const entries = Object.entries(AGENCY_CODES).filter(([code]) => code !== 'AA');
    const exact = entries.find(([, name]) => name === clean);
    if (exact) return exact[0];
    // Longest match wins. Iterating in key order returned the PARENT company for a
    // branch office — "台灣檢驗科技股份有限公司高雄分公司" resolved to code 35
    // (the parent) instead of 105, because the parent's shorter name is a prefix and
    // its numeric key comes first. A wrong 檢測機構許可證號 on an official filing.
    const hits = entries
      .filter(([, name]) => clean.includes(name) || name.includes(clean))
      .sort((a, b) => String(b[1]).length - String(a[1]).length);
    return hits.length ? hits[0][0] : '';
  },

  UNIT_ALIASES: {
    '度': '161', 'pH': '161',
    'μg/m3': '127', 'ug/m3': '127', 'μg/m^3': '127', 'ug/m^3': '127',
    'μg/L': '126', 'ug/L': '126',
    'mg/L': '47', 'MG/L': '47',
  },
  // Some lab reports print a chemically-precise unit (e.g. "mg P/L" for phosphorus-as-P)
  // that is a real, valid code in the table, but doesn't match this filing's convention
  // of using the plain "mg/L" code for that item. Confirmed reporting-convention
  // override, not a guess — applied regardless of what unit text the report shows.
  ITEM_UNIT_OVERRIDES: {
    '總磷': '47',
    // The report prints "mg SO42-/L" for sulphate. Confirmed by the person filing
    // these as a clerical slip in the lab's own template — the correct unit for this
    // filing is plain mg/L — so it is fixed outright here rather than being flagged
    // for review every quarter.
    '硫酸鹽': '47',
  },
  /** Reverse-lookup a unit code from a free-text unit symbol (as printed on a lab report).
   * Returns { code, confident }. confident=false means either no exact match was found
   * (an alias/fallback was used or nothing matched at all) or more than one official code
   * shares the same printed symbol (e.g. "ug/m^3" and "μg/m3" are both valid officially-listed
   * unit codes for the same physical unit) — these cases should be flagged for human review
   * rather than silently trusted, since an environmental filing with the wrong unit code is a
   * meaningful error. */
  reverseUnitLookup(text, itemName) {
    if (itemName && this.ITEM_UNIT_OVERRIDES[itemName]) {
      return { code: this.ITEM_UNIT_OVERRIDES[itemName], confident: true };
    }
    if (!text) return { code: '', confident: false };
    const clean = String(text).trim();
    const exactMatches = Object.entries(UNIT_CODES).filter(([, name]) => name === clean);
    if (exactMatches.length === 1) return { code: exactMatches[0][0], confident: true };
    if (exactMatches.length > 1) return { code: exactMatches[0][0], confident: false }; // ambiguous: multiple official codes print identically

    if (this.UNIT_ALIASES[clean]) return { code: this.UNIT_ALIASES[clean], confident: false };

    // Chemically-qualified units — "mg SO42-/L" for sulphate-as-SO4, "mg P/L",
    // "mg N/L", "mg CaCO3/L" — are the same physical unit as the plain form, but the
    // official code table only lists the plain one, so an exact match never happens
    // and the unit came out BLANK. Fall back to the plain unit and mark it
    // unconfident, so it is filled in but still listed for checking.
    const species = clean.match(/^(mg|μg|ug|g|ng|kg)\s+\S+\s*\/\s*(L|mL|m3|m\^3|kg|g)$/i);
    if (species) {
      const base = `${species[1]}/${species[2]}`.toLowerCase();
      const baseMatch = Object.entries(UNIT_CODES).filter(([, name]) => String(name).toLowerCase() === base);
      if (baseMatch.length >= 1) return { code: baseMatch[0][0], confident: false };
    }

    // loose fallback: case-insensitive
    const lower = clean.toLowerCase();
    const looseMatches = Object.entries(UNIT_CODES).filter(([, name]) => String(name).toLowerCase() === lower);
    if (looseMatches.length >= 1) return { code: looseMatches[0][0], confident: false };

    return { code: '', confident: false };
  },

  /** 樣品特性 free text -> 檢測類別 enum used by the water template. */
  sampleTypeToCategory(text) {
    const s = String(text || '');
    if (/放流|廢水/.test(s)) return '污廢水';
    if (/河川|溪|圳|排水/.test(s)) return '河川';
    if (/地下水/.test(s)) return '地下水';
    if (/海域|海水|海灘/.test(s)) return '海域海灘';
    if (/水庫/.test(s)) return '水庫';
    if (/自來水/.test(s)) return '自來水';
    if (/回收水/.test(s)) return '回收水';
    if (/雨水/.test(s)) return '雨水';
    if (/飲用水/.test(s)) return '飲用水';
    return '';
  },

  /** 樣品特性 free text -> the 檢測類別 enum used by the 地質 template. Kept separate
   *  from sampleTypeToCategory because "底泥"/"土壤" must never be classed as water
   *  and vice versa — this is also what keeps the water and geo sheet parsers from
   *  claiming each other's reports during batch auto-detection, since both report
   *  types share the same 檢驗項目 table layout. */
  sedimentTypeToCategory(text) {
    const s = String(text || '');
    if (/底泥|底質|沉積物/.test(s)) return '底泥品質';
    if (/土壤/.test(s)) return '土壤品質';
    if (/廢棄物|灰渣|污泥/.test(s)) return '廢棄物';
    if (/毒化|毒性化學/.test(s)) return '毒化物質';
    if (/土砂|沖蝕|淤積/.test(s)) return '土砂觀測';
    return '';
  },

  /** True when a 樣品特性 clearly describes a solid/sediment sample rather than water. */
  isGeoSampleType(text) { return !!this.sedimentTypeToCategory(text); },
  /** True when a 樣品特性 clearly describes a water sample. */
  isWaterSampleType(text) { return !!this.sampleTypeToCategory(text); },

  /** Normalize a raw test-item name to the conventional short form used in filings. */
  normalizeItemName(text) {
    return String(text || '').replace(/[（(].*?[）)]/g, '').trim();
  },

  /**
   * Trim floating-point noise and unnecessary trailing zeros ("4.0" -> "4",
   * "36.4194255433345" -> "36.4") WITHOUT ever coarsening a value the report
   * actually printed.
   *
   * The previous version rounded to a fixed number of decimals, which quietly
   * destroyed trace results: mercury reported as 0.0006 mg/L was filed as 0.001
   * (+67%), and a mercury detection limit of 0.0004 was filed as 0. Both appear in
   * real 水質 reports. The rule now is: round only to strip float noise, and never
   * to fewer decimals than the source itself showed.
   */
  formatNumber(v, decimals = 1) {
    const raw = String(v ?? '').trim();
    const n = parseFloat(raw);
    if (isNaN(n)) return raw;
    if (n === 0) return '0';
    // A value carrying more than 10 significant digits is not a reported
    // measurement — it is a spreadsheet-computed average whose full float sits in
    // the cell (36.4194255433345, 0.30000000000000004). Those get the caller's
    // requested precision. Anything a lab actually PRINTED keeps every digit it
    // printed, which is what protects 0.0004 and 0.0006.
    const significantDigits = String(Math.abs(n)).replace(/[^0-9]/g, '').replace(/^0+/, '').length;
    if (significantDigits > 10) return String(Number(n.toFixed(decimals)));
    const printedDecimals = (raw.split('.')[1] || '').replace(/[^0-9].*$/, '').length;
    return String(Number(n.toFixed(Math.min(Math.max(decimals, printedDecimals), 12))));
  },

  /** Parse a raw "檢測值" cell into {比較關係, 檢測數值} pairs, handling ND / </> / scientific-ish notation. */
  parseValueCell(raw) {
    // Real reports mix full-width and half-width forms freely — the same
    // consultant's own boilerplate writes "以ＮＤ表示". A full-width ＮＤ used to
    // fall through as literal text, which left 比較關係 empty, put non-numeric text
    // in 檢測數值, AND suppressed the detection limit (limitApplies never became
    // true). Normalize once, up front.
    const s = String(raw ?? '').trim()
      .replace(/[＜﹤〈]/g, '<').replace(/[＞﹥〉]/g, '>')
      .replace(/[Ｎｎ]/g, 'N').replace(/[Ｄｄ]/g, 'D').replace(/[Ａａ]/g, 'A')
      .replace(/／/g, '/').replace(/（/g, '(').replace(/）/g, ')')
      .replace(/[０-９]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0))
      .replace(/[．]/g, '.')
      .trim();
    if (s === '') return { cmp: '', val: '', note: '' };
    // ND: 比較關係 filled, 檢測數值 stays BLANK — "ND" is not itself a measured
    // value, so it has no business sitting in the value field once 比較關係
    // already records it. Detection limit (if the report has one) still applies
    // here — see call sites below.
    if (/^N\.?D\.?$/i.test(s)) return { cmp: 'ND', val: '', note: '' };
    // NA / 未檢測: neither 比較關係 nor 檢測數值 should carry "未檢測" text — that
    // belongs in 備註 instead (returned via `note`), since 比較關係 is meant to
    // hold only the ND/</> symbols, not free-text explanations. Call sites are
    // responsible for merging `note` into the row's 備註 field.
    if (/^(NA|未檢測|N\.A\.?)$/i.test(s)) return { cmp: '', val: '', note: '未檢測' };
    // "<10.0(6.9)" style: the instrument actually read 6.9, but that's below the
    // report's required first calibration-curve point (a different threshold from
    // the plain detection limit), so the filing must show "<10.0" — the parenthetical
    // raw reading is informational only and is intentionally dropped, not stored.
    let m = s.match(/^([<>])\s*([\d.]+)\s*\(\s*[\d.]+\s*\)$/);
    if (m) return { cmp: m[1], val: m[2], note: '' };
    m = s.match(/^([<>])\s*([\d.]+)$/);
    if (m) return { cmp: m[1], val: m[2], note: '' };
    // e.g. "6.0×104" meaning 6.0×10^4 (single trailing exponent digit)
    m = s.match(/^([\d.]+)\s*[×xX]\s*10\^?(\d)$/);
    if (m) {
      const val = parseFloat(m[1]) * Math.pow(10, parseInt(m[2], 10));
      return { cmp: '', val: String(val), note: '' };
    }
    return { cmp: '', val: s, note: '' };
  },

  // ---------- report-type detectors & parsers ----------

  /** BN (固定音源噪音-營建工程) / BV (營建工程振動) / LFN (低頻噪音) single-event report.
   *  One sheet -> 1-2 rows. */
  parseNoiseEventSheet(grid, sheetName = '') {
    const title = this.cellStr(grid[0]?.[0]) + ' ' + this.cellStr(grid[1]?.[0]) + ' ' + this.cellStr(grid[1]?.[1]);
    const sampleChar = this.labelValue(grid, /樣品特性[:：]/) || '';
    const isLFN = /低頻噪音/.test(title) || /低頻噪音/.test(sampleChar);
    // A construction VIBRATION sheet sits in the same workbook as the construction
    // NOISE sheets and shares its wording ("固定音源噪音振動測定報告(營建工程)"), so it
    // has to be told apart by what it actually measures: a 振動測值 table of
    // Lveq/Lvmax/Lv5… rather than a 整體營建噪音值 figure. Without this it matched the
    // noise branch, found no noise value, and the whole sheet was silently dropped.
    const vibHit = this.findCell(grid, /^振動測值$/);
    const isBV = !isLFN && (!!vibHit || /^BV/i.test(sheetName));
    const isBN = !isLFN && !isBV && (/固定音源噪音|營建工程/.test(title) || /固定音源噪音|營建工程/.test(sampleChar));
    if (!isLFN && !isBN && !isBV) return null;

    // 監測日期 on the noise sheets, 測定日期 on the vibration ones — same field.
    const dateRaw = this.labelValue(grid, /監測日期[:：]/) || this.labelValue(grid, /測定日期[:：]/)
      || this.labelValue(grid, /檢測日期[:：]/);
    const dateISO = this.rocDateToISO(dateRaw);
    const timeRaw = this.labelValue(grid, /測定時間[:：]/) || this.labelValue(grid, /監測時間[:：]/);
    const [tStart, tEnd] = this.splitTimeRange(timeRaw);
    // A night-time construction measurement written "23:50~00:10" ends on the NEXT
    // day. Filing 日期(迄) = 日期(起) gave a window that ends 23h40m before it starts.
    const dateEndISO = (tStart && tEnd && tEnd < tStart) ? this.addDaysISO(dateISO, 1) : dateISO;
    // 監測地點 on the noise sheets, 名稱或地點 on the vibration ones.
    const location = this.labelValue(grid, /監測地點[:：]/) || this.labelValue(grid, /名稱或地點[:：]/) || '';
    const coordHit = this.findCell(grid, /大地座標/);
    let coordX = '', coordY = '';
    if (coordHit) {
      const vals = this.valuesRightOf(grid, coordHit.r, coordHit.c, 2);
      coordX = vals[0] || ''; coordY = vals[1] || '';
    }
    const methodRaw = this.labelValue(grid, /採樣方法[:：]/);
    const method = this.extractMethodCode(methodRaw);
    const agencyRaw = this.labelValue(grid, /採樣單位[:：]/);
    const agencyCode = this.reverseAgencyLookup(agencyRaw);
    // 管制區: scan ALL 噪音管制區 mentions, not just the first. On the LFN template
    // the first hit is a bare caption whose value ("二") sits in the neighbouring
    // cell, so stopping at the first match filed every 低頻噪音 row with a blank
    // 管制區 even though the report plainly states 第二類.
    const zone = (() => {
      const maxCol = this.lastCol(grid);
      for (let r = 0; r < grid.length; r++) {
        const row = grid[r] || [];
        const end = Math.min(row.length, maxCol + 1);
        for (let c = 0; c < end; c++) {
          const cell = this.cellStr(row[c]);
          if (!/噪音管制區/.test(cell)) continue;
          const inCell = this.extractZone(cell);
          if (inCell) return inCell;
          const nextVal = this.valuesRightOf(grid, r, c, 1)[0] || '';
          const beside = this.extractZone(nextVal) || this.extractZone(`第${nextVal}類`);
          if (beside) return beside;
        }
      }
      return '';
    })();
    const tod = tStart ? this.hourToTod(tStart.split(':')[0]) : '日間';
    const siteCode = this.labelValue(grid, /測點編號[:：]/) || '';

    const baseRow = {
      '日期(起)': dateISO, '時間(起)': tStart, '日期(迄)': dateEndISO, '時間(迄)': tEnd,
      '監測地點': location, '座標系統': coordX ? '3' : '', '採樣座標-經度 X': coordX, '採樣座標-緯度 Y': coordY,
      '監測時段': tod, '監測方法': method, '檢測機構許可證號': agencyCode, '其他檢測機構名稱': '',
      _siteCode: siteCode, _rawLocation: location,
    };

    const rows = [];
    if (isBN) {
      const hit = this.findCell(grid, /整體營建噪音值\(L3\)|整體营建噪音值\(L3\)/);
      let leq = '', lmax = '';
      if (hit) {
        const vals = this.valuesRightOf(grid, hit.r, hit.c, 2);
        leq = vals[0] || ''; lmax = vals[1] || '';
      }
      const common = { ...baseRow, '管制標準': '營建工程', '管制區': zone, '環境音量標準': '0', '頻率範圍': '20 Hz 至 20kHz', '檢測類別': '營建工程噪音', '監測單位': '16' };
      if (leq) rows.push({ ...common, '音源發聲特性': '均能音量(Leq)', '監測數值': this.formatNumber(leq) });
      if (lmax) rows.push({ ...common, '音源發聲特性': '最大音量(Lmax)', '監測數值': this.formatNumber(lmax) });
    } else if (isBV) {
      // 振動測值 header row: Lveq | Lvmax | Lv5 | Lv10 | ... with the numbers directly
      // underneath. Only the two metrics that have an official 音源發聲特性 code are
      // emitted — Lv5/Lv50/Lv90 have no valid code, so importing them would produce a
      // row the filing system can't accept.
      const methodRawV = this.labelValue(grid, /量測方法依據[:：]?/) || this.labelValue(grid, /採樣方法[:：]/);
      const methodV = this.extractMethodCode(methodRawV) || 'NIEA P204';
      const headerR = vibHit ? vibHit.r : -1;
      if (headerR >= 0) {
        const labelRow = grid[headerR] || [];
        const valueRow = grid[headerR + 1] || [];
        const colOf = (re) => { for (let c = 0; c < labelRow.length; c++) if (re.test(this.cellStr(labelRow[c]))) return c; return -1; };
        const metrics = [
          { col: colOf(/^Lveq$/i), item: '事件振動位準(Lveq)' },
          { col: colOf(/^Lvmax$/i), item: '最大振動位準(Lvmax)' },
        ];
        metrics.forEach(m => {
          if (m.col < 0) return;
          const v = this.cellStr(valueRow[m.col]);
          if (v === '' || isNaN(parseFloat(v))) return;
          rows.push({
            ...baseRow, '管制標準': '無', '管制區': '無', '環境音量標準': '0', '頻率範圍': '',
            '檢測類別': '振動', '監測時段': tod, '音源發聲特性': m.item,
            '監測單位': '159', '監測數值': this.formatNumber(v), '監測方法': methodV,
          });
        });
      }
    } else if (isLFN) {
      const hit = this.findCell(grid, /整體低頻噪音測值\(L1\)/);
      let leqLF = '';
      if (hit) {
        const vals = this.valuesRightOf(grid, hit.r, hit.c, 1);
        leqLF = vals[0] || '';
      }
      rows.push({
        ...baseRow, '管制標準': '營建工程', '管制區': zone, '環境音量標準': '0', '頻率範圍': '20 Hz 至 200 Hz',
        '檢測類別': '低頻噪音', '音源發聲特性': '均能音量(Leq,LF)', '監測單位': '16', '監測數值': this.formatNumber(leqLF),
      });
    }
    return rows.length ? rows : null;
  },

  /** N-xx(平日/假日) 24hr ambient noise, or V-xx(平日/假日) 24hr vibration. */
  parseNoise24hrSheet(sheetName, grid) {
    const isVib = /^V-?\d/i.test(sheetName)
      || !!this.findCell(grid, /Lv日\(Lv10\)=/)
      // 只寫夜間值的振動工作表原本認不出來，整張會被默默丟掉。
      || !!this.findCell(grid, /Lv夜\(Lv10\)=/);
    // Don't require the specific "(6~20)" hours — 道路交通噪音 reports use a
    // different daytime window (e.g. "(7~20)"), and hardcoding the hours meant
    // those sheets never matched here at all.
    const isNoise = /^N-?\d/i.test(sheetName) || !!this.findCell(grid, /^L日\(/);
    if (!isVib && !isNoise) return null;

    const dateRaw = this.labelValue(grid, /監測日期[:：]/);
    const dateISO = this.rocDateToISO(dateRaw);
    const location = this.labelValue(grid, /監測地點[:：]/) || '';
    const coordXHit = this.findCell(grid, /座標X/);
    const coordYHit = this.findCell(grid, /座標Y/);
    const coordX = coordXHit ? (this.valuesRightOf(grid, coordXHit.r, coordXHit.c, 1)[0] || '') : '';
    const coordY = coordYHit ? (this.valuesRightOf(grid, coordYHit.r, coordYHit.c, 1)[0] || '') : '';
    const agencyRaw = this.labelValue(grid, /採樣單位[:：]/);
    const agencyCode = this.reverseAgencyLookup(agencyRaw);
    const siteCode = this.labelValue(grid, /測點編號[:：]/) || '';

    // 樣品特性 tells us which noise category this report actually is — different
    // report types (一般環境噪音 vs 道路交通噪音 etc.) use different daytime windows
    // and regulatory bases, so this must be read from the report, not assumed.
    const sampleChar = this.labelValue(grid, /樣品特性[:：]/) || '';
    let noiseCategory = '環境噪音';
    if (/道路交通/.test(sampleChar)) noiseCategory = '道路交通噪音';
    else if (/公私場所/.test(sampleChar)) noiseCategory = '公私場所噪音';
    else if (/航空/.test(sampleChar)) noiseCategory = '航空噪音';

    const baseRow = {
      '日期(起)': dateISO, '時間(起)': '00:00:00',
      '日期(迄)': this.addDaysISO(dateISO, 1), '時間(迄)': '00:00:00',
      '監測地點': location, '座標系統': coordX ? '3' : '', '採樣座標-經度 X': coordX, '採樣座標-緯度 Y': coordY,
      '檢測機構許可證號': agencyCode, '其他檢測機構名稱': '',
      _siteCode: siteCode, _rawLocation: location,
    };

    const rows = [];
    if (isNoise) {
      const methodRaw = this.labelValue(grid, /採樣方法[:：]/);
      const method = this.extractMethodCode(methodRaw) || 'NIEA P201';
      // Find the "L日(...)" label without pinning to specific clock hours — different
      // control-zone classes (第一/二類 vs 第三/四類) and regulatory bases legitimately
      // use different day/evening/night windows (e.g. "L日(6~20)" vs "L日(7~20)"), so
      // matching only the "L日(" prefix works across all of them uniformly.
      const labelHit = this.findCell(grid, /^L日\(/);
      if (labelHit) {
        const labelRow = grid[labelHit.r];
        const valueRow = grid[labelHit.r + 1] || [];
        // Find each period's OWN column in the label row rather than assuming they
        // sit at label.col, +1, +2 — merged cells of different widths can space
        // "L日"/"L晚"/"L夜" (and their value row underneath) much further apart than
        // that, which silently dropped 晚間/夜間 on reports using a wider layout.
        const findCol = (regex) => {
          for (let c = 0; c < labelRow.length; c++) if (regex.test(this.cellStr(labelRow[c]))) return c;
          return -1;
        };
        const periods = [
          { key: 'L日', tod: '日間', col: findCol(/^L日/) },
          { key: 'L晚', tod: '晚間', col: findCol(/^L晚/) },
          { key: 'L夜', tod: '夜間', col: findCol(/^L夜/) },
        ];
        periods.forEach(p => {
          if (p.col < 0) return;
          const v = this.cellStr(valueRow[p.col]);
          if (v !== '' && !isNaN(parseFloat(v))) {
            rows.push({
              ...baseRow, '管制標準': '噪音管制法第7條第1項', '管制區': '', '環境音量標準': '', '頻率範圍': '20 Hz 至 20kHz',
              '檢測類別': noiseCategory, '監測時段': p.tod, '音源發聲特性': '均能音量(Leq)',
              '監測單位': '16', '監測數值': String(Math.round(parseFloat(v) * 10) / 10), '監測方法': method,
            });
          }
        });
      }
    } else if (isVib) {
      const methodRaw = this.labelValue(grid, /量測方法依據[:：]?/) || this.labelValue(grid, /採樣方法[:：]/);
      const method = this.extractMethodCode(methodRaw) || 'NIEA P204';
      // Extract BOTH the Lvd(10) and Leq-style vibration summaries when the report
      // provides them — which ones actually end up in the filing is the person's
      // call via the item-selection checklist, not something to decide here.
      // `secondary: true` means "read it, but don't tick it by default" — a 24-hour
      // environmental vibration filing reports Lvd(10); Lveq is stated in the report
      // as well but isn't what normally gets submitted, so it waits, unticked, in the
      // import preview's collapsed "其他測項" list instead of quietly doubling the
      // number of rows imported. (Construction vibration, BV, is different: Lveq and
      // Lvmax are both primary there, the same way BN reports Leq and Lmax.)
      // Lv10 的音源發聲特性是**看時段決定的**：日間 Lvd(10)、夜間 Lvn(10)
      // （官方噪音資料辭典第 21 列，d=day / n=night，見 schema.js 的 vibLv10ItemFor）。
      // v4.29 以前日夜共用一個 itemLabel，夜間那一筆填成 Lvd(10)，是錯的。
      //
      // Lveq 沒有日夜之分（官方代碼只有「事件振動位準(Lveq)」一個），所以它維持
      // 固定名稱——用 itemFor 這個函式而不是把規則寫在迴圈裡，就是為了讓兩者
      // 的差別留在各自的定義上，不會有人日後「順手」把規則套到 Lveq 去。
      const vibMetrics = [
        { labelKey: 'Lv10', dayRegex: /Lv日\(Lv10\)=/, nightRegex: /Lv夜\(Lv10\)=/, itemFor: (tod) => vibLv10ItemFor(tod) },
        { labelKey: 'Lveq', dayRegex: /Lv日\(Lveq\)=/, nightRegex: /Lv夜\(Lveq\)=/, itemFor: () => '事件振動位準(Lveq)', secondary: true },
      ];
      vibMetrics.forEach(metric => {
        const dayVal = this.labelValue(grid, metric.dayRegex);
        const nightVal = this.labelValue(grid, metric.nightRegex);
        [{ v: dayVal, tod: '日間' }, { v: nightVal, tod: '夜間' }].forEach(p => {
          if (p.v && !isNaN(parseFloat(p.v))) {
            const itemLabel = metric.itemFor(p.tod);
            if (!itemLabel) return; // 沒有對應官方代碼的時段就不產生資料列
            rows.push({
              ...baseRow, '管制標準': '無', '管制區': '無', '環境音量標準': '0', '頻率範圍': '',
              '檢測類別': '振動', '監測時段': p.tod, '音源發聲特性': itemLabel,
              '監測單位': '159',
              // 報告上印的就是一位小數（儲存格格式 0.0），所以先四捨五入到一位，
              // 再依官方「小數點2位數」的規定補零成 39.20——數值不變，只是寫法。
              '監測數值': String(Math.round(parseFloat(p.v) * 10) / 10),
              '監測方法': method,
              _secondaryItem: !!metric.secondary,
            });
          }
        });
      });
    }
    return rows.length ? rows : null;
  },

  /** For reports where "label：value" sits in a single cell (unlike separate label/value cells elsewhere). */
  labelValueSameCell(grid, regex) {
    const hit = this.findCell(grid, regex);
    if (!hit) return null;
    const text = this.cellStr(grid[hit.r][hit.c]);
    const parts = text.split(/[:：]/);
    return parts.length > 1 ? parts.slice(1).join('：').trim() : null;
  },

  /**
   * 24hr ambient air-quality report (continuous analyzers + TSP/PM2.5 side table).
   * Column positions are anchored by header-label text, not fixed indices, since a
   * merged-cell layout otherwise shifts. Extracts EVERY pollutant column actually
   * present in the report (not a fixed list) — which items end up in the filing is
   * a decision the person makes via the item-selection checklist in the import UI,
   * since different projects/plans report different sets of items.
   */
  /**
   * The summary rows a 24hr ambient air-quality report states for each pollutant.
   *
   * `plainItemName: true` means rows from this statistic keep the bare pollutant name
   * ("CO"), because the daily average is what a filing normally reports and the
   * official template's 檢測項目 column is written that way. Any OTHER statistic gets
   * its name spelled out in 檢測項目 ("CO最大8小時平均值") — otherwise two rows for the
   * same site, day and pollutant would be indistinguishable both to a reader and to
   * this app's own duplicate detection, which keys on date + time + site + item.
   */
  AIR_STAT_ROWS: [
    { key: 'avg', label: '日平均值', re: /^日平均值或$|^日平均值$/, plainItemName: true },
    { key: 'maxHour', label: '最大小時平均值', re: /^最大小時平均值$/ },
    { key: 'minHour', label: '最小小時平均值', re: /^最小小時平均值$/ },
    { key: 'max8', label: '最大8小時平均值', re: /^最大8小時平均值$|^最大八小時平均值$/ },
  ],
  /** The hourly block itself, offered as one more choice alongside the summary rows.
   *  Its rows carry their own clock hour, so they never collide with each other. */
  AIR_HOURLY_STAT: { key: 'hourly', label: '每小時測值', plainItemName: true },

  AIR_POLLUTANT_DEFS: [
    { key: 'SO2', unit: '113' }, { key: 'NO2', unit: '113', methodFrom: 'NOX' }, { key: 'NOx', unit: '113' },
    { key: 'NO', unit: '113' }, { key: 'CO', unit: '113' }, { key: 'O3', unit: '113' },
    { key: 'CH4', unit: '113' }, { key: 'NMHC', unit: '113' }, { key: 'THC', unit: '113' },
    { key: 'PM10', unit: '127' },
  ],

  /**
   * 落塵量 (dustfall) style report: unlike the 24hr hourly table, this is a simple
   * vertical list with ONE ROW PER SITE (not per hour), all reporting the same single
   * item. Detected by the distinctive column combination 測點編號+監測地點+監測日期
   * all present as separate header columns, which the hourly table format doesn't have.
   */
  parseAirDustfallSheet(grid) {
    const headerHit = this.findCell(grid, /^監測項目$/);
    if (!headerHit) return null;
    const headerRow = grid[headerHit.r];
    const colOf = (regex) => {
      for (let c = 0; c < headerRow.length; c++) if (regex.test(this.cellStr(headerRow[c]))) return c;
      return -1;
    };
    const cols = {
      item: headerHit.c,
      siteCode: colOf(/^測點編號$/),
      location: colOf(/^監測地點$/),
      dateRange: colOf(/^監測日期$/),
      value: colOf(/濃度/),
    };
    if (cols.siteCode < 0 || cols.location < 0 || cols.dateRange < 0 || cols.value < 0) return null;

    const valueHeaderText = this.cellStr(headerRow[cols.value]);
    const unitMatch = valueHeaderText.match(/濃度\s*(.+)$/);
    const unitText = unitMatch ? unitMatch[1].trim() : '';
    const unitLookup = unitText ? this.reverseUnitLookup(unitText) : { code: '', confident: false };

    const receiveDateRaw = this.labelValueSameCell(grid, /收樣日期[:：]/) || this.labelValueSameCell(grid, /報告日期[:：]/);
    const rocYearMatch = String(receiveDateRaw || '').match(/(\d{2,3})年/);
    const rocYear = rocYearMatch ? parseInt(rocYearMatch[1], 10) : null;

    const methodRaw = this.labelValueSameCell(grid, /採樣方法[:：]/);
    const method = this.extractMethodCode(methodRaw);
    const agencyRaw = this.labelValueSameCell(grid, /採樣單位[:：]/);
    const agencyCode = this.reverseAgencyLookup(agencyRaw);

    const rows = [];
    for (let r = headerHit.r + 1; r < grid.length; r++) {
      const row = grid[r];
      const itemName = this.cellStr(row[cols.item]);
      const location = this.cellStr(row[cols.location]);
      if (!itemName && !location) continue;
      if (/以下空白/.test(row.map(c => this.cellStr(c)).join(''))) break;
      const siteCode = this.cellStr(row[cols.siteCode]);
      const dateRangeRaw = this.cellStr(row[cols.dateRange]);
      const valRaw = this.cellStr(row[cols.value]);
      if (!location || valRaw === '') continue;

      let dateStart = '', dateEnd = '';
      const trMatch = dateRangeRaw.match(/(\d{1,2})\/(\d{1,2})\s*~\s*(\d{1,2})\/(\d{1,2})/);
      if (trMatch && rocYear) {
        const y = rocYear + 1911;
        const pad = n => String(n).padStart(2, '0');
        const [, m1, d1, m2, d2] = trMatch;
        dateStart = `${y}-${pad(m1)}-${pad(d1)}`;
        const y2 = parseInt(m2, 10) < parseInt(m1, 10) ? y + 1 : y;
        dateEnd = `${y2}-${pad(m2)}-${pad(d2)}`;
      }

      const { cmp, val, note } = this.parseValueCell(valRaw);
      rows.push({
        '日期(起)': dateStart, '時間(起)': '', '日期(迄)': dateEnd, '時間(迄)': '',
        '採樣地點': location, '座標系統': '', '採樣座標-經度 X': '', '採樣座標-緯度 Y': '',
        '場所編號': '', '採樣地點高度(公尺)': '', '污染物採樣高度(公尺)': '', '管制編號': '', '煙道編號': '',
        '檢測類別': '周界空氣品質', '檢測項目': itemName || '落塵量',
        '檢測濃度/質量單位': unitLookup.code, '其他檢測濃度/質量單位': unitLookup.code ? '' : unitText,
        '比較關係': cmp, '檢測數值': /^[\d.]+$/.test(val) ? this.formatNumber(val, 3) : val,
        '檢測方法': method, '檢測機構許可證號': agencyCode, '其他檢測機構名稱': '',
        '備註': note || '',
        _siteCode: siteCode, _rawLocation: location,
        _uncertainUnit: !!unitText && !unitLookup.confident,
      });
    }
    return rows.length ? rows : null;
  },

  parseAirQualitySheet(grid) {
    const title = this.cellStr(grid[1]?.[0] || '') + this.cellStr(grid[1]?.[18] || '');
    if (!/空氣品質檢測報告/.test(title) && !this.findCell(grid, /空氣品質檢測報告/)) return null;

    const headerRowHit = this.findCell(grid, /^監測項目$/);
    if (!headerRowHit) return null;
    const headerRow = headerRowHit.r; // row 13 in the sample: pollutant column headers
    const colOf = (regex) => {
      const row = grid[headerRow];
      for (let c = 0; c < row.length; c++) if (regex.test(this.cellStr(row[c]))) return c;
      return -1;
    };
    const cols = { TSP: colOf(/^TSP$/), method: colOf(/^檢測方法$/), location: colOf(/^監測地點$/),
      timeRange: colOf(/^監測時間$/), pm25: colOf(/PM2\.5濃度/), standard: colOf(/^空氣品質標準$/) };
    // dynamically locate whichever pollutant columns this report actually has
    this.AIR_POLLUTANT_DEFS.forEach(def => { cols[def.key] = colOf(new RegExp(`^${def.key}$`, 'i')); });

    // hourly data block: rows after the unit row (headerRow+2) up to the "日平均值或" row
    const unitRow = headerRow + 2;
    // Every summary row this report format can carry. The report states several
    // DIFFERENT statistics for the same pollutant on the same day — the daily
    // average, the highest and lowest single-hour averages, and the highest rolling
    // 8-hour average — and which one belongs in a filing is the person's decision,
    // not something to hard-code. All of them are read; the import preview then asks.
    const statHits = this.AIR_STAT_ROWS
      .map(def => ({ def, hit: this.findCell(grid, def.re) }))
      .filter(x => x.hit);
    if (statHits.length === 0) return null;
    // The hourly block runs from just under the unit row down to whichever summary
    // row sits HIGHEST on the sheet — taken by row position, not by the order the
    // statistics happen to be declared in AIR_STAT_ROWS, so a report that prints them
    // in a different order can't leave a summary row inside the hourly block.
    const avgHit = statHits.reduce((best, x) => (!best || x.hit.r < best.r ? x.hit : best), null);

    const location = this.cellStr(grid[unitRow]?.[cols.location]) || (this.labelValueSameCell(grid, /監測地點[:：]/) || '');
    const siteCode = this.labelValueSameCell(grid, /測點編號[:：]/) || '';
    const timeRangeRaw = this.cellStr(grid[unitRow]?.[cols.timeRange]); // e.g. "2/12 ~ 2/13"
    const receiveDateRaw = this.labelValueSameCell(grid, /收樣日期[:：]/) || this.labelValueSameCell(grid, /報告日期[:：]/);
    const rocYearMatch = String(receiveDateRaw || '').match(/(\d{2,3})年/);
    const rocYear = rocYearMatch ? parseInt(rocYearMatch[1], 10) : null;

    let dateStart = '', dateEnd = '';
    const trMatch = timeRangeRaw.match(/(\d{1,2})\/(\d{1,2})\s*~\s*(\d{1,2})\/(\d{1,2})/);
    if (trMatch && rocYear) {
      const y = rocYear + 1911;
      const pad = n => String(n).padStart(2, '0');
      const [, m1, d1, m2, d2] = trMatch;
      dateStart = `${y}-${pad(m1)}-${pad(d1)}`;
      const y2 = parseInt(m2, 10) < parseInt(m1, 10) ? y + 1 : y;
      dateEnd = `${y2}-${pad(m2)}-${pad(d2)}`;
    }

    // Not every version of this report has a 監測時間 column. A very common variant
    // writes the sampling window DOWN THE FIRST COLUMN instead, one character per
    // merged cell — "115 / 年 / 6 / 月 / 25 / 日 / 至 / 115 / 年 / 6 / 月 / 26 / 日" —
    // which is invisible to a column lookup. Stitching that column back together and
    // splitting it on 至 recovers both ends of the window; without this the whole
    // report imports with EMPTY dates, which is exactly the sort of silent gap that
    // only shows up after the filing is submitted.
    if (!dateStart) {
      const stitched = [];
      for (let r = headerRow; r < avgHit.r; r++) stitched.push(this.cellStr(grid[r]?.[0]));
      const joined = stitched.join('');
      const parts = joined.split(/至|~|～|－|—/);
      const a = DateTimeUtil.toISODate(parts[0] || '');
      const b = parts.length > 1 ? DateTimeUtil.toISODate(parts[1]) : '';
      if (/^\d{4}-\d{2}-\d{2}$/.test(a)) {
        dateStart = a;
        dateEnd = /^\d{4}-\d{2}-\d{2}$/.test(b) ? b : a;
      }
    }

    // first/last hourly rows determine the start/end clock time (rolling 24hr window)
    let timeStart = '', timeEnd = '';
    for (let r = unitRow + 1; r < avgHit.r; r++) {
      const hourCell = this.cellStr(grid[r]?.[1]);
      const hm = hourCell.match(/(\d{1,2})\s*~\s*(\d{1,2})/);
      if (hm) { if (!timeStart) timeStart = `${hm[1].padStart(2, '0')}:00:00`; timeEnd = `${hm[2].padStart(2, '0')}:00:00`; }
    }

    // TSP: single discrete reading somewhere in its hourly column (not itself hourly-averaged)
    let tspVal = '';
    if (cols.TSP >= 0) {
      for (let r = unitRow + 1; r < avgHit.r; r++) {
        const v = this.cellStr(grid[r]?.[cols.TSP]);
        if (v !== '' && !isNaN(parseFloat(v))) { tspVal = v; break; }
      }
    }

    const methodNoteHit = this.findCell(grid, /檢測方法[:：].*NIEA/);
    const methodMap = {};
    if (methodNoteHit) {
      const noteText = this.cellStr(grid[methodNoteHit.r][methodNoteHit.c]);
      const re = /([A-Za-z0-9.]+):NIEA\s*([A-Z]?\d+)/g;
      let m;
      while ((m = re.exec(noteText))) methodMap[m[1].toUpperCase()] = `NIEA ${m[2]}`;
    }
    const pm25Method = this.extractMethodCode(this.cellStr(grid[unitRow]?.[cols.method])) || 'NIEA A205';

    const agencyRaw = this.labelValueSameCell(grid, /採樣單位[:：]/);
    const agencyCode = this.reverseAgencyLookup(agencyRaw);

    const baseRow = {
      '日期(起)': dateStart, '時間(起)': timeStart, '日期(迄)': dateEnd, '時間(迄)': timeEnd,
      '採樣地點': location, '座標系統': '', '採樣座標-經度 X': '', '採樣座標-緯度 Y': '',
      '場所編號': '', '採樣地點高度(公尺)': '', '污染物採樣高度(公尺)': '', '管制編號': '', '煙道編號': '',
      '檢測類別': '周界空氣品質', '檢測機構許可證號': agencyCode, '其他檢測機構名稱': '',
      _siteCode: siteCode, _rawLocation: location,
    };

    const rows = [];
    const skippedPlaceholderItems = [];
    // Whether a pollutant column is actually monitored at this site is determined
    // by the spreadsheet's own hidden-column flag, not by guessing from the data —
    // a lab reusing one shared report template across sites with different
    // instrumentation hides the columns for items a given site doesn't measure,
    // while the placeholder text stays in the (hidden) cells. Trying to infer this
    // from data patterns instead (e.g. "this column reads the same value all day")
    // was tried and is unreliable both ways: it can wrongly flag a genuinely clean
    // site's real SO2/O3 readings as fake, and wrongly accept a hidden placeholder
    // column that happens to vary. The hidden-column flag has neither failure mode.
    const hiddenCols = grid._hiddenCols || new Set();

    // Use the same ND/"< X" parser as the water-table reader — a daily average like
    // "< 0.3" (below detection limit) must never be silently dropped just because
    // parseFloat can't read the "<" prefix; every pollutant column can show this.
    // "Has content" now checks cmp/note rather than val alone, since ND's val is
    // intentionally blank (see parseValueCell) — checking val==='' alone would
    // wrongly treat a real "ND" reading as an empty cell to skip.
    // One pass PER SUMMARY ROW the report states (日平均值 / 最大小時平均值 /
    // 最小小時平均值 / 最大8小時平均值). Each row is tagged with `_statKind`, and the
    // import preview turns those tags into a checklist so the person chooses which
    // statistic actually goes into the filing — 日平均值 is pre-selected because that
    // is what these filings normally report, but nothing is silently discarded.
    statHits.forEach(({ def: statDef, hit }) => {
      const statRow = grid[hit.r] || [];
      this.AIR_POLLUTANT_DEFS.forEach(def => {
        const col = cols[def.key];
        if (col < 0) return;
        if (hiddenCols.has(col)) {
          if (statDef.key === 'avg' && !skippedPlaceholderItems.includes(def.key)) skippedPlaceholderItems.push(def.key);
          return;
        }
        const v = this.cellStr(statRow[col]);
        if (v === '') return;
        const { cmp, val, note } = this.parseValueCell(v);
        const hasContent = cmp !== '' || note !== '' || (val !== '' && !isNaN(parseFloat(val)));
        if (!hasContent) return;
        const methodKey = (def.methodFrom || def.key).toUpperCase();
        rows.push({
          ...baseRow,
          '檢測項目': statDef.plainItemName ? def.key : `${def.key}${statDef.label}`,
          '檢測濃度/質量單位': def.unit,
          '比較關係': cmp, '檢測數值': /^[\d.]+$/.test(val) ? this.formatNumber(val, 3) : val,
          '檢測方法': methodMap[methodKey] || '', '備註': note || '',
          _statKind: statDef.key, _statLabel: statDef.label,
        });
      });
    });

    // The hourly block itself — 24 rows per pollutant, each with its own clock hour.
    // Offered as one more choice in the same checklist for anyone who needs to file
    // hour-by-hour readings rather than a summary.
    this._pushAirHourlyRows(grid, cols, hiddenCols, methodMap, baseRow, unitRow, avgHit.r, dateStart, dateEnd, rows);

    if (cols.TSP >= 0 && hiddenCols.has(cols.TSP)) {
      skippedPlaceholderItems.push('TSP');
    } else if (tspVal) {
      const { cmp, val, note } = this.parseValueCell(tspVal);
      rows.push({
        ...baseRow, '檢測項目': 'TSP', '檢測濃度/質量單位': '127',
        '比較關係': cmp, '檢測數值': /^[\d.]+$/.test(val) ? this.formatNumber(val, 3) : val,
        '檢測方法': methodMap.TSP || '', '備註': note || '',
      });
    }
    if (cols.pm25 >= 0 && hiddenCols.has(cols.pm25)) {
      skippedPlaceholderItems.push('PM2.5');
    } else if (cols.pm25 >= 0) {
      const pm25Raw = this.cellStr(grid[unitRow]?.[cols.pm25]);
      if (pm25Raw !== '') {
        const { cmp, val, note } = this.parseValueCell(pm25Raw);
        const hasContent = cmp !== '' || note !== '' || (val !== '' && !isNaN(parseFloat(val)));
        if (hasContent) {
          rows.push({
            ...baseRow, '檢測項目': 'PM2.5', '檢測濃度/質量單位': '127',
            '比較關係': cmp, '檢測數值': /^[\d.]+$/.test(val) ? this.formatNumber(val, 3) : val,
            '檢測方法': pm25Method, '備註': note || '',
          });
        }
      }
    }
    // tag every row with which columns this sheet skipped as likely-unmonitored
    // placeholders, so the import UI can tell the person what was left out and why
    // (rather than silently never showing those items at all).
    if (skippedPlaceholderItems.length) {
      rows.forEach(r => { r._skippedPlaceholderItems = skippedPlaceholderItems; });
    }
    // Rows the summary passes produced carry a stat tag; TSP/PM2.5 (single 24-hour
    // composite samples, not hourly statistics) don't, so tag them as the daily
    // figure — otherwise they'd be filtered out whenever the person picks a
    // statistic, even though there is only ever one value for them.
    rows.forEach(r => { if (!r._statKind) { r._statKind = 'avg'; r._statLabel = '日平均值'; } });
    return rows.length ? rows : null;
  },

  /** Reads the hour-by-hour block of a 24hr air report into one row per pollutant per
   *  hour. The report writes hours as "08 ~ 09" and simply rolls past midnight without
   *  restating the date, so the calendar day is advanced whenever the clock wraps. */
  _pushAirHourlyRows(grid, cols, hiddenCols, methodMap, baseRow, unitRow, endRow, dateStart, dateEnd, out) {
    let day = dateStart;
    let prevHour = null;
    for (let r = unitRow + 1; r < endRow; r++) {
      const hm = this.cellStr(grid[r]?.[1]).match(/(\d{1,2})\s*~\s*(\d{1,2})/);
      if (!hm) continue;
      const h1 = parseInt(hm[1], 10), h2 = parseInt(hm[2], 10);
      if (prevHour !== null && h1 < prevHour) day = dateEnd || this.addDaysISO(day, 1);
      prevHour = h1;
      const pad = n => String(n).padStart(2, '0');
      const tStart = `${pad(h1)}:00:00`;
      // "23 ~ 24" means 23:00 today through 00:00 tomorrow
      const rollsOver = h2 >= 24 || h2 <= h1;
      const tEnd = `${pad(h2 % 24)}:00:00`;
      const endDay = rollsOver ? this.addDaysISO(day, 1) : day;
      this.AIR_POLLUTANT_DEFS.forEach(def => {
        const col = cols[def.key];
        if (col < 0 || hiddenCols.has(col)) return;
        const v = this.cellStr(grid[r]?.[col]);
        if (v === '') return;
        const { cmp, val, note } = this.parseValueCell(v);
        const hasContent = cmp !== '' || note !== '' || (val !== '' && !isNaN(parseFloat(val)));
        if (!hasContent) return;
        const methodKey = (def.methodFrom || def.key).toUpperCase();
        out.push({
          ...baseRow,
          '日期(起)': day, '時間(起)': tStart, '日期(迄)': endDay, '時間(迄)': tEnd,
          '檢測項目': def.key, '檢測濃度/質量單位': def.unit,
          '比較關係': cmp, '檢測數值': /^[\d.]+$/.test(val) ? this.formatNumber(val, 3) : val,
          '檢測方法': methodMap[methodKey] || '', '備註': note || '',
          _statKind: this.AIR_HOURLY_STAT.key, _statLabel: this.AIR_HOURLY_STAT.label,
        });
      });
    }
  },

  /** W (河川/地下水等) / WU (放流水) style vertical water-quality test-item table. */
  parseWaterTableSheet(grid) {
    return this.parseLabItemTableSheet(grid, 'water');
  },

  /**
   * 地質 lab report — the 底泥檢測報告 (e.g. 06525DS27) and its 土壤/廢棄物 siblings.
   * These use exactly the same vertical "檢驗項目 / 檢測值 / 偵測極限 / 單位 / 檢測方法"
   * table as the water reports from the same consultant, with a 樣品特性 of 底泥/土壤
   * instead of a water type — so it shares the parser below and only differs in which
   * template's field names it writes into and how 檢測類別 is derived.
   */
  parseGeoSedimentSheet(grid) {
    return this.parseLabItemTableSheet(grid, 'geo');
  },

  /**
   * Shared reader for the consultant's vertical lab-report table, used by both the
   * water and the 地質 (底泥/土壤) categories.
   *
   * `kind` decides two things and nothing else:
   *   - which template's field names the output rows use (採樣深度/採樣水深 differ), and
   *   - how 樣品特性 becomes 檢測類別.
   * It also acts as the gate that stops the two from stealing each other's reports
   * during batch auto-detection: a sheet whose 樣品特性 says 底泥 is refused by the
   * water reader, and one that says 放流水 is refused by the geo reader.
   */
  parseLabItemTableSheet(grid, kind) {
    const hasItemHeader = this.findCell(grid, /檢驗項目|檢測項目/);
    if (!hasItemHeader) return null;

    const sampleCharEarly = this.labelValue(grid, /樣品特性[:：]/) || '';
    if (kind === 'water' && this.isGeoSampleType(sampleCharEarly)) return null;
    if (kind === 'geo') {
      // Only claim the sheet when it actually looks like a solid-sample report:
      // either 樣品特性 says so, or the report titles itself 底泥/土壤/廢棄物.
      const title = [0, 1, 2, 3].map(r => this.cellStr(grid[r]?.[0]) + this.cellStr(grid[r]?.[1])).join(' ');
      const titleSaysGeo = /底泥|底質|土壤|廢棄物|事業廢棄物/.test(title)
        || !!this.findCell(grid, /底泥檢測報告|土壤檢測報告|廢棄物檢測報告/);
      if (!this.isGeoSampleType(sampleCharEarly) && !titleSaysGeo) return null;
      if (this.isWaterSampleType(sampleCharEarly) && !this.isGeoSampleType(sampleCharEarly)) return null;
    }

    const location = this.labelValue(grid, /採樣地點[:：]/) || '';
    const sampleTimeRaw = this.labelValue(grid, /採樣時間[:：]/) || this.labelValue(grid, /採樣日期[:：]/);
    const sampleDateISO = this.rocDateToISO(sampleTimeRaw);
    let sampleTime = '';
    if (sampleTimeRaw) {
      const m = String(sampleTimeRaw).match(/(\d{1,2})時(\d{1,2})分/);
      if (m) sampleTime = `${m[1].padStart(2, '0')}:${m[2].padStart(2, '0')}:00`;
      else {
        const hm = String(sampleTimeRaw).match(/(\d{1,2}):(\d{2})/);
        if (hm) sampleTime = `${hm[1].padStart(2, '0')}:${hm[2]}:00`;
      }
    }
    const sampleChar = sampleCharEarly;
    // 地質: the person picks 檢測類別 themselves in the import preview (see
    // SITE_PROFILE_FIELDS.geo) — leave it blank here rather than guessing.
    const category = kind === 'geo' ? '' : this.sampleTypeToCategory(sampleChar);
    const methodRaw = this.labelValue(grid, /採樣方法[:：]/);
    const defaultMethod = this.extractMethodCode(methodRaw);
    const agencyRaw = this.labelValue(grid, /採樣單位[:：]/) || this.labelValue(grid, /公司名稱[:：]/);
    const agencyCode = this.reverseAgencyLookup(agencyRaw);
    const siteCode = this.labelValue(grid, /測點編號[:：]/) || '';

    const headerHit = hasItemHeader;

    // This reader locates each value BY POSITION (item, then value, then 偵測極限,
    // then 單位, then 檢測方法), which is only correct for the template it was written
    // against. Another lab can use the very same heading words in a different order —
    // 檢驗項目 / 檢測結果 / 單位 / 方法偵測極限 / 分析方法 is a real and common variant —
    // and reading that by position silently swaps the unit and the detection limit.
    // So: confirm the heading order before claiming the sheet, and hand anything else
    // to AutoDetect, which maps by heading text instead of by position.
    const headerCells = [];
    for (let r = headerHit.r; r < Math.min(headerHit.r + 3, grid.length); r++) {
      (grid[r] || []).forEach((v, c) => {
        const s = this.cellStr(v);
        if (s !== '' && headerCells[c] === undefined) headerCells[c] = s;
      });
    }
    const colOfLabel = (re) => headerCells.findIndex(s => s !== undefined && re.test(s));
    const cItem = colOfLabel(/檢驗項目|檢測項目/);
    const cValue = colOfLabel(/檢測值|檢測結果|檢驗結果|測定值|結\s*果/);
    const cLimit = colOfLabel(/偵測\s*極限|檢測\s*極限/);
    const cUnit = colOfLabel(/^單\s*位$/);
    const cMethod = colOfLabel(/檢測方法|分析方法|檢驗方法/);
    if (cLimit >= 0 && cUnit >= 0 && cUnit < cLimit) return null;
    if (cUnit >= 0 && cMethod >= 0 && cMethod < cUnit) return null;
    /*
     * 有抓到欄位位置就**照欄位讀**，不要「刪掉空白格再照順序取」。
     *
     * 舊寫法是 row.filter(v => v !== '') 之後用 cells[0..4]。只要中間有一格是空的，
     * 整列就往左移一格：水溫、pH、導電度、懸浮固體這些本來就沒有偵測極限的項目
     * （很多實驗室那一格是空的，不寫「--」），單位會被讀成偵測極限、
     * 檢測方法會被讀成單位——而且因為單位讀不到，程式會走「沒有單位欄 ⇒ 無(161)」
     * 那條路並且回報 confident:true，所以連「單位可能有問題」的提醒都不會出現。
     * 實測：`["水溫","28.5","","℃","NIEA W217"]` → 單位 161（應為 4）。
     *
     * 找不到欄位位置時（版型太特別）才退回原本的位置式讀法，行為與舊版相同。
     */
    const byColumn = cItem >= 0 && cValue >= 0 && cValue > cItem;

    // find the literal "檢測值" marker row that precedes the item rows
    let startRow = headerHit.r + 1;
    for (let r = headerHit.r + 1; r < Math.min(headerHit.r + 4, grid.length); r++) {
      if (grid[r].some(v => /^檢測值$/.test(this.cellStr(v)))) { startRow = r + 1; break; }
    }

    const rows = [];
    for (let r = startRow; r < grid.length; r++) {
      const row = grid[r];
      const cells = row.map(v => this.cellStr(v)).filter(v => v !== '');
      if (cells.length === 0) continue;
      const first = cells[0];
      if (/以下空白|備註|聲明書|本報告|公司名稱|負責人|檢驗室主管|第\d+頁/.test(first)) break;
      if (cells.length < 2) continue;
      const dash = (v) => (/^[-–—─]{1,2}$/.test(v || '') ? '' : (v || ''));
      let itemName, valueRaw, limitRaw, unitText, methodText;
      if (byColumn) {
        const at = (c) => (c >= 0 ? this.cellStr(row[c]) : '');
        itemName = this.normalizeItemName(at(cItem));
        valueRaw = at(cValue);
        limitRaw = dash(at(cLimit));
        unitText = at(cUnit);
        methodText = at(cMethod);
        if (itemName === '' || valueRaw === '') continue;
      } else {
        itemName = this.normalizeItemName(cells[0]);
        valueRaw = cells[1];
        let idx = 2;
        // "--", "—", "─" and friends all mean "no detection limit applies here"
        limitRaw = dash(cells[idx]);
        if (cells[idx] !== undefined) idx++;
        unitText = '';
        /*
         * 檢測方法不是只有 NIEA。CNS、EPA、APHA、電極法、碘定量法都是實驗室
         * 真的會寫的東西，而 extractMethodCode 本來就認得它們——只有這裡寫死
         * 了 /^NIEA/，於是非 NIEA 的方法會被當成單位吃掉，該列的檢測方法則
         * 退回工作表上的「採樣方法」（那是採樣，不是分析方法）。
         */
        const looksLikeMethod = (v) => !!v && (/^(NIEA|CNS|EPA|ASTM|APHA|ISO)/i.test(v) || /法$/.test(v));
        if (cells[idx] && !looksLikeMethod(cells[idx])) { unitText = cells[idx]; idx++; }
        methodText = '';
        if (cells[idx] && looksLikeMethod(cells[idx])) { methodText = cells[idx]; idx++; }
      }

      const { cmp, val, note } = this.parseValueCell(valueRaw);
      // No unit column at all for this item (e.g. pH is dimensionless) means "無" —
      // official unit code 161 — not a blank/unknown unit. Only missing/unmatched
      // unit *text* (unitText present but not found in the code table) is uncertain.
      const unitLookup = unitText ? this.reverseUnitLookup(unitText, itemName) : (this.ITEM_UNIT_OVERRIDES[itemName] ? { code: this.ITEM_UNIT_OVERRIDES[itemName], confident: true } : { code: '161', confident: true });
      const valFormatted = /^[\d.]+$/.test(val) ? this.formatNumber(val, 3) : val;
      // 檢測極限 only makes sense alongside "ND" (below detection limit entirely)
      // or "<X" (below this specific threshold) — a plain measured value or a
      // ">X" reading isn't characterized by a detection limit, so leave it blank
      // even if the report's own 偵測極限 column happened to have something in it.
      const limitApplies = cmp === 'ND' || cmp === '<';
      // A limit that isn't a bare number is passed through UNCHANGED rather than
      // silently blanked. 檢測極限 may only hold a number in the filing, but the
      // person is the one who should decide what to do with a report that wrote
      // "<0.002" or "0.0004 mg/L" there — the import preview lists every such value
      // and offers 清成空白 / 只保留數字 / 照原樣匯入. Blanking it here would hide
      // both the report's real content and the choice.
      const limitFormatted = !limitApplies ? ''
        : (/^[\d.]+$/.test(limitRaw) ? this.formatNumber(limitRaw, 3) : limitRaw);

      const depthFields = kind === 'geo'
        ? { '採樣深度(公尺)': '' }
        : { '採樣深度(公尺)': '', '採樣水深(公尺)': '' };
      rows.push({
        '日期(起)': sampleDateISO, '時間(起)': sampleTime, '日期(迄)': sampleDateISO, '時間(迄)': sampleTime,
        '採樣地點': location, '座標系統': '', '採樣座標-經度 X': '', '採樣座標-緯度 Y': '',
        ...depthFields, '管制編號': '',
        '檢測類別': category, '檢測項目': itemName,
        '檢測濃度/質量單位': unitLookup.code, '其他檢測濃度/質量單位': unitLookup.code ? '' : unitText,
        '比較關係': cmp, '檢測數值': valFormatted, '檢測極限': limitFormatted,
        '檢測方法': this.extractMethodCode(methodText) || defaultMethod,
        '檢測機構許可證號': agencyCode, '其他檢測機構名稱': '',
        '備註': note || '',
        _siteCode: siteCode, _rawLocation: location,
        _uncertainUnit: unitText && !unitLookup.confident,
      });
    }
    return rows.length ? rows : null;
  },

  /**
   * Try every known parser for a sheet, category-scoped. Returns rows[] or null.
   *
   * Two passes, in this order:
   *   1. The template-specific parsers above — precise, written against known report
   *      layouts, and always preferred when one of them recognizes the sheet.
   *   2. AutoDetect — the layout-agnostic reader that hunts for the seven required
   *      fields by heading text. It fires only when no specific parser matched, so an
   *      unfamiliar lab's report still yields data instead of nothing at all. Rows it
   *      produces carry `_autoDetected` so the import preview can flag them for
   *      manual checking. Pass `{ allowAutoDetect:false }` to run pass 1 only (used by
   *      batch auto-detection, where guessing which CATEGORY an unknown sheet belongs
   *      to would be a guess on top of a guess).
   */
  parseSheet(category, sheetName, grid, { allowAutoDetect = true } = {}) {
    if (!grid || grid.length === 0) return null;
    let rows = null;
    if (category === 'noise') {
      rows = this.parseNoiseEventSheet(grid, sheetName) || this.parseNoise24hrSheet(sheetName, grid);
    } else if (category === 'water') {
      rows = this.parseWaterTableSheet(grid);
    } else if (category === 'geo') {
      rows = this.parseGeoSedimentSheet(grid);
    } else if (category === 'air') {
      rows = this.parseAirDustfallSheet(grid) || this.parseAirQualitySheet(grid);
    }
    if (rows && rows.length) return rows;
    if (allowAutoDetect && typeof AutoDetect !== 'undefined') {
      return AutoDetect.parseSheet(category, sheetName, grid);
    }
    return null;
  },

  /** Parse every sheet of a workbook (given as {sheetName: grid}) for a category. */
  parseWorkbook(category, sheetGrids) {
    const rows = [];
    const matchedSheets = [];
    const skippedSheets = [];
    for (const [sheetName, grid] of Object.entries(sheetGrids)) {
      const parsed = this.parseSheet(category, sheetName, grid);
      if (parsed && parsed.length) {
        rows.push(...parsed);
        matchedSheets.push(sheetName);
      } else {
        skippedSheets.push(sheetName);
      }
    }
    // Group rows by the report's own site code (測點編號) when available — this is
    // the reliable key, because the raw report's free-text location name doesn't
    // always match the official site name used in the filing (e.g. a construction-
    // noise report covers "Y1車站~Y4車站" but the filing names the specific point
    // "YC01標工區周界外"). Rows without a site code fall back to grouping by their
    // raw location text.
    const sites = {}; // key -> { siteCode, rawLocation, rowIndices: [] }
    rows.forEach((row, i) => {
      const key = row._siteCode || row._rawLocation || `row${i}`;
      if (!sites[key]) sites[key] = { siteCode: row._siteCode || '', rawLocation: row._rawLocation || '', rowIndices: [] };
      sites[key].rowIndices.push(i);
    });

    return { rows, matchedSheets, skippedSheets, sites };
  },
};
