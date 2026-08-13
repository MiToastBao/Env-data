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
  },

  // ---------- grid utilities ----------
  // grid: array of rows, each row an array of cell values (strings/numbers/''),
  // as produced by XLSX.utils.sheet_to_json(ws, {header:1, defval:''}).

  cellStr(v) { return v === undefined || v === null ? '' : String(v).trim(); },

  /** Scan the whole grid for the first cell matching regex; return {r,c} or null. */
  findCell(grid, regex) {
    for (let r = 0; r < grid.length; r++) {
      const row = grid[r];
      for (let c = 0; c < row.length; c++) {
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

  /** Find a label anywhere in the grid and return the next non-empty value(s) to its right. */
  labelValue(grid, regex, n = 1) {
    const hit = this.findCell(grid, regex);
    if (!hit) return n === 1 ? null : [];
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
  /** Extract a bare NIEA-style method code, dropping the version suffix: "NIEA W217.51A" -> "NIEA W217" */
  extractMethodCode(text) {
    if (!text) return '';
    const m = String(text).match(/NIEA\s*([A-Z]?\d+)/i);
    return m ? `NIEA ${m[1]}` : String(text).trim();
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
    for (const [code, name] of Object.entries(AGENCY_CODES)) {
      if (code === 'AA') continue;
      if (clean.includes(name) || name.includes(clean)) return code;
    }
    return '';
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

  /** Normalize a raw test-item name to the conventional short form used in filings. */
  normalizeItemName(text) {
    return String(text || '').replace(/[（(].*?[）)]/g, '').trim();
  },

  /** Trim floating-point noise and unnecessary trailing zeros ("4.0" -> "4", "36.4194255433345" -> "36.4"). */
  formatNumber(v, decimals = 1) {
    const n = parseFloat(v);
    if (isNaN(n)) return String(v ?? '');
    const rounded = Math.round(n * Math.pow(10, decimals)) / Math.pow(10, decimals);
    return String(rounded);
  },

  /** Parse a raw "檢測值" cell into {比較關係, 檢測數值} pairs, handling ND / </> / scientific-ish notation. */
  parseValueCell(raw) {
    const s = String(raw ?? '').trim();
    if (s === '') return { cmp: '', val: '' };
    if (/^ND$/i.test(s)) return { cmp: 'ND', val: 'ND' };
    if (/^(NA|未檢測|N\.A\.?)$/i.test(s)) return { cmp: '未檢測', val: '未檢測' };
    let m = s.match(/^([<>])\s*([\d.]+)$/);
    if (m) return { cmp: m[1], val: m[2] };
    // e.g. "6.0×104" meaning 6.0×10^4 (single trailing exponent digit)
    m = s.match(/^([\d.]+)\s*[×xX]\s*10\^?(\d)$/);
    if (m) {
      const val = parseFloat(m[1]) * Math.pow(10, parseInt(m[2], 10));
      return { cmp: '', val: String(val) };
    }
    return { cmp: '', val: s };
  },

  // ---------- report-type detectors & parsers ----------

  /** BN (固定音源噪音-營建工程) / LFN (低頻噪音) single-event report. One sheet -> 1-2 rows. */
  parseNoiseEventSheet(grid) {
    const title = this.cellStr(grid[0]?.[0]) + ' ' + this.cellStr(grid[1]?.[0]);
    const sampleChar = this.labelValue(grid, /樣品特性[:：]/) || '';
    const isLFN = /低頻噪音/.test(title) || /低頻噪音/.test(sampleChar);
    const isBN = !isLFN && (/固定音源噪音|營建工程/.test(title) || /固定音源噪音|營建工程/.test(sampleChar));
    if (!isLFN && !isBN) return null;

    const dateRaw = this.labelValue(grid, /監測日期[:：]/);
    const dateISO = this.rocDateToISO(dateRaw);
    const timeRaw = this.labelValue(grid, /測定時間[:：]/);
    const [tStart, tEnd] = this.splitTimeRange(timeRaw);
    const location = this.labelValue(grid, /監測地點[:：]/) || '';
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
    const zoneText = this.findCell(grid, /噪音管制區/);
    const zone = zoneText ? this.extractZone(this.cellStr(grid[zoneText.r][zoneText.c])) : '';
    const tod = tStart ? this.hourToTod(tStart.split(':')[0]) : '日間';
    const siteCode = this.labelValue(grid, /測點編號[:：]/) || '';

    const baseRow = {
      '日期(起)': dateISO, '時間(起)': tStart, '日期(迄)': dateISO, '時間(迄)': tEnd,
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
    const isVib = /^V-?\d/i.test(sheetName) || !!this.findCell(grid, /Lv日\(Lv10\)=/);
    const isNoise = /^N-?\d/i.test(sheetName) || !!this.findCell(grid, /^L日\(6~20\)/);
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
      const labelHit = this.findCell(grid, /^L日\(6~20\)/);
      if (labelHit) {
        const valueRow = grid[labelHit.r + 1] || [];
        const periods = [
          { key: 'L日', tod: '日間', col: labelHit.c },
          { key: 'L晚', tod: '晚間', col: labelHit.c + 1 },
          { key: 'L夜', tod: '夜間', col: labelHit.c + 2 },
        ];
        periods.forEach(p => {
          const v = this.cellStr(valueRow[p.col]);
          if (v !== '' && !isNaN(parseFloat(v))) {
            rows.push({
              ...baseRow, '管制標準': '噪音管制法第7條第1項', '管制區': '', '環境音量標準': '', '頻率範圍': '20 Hz 至 20kHz',
              '檢測類別': '環境噪音', '監測時段': p.tod, '音源發聲特性': '均能音量(Leq)',
              '監測單位': '16', '監測數值': String(Math.round(parseFloat(v) * 10) / 10), '監測方法': method,
            });
          }
        });
      }
    } else if (isVib) {
      const methodRaw = this.labelValue(grid, /量測方法依據[:：]?/) || this.labelValue(grid, /採樣方法[:：]/);
      const method = this.extractMethodCode(methodRaw) || 'NIEA P204';
      const dayVal = this.labelValue(grid, /Lv日\(Lv10\)=/);
      const nightVal = this.labelValue(grid, /Lv夜\(Lv10\)=/);
      [{ v: dayVal, tod: '日間' }, { v: nightVal, tod: '夜間' }].forEach(p => {
        if (p.v && !isNaN(parseFloat(p.v))) {
          rows.push({
            ...baseRow, '管制標準': '無', '管制區': '無', '環境音量標準': '0', '頻率範圍': '',
            '檢測類別': '振動', '監測時段': p.tod, '音源發聲特性': 'Lvd(10)',
            '監測單位': '159', '監測數值': String(Math.round(parseFloat(p.v) * 10) / 10), '監測方法': method,
          });
        }
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

      const { cmp, val } = this.parseValueCell(valRaw);
      rows.push({
        '日期(起)': dateStart, '時間(起)': '', '日期(迄)': dateEnd, '時間(迄)': '',
        '採樣地點': location, '座標系統': '', '採樣座標-經度 X': '', '採樣座標-緯度 Y': '',
        '場所編號': '', '採樣地點高度(公尺)': '', '污染物採樣高度(公尺)': '', '管制編號': '', '煙道編號': '',
        '檢測類別': '周界空氣品質', '檢測項目': itemName || '落塵量',
        '檢測濃度/質量單位': unitLookup.code, '其他檢測濃度/質量單位': unitLookup.code ? '' : unitText,
        '比較關係': cmp, '檢測數值': /^[\d.]+$/.test(val) ? this.formatNumber(val, 3) : val,
        '檢測方法': method, '檢測機構許可證號': agencyCode, '其他檢測機構名稱': '',
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
    const avgHit = this.findCell(grid, /^日平均值或$/);
    if (!avgHit) return null;
    const avgRow = grid[avgHit.r];

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
    // Deliberately narrow, not a general rule: SO2/O3/etc. can legitimately read
    // "< X" for every single hour when the air is genuinely clean and the detection
    // limit is low — that's a real, common result and must NOT be discarded (that was
    // tried and caused a real false-negative on SO2/O3 in a verified report). CH4 is
    // different: ambient background CH4 is ~1.8–2.0 ppm essentially everywhere on
    // Earth, so a report showing "< 1.0 ppm" for every hour of the day is physically
    // implausible — it means this site's hydrocarbon analyzer (which reports CH4,
    // NMHC, and THC together as one instrument) isn't actually installed, and the
    // report template just carries the three columns with a placeholder value.
    // Confirmed against a verified ground-truth filing where this exact pattern
    // corresponds to CH4/NMHC/THC being correctly absent from the official submission.
    const isImplausiblyLowCH4 = (col) => {
      const vals = [];
      for (let r = unitRow + 1; r < avgHit.r; r++) {
        const v = this.cellStr(grid[r]?.[col]);
        if (v !== '') vals.push(v);
      }
      if (vals.length < 3 || new Set(vals).size !== 1) return false; // not a flat-all-day pattern
      const m = vals[0].match(/^<\s*([\d.]+)$/);
      return !!m && parseFloat(m[1]) < 1.0; // far below real-world CH4 background levels
    };
    let hydrocarbonAnalyzerMissing = false;
    if (cols.CH4 >= 0 && isImplausiblyLowCH4(cols.CH4)) {
      hydrocarbonAnalyzerMissing = true;
      skippedPlaceholderItems.push('CH4', 'NMHC', 'THC');
    }

    // Use the same ND/"< X" parser as the water-table reader — a daily average like
    // "< 0.3" (below detection limit) must never be silently dropped just because
    // parseFloat can't read the "<" prefix; every pollutant column can show this.
    this.AIR_POLLUTANT_DEFS.forEach(def => {
      const col = cols[def.key];
      if (col < 0) return;
      if (hydrocarbonAnalyzerMissing && ['CH4', 'NMHC', 'THC'].includes(def.key)) return;
      const v = this.cellStr(avgRow[col]);
      if (v === '') return;
      const { cmp, val } = this.parseValueCell(v);
      if (val === '' || (cmp === '' && isNaN(parseFloat(val)))) return;
      const methodKey = (def.methodFrom || def.key).toUpperCase();
      rows.push({
        ...baseRow, '檢測項目': def.key, '檢測濃度/質量單位': def.unit,
        '比較關係': cmp, '檢測數值': /^[\d.]+$/.test(val) ? this.formatNumber(val, 3) : val,
        '檢測方法': methodMap[methodKey] || '',
      });
    });
    if (tspVal) {
      const { cmp, val } = this.parseValueCell(tspVal);
      rows.push({
        ...baseRow, '檢測項目': 'TSP', '檢測濃度/質量單位': '127',
        '比較關係': cmp, '檢測數值': /^[\d.]+$/.test(val) ? this.formatNumber(val, 3) : val,
        '檢測方法': methodMap.TSP || '',
      });
    }
    if (cols.pm25 >= 0) {
      const pm25Raw = this.cellStr(grid[unitRow]?.[cols.pm25]);
      if (pm25Raw !== '') {
        const { cmp, val } = this.parseValueCell(pm25Raw);
        if (val !== '' && !(cmp === '' && isNaN(parseFloat(val)))) {
          rows.push({
            ...baseRow, '檢測項目': 'PM2.5', '檢測濃度/質量單位': '127',
            '比較關係': cmp, '檢測數值': /^[\d.]+$/.test(val) ? this.formatNumber(val, 3) : val,
            '檢測方法': pm25Method,
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
    return rows.length ? rows : null;
  },

  /** W (河川/地下水等) / WU (放流水) style vertical water-quality test-item table. */
  parseWaterTableSheet(grid) {
    const hasItemHeader = this.findCell(grid, /檢驗項目/);
    if (!hasItemHeader) return null;

    const location = this.labelValue(grid, /採樣地點[:：]/) || '';
    const sampleTimeRaw = this.labelValue(grid, /採樣時間[:：]/);
    const sampleDateISO = this.rocDateToISO(sampleTimeRaw);
    let sampleTime = '';
    if (sampleTimeRaw) {
      const m = String(sampleTimeRaw).match(/(\d{1,2})時(\d{1,2})分/);
      if (m) sampleTime = `${m[1].padStart(2, '0')}:${m[2].padStart(2, '0')}:00`;
    }
    const sampleChar = this.labelValue(grid, /樣品特性[:：]/) || '';
    const category = this.sampleTypeToCategory(sampleChar);
    const methodRaw = this.labelValue(grid, /採樣方法[:：]/);
    const defaultMethod = this.extractMethodCode(methodRaw);
    const agencyRaw = this.labelValue(grid, /採樣單位[:：]/) || this.labelValue(grid, /公司名稱[:：]/);
    const agencyCode = this.reverseAgencyLookup(agencyRaw);
    const siteCode = this.labelValue(grid, /測點編號[:：]/) || '';

    const headerHit = hasItemHeader;
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
      const itemName = this.normalizeItemName(cells[0]);
      const valueRaw = cells[1];
      let idx = 2;
      const limitRaw = /^--$/.test(cells[idx] || '') ? '' : (cells[idx] || '');
      if (cells[idx] !== undefined) idx++;
      let unitText = '';
      if (cells[idx] && !/^NIEA/i.test(cells[idx])) { unitText = cells[idx]; idx++; }
      let methodText = '';
      if (cells[idx] && /^NIEA/i.test(cells[idx])) { methodText = cells[idx]; idx++; }

      const { cmp, val } = this.parseValueCell(valueRaw);
      // No unit column at all for this item (e.g. pH is dimensionless) means "無" —
      // official unit code 161 — not a blank/unknown unit. Only missing/unmatched
      // unit *text* (unitText present but not found in the code table) is uncertain.
      const unitLookup = unitText ? this.reverseUnitLookup(unitText, itemName) : (this.ITEM_UNIT_OVERRIDES[itemName] ? { code: this.ITEM_UNIT_OVERRIDES[itemName], confident: true } : { code: '161', confident: true });
      const valFormatted = /^[\d.]+$/.test(val) ? this.formatNumber(val, 3) : val;
      const limitFormatted = /^[\d.]+$/.test(limitRaw) ? this.formatNumber(limitRaw, 3) : limitRaw;

      rows.push({
        '日期(起)': sampleDateISO, '時間(起)': sampleTime, '日期(迄)': sampleDateISO, '時間(迄)': sampleTime,
        '採樣地點': location, '座標系統': '', '採樣座標-經度 X': '', '採樣座標-緯度 Y': '',
        '採樣深度(公尺)': '', '採樣水深(公尺)': '', '管制編號': '',
        '檢測類別': category, '檢測項目': itemName,
        '檢測濃度/質量單位': unitLookup.code, '其他檢測濃度/質量單位': unitLookup.code ? '' : unitText,
        '比較關係': cmp, '檢測數值': valFormatted, '檢測極限': limitFormatted,
        '檢測方法': this.extractMethodCode(methodText) || defaultMethod,
        '檢測機構許可證號': agencyCode, '其他檢測機構名稱': '',
        _siteCode: siteCode, _rawLocation: location,
        _uncertainUnit: unitText && !unitLookup.confident,
      });
    }
    return rows.length ? rows : null;
  },

  /** Try every known parser for a sheet, category-scoped. Returns rows[] or null if unrecognized. */
  parseSheet(category, sheetName, grid) {
    if (!grid || grid.length === 0) return null;
    if (category === 'noise') {
      return this.parseNoiseEventSheet(grid) || this.parseNoise24hrSheet(sheetName, grid);
    }
    if (category === 'water') {
      return this.parseWaterTableSheet(grid);
    }
    if (category === 'air') {
      return this.parseAirDustfallSheet(grid) || this.parseAirQualitySheet(grid);
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
