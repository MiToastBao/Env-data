// import.js
// Reads an uploaded file (xlsx/xls/csv/pdf), turns it into raw rows of
// {header: value}, then auto-maps headers onto the target category's
// field keys using exact match + alias table + fuzzy match. The user
// confirms/corrects the mapping in the UI before rows are committed.

const ImportEngine = {

  // Common aliases seen in lab-report exports that don't match our exact field names.
  ALIASES: {
    '日期(起)': ['起始日期', '開始日期', '採樣日期(起)', '監測日期(起)', '日期起', '起日期'],
    '時間(起)': ['起始時間', '開始時間', '採樣時間(起)', '監測時間(起)', '時間起', '起時間'],
    '日期(迄)': ['結束日期', '採樣日期(迄)', '監測日期(迄)', '日期迄', '迄日期'],
    '時間(迄)': ['結束時間', '採樣時間(迄)', '監測時間(迄)', '時間迄', '迄時間'],
    '採樣地點': ['監測地點', '測點', '測站', '採樣點', '地點', '測點名稱'],
    '監測地點': ['採樣地點', '測點', '測站', '地點', '測點名稱'],
    '調查地點': ['採樣地點', '調查點', '測點'],
    '座標系統': ['坐標系統', '座標系', 'coordinate system'],
    '採樣座標-經度 X': ['經度', 'X座標', 'x', '座標X', '經度X', 'TWD97 X', 'WGS84經度'],
    '採樣座標-緯度 Y': ['緯度', 'Y座標', 'y', '座標Y', '緯度Y', 'TWD97 Y', 'WGS84緯度'],
    '檢測類別': ['監測類別', '類別', '檢測類型'],
    '檢測項目': ['監測項目', '項目', '測項', '分析項目'],
    '檢測濃度/質量單位': ['單位', '濃度單位', '檢測單位', '單位代碼'],
    '監測單位': ['單位', '檢測單位', '單位代碼'],
    '比較關係': ['比較符號', '符號'],
    '檢測數值': ['監測數值', '數值', '結果', '檢測結果', '分析結果'],
    '監測數值': ['檢測數值', '數值', '結果'],
    '檢測極限': ['偵測極限', 'MDL', '定量極限', 'RL'],
    '檢測方法': ['分析方法', '方法', '檢驗方法'],
    '檢測機構許可證號': ['檢測機構', '機構代碼', '檢驗機構', '委託檢測機構', '機構許可證號'],
    '其他檢測機構名稱': ['檢測機構名稱', '機構名稱'],
    '備註': ['說明', 'remark', 'note'],
    '管制編號': ['EMS編號', '管制編號(EMS)'],
    '學名': ['scientific name'],
    '中文名': ['俗名', '中文俗名', '物種名稱'],
    '數量': ['個體數', '數目'],
    '保育分類': ['保育等級', '保育類別'],
  },

  normalize(s) {
    return String(s || '').trim().toLowerCase().replace(/[\s\u3000（）()\-_/]/g, '');
  },

  /** Read a File object and return { headers: [...], rows: [{header: value}, ...] } */
  async readFile(file) {
    const name = file.name.toLowerCase();
    if (name.endsWith('.xlsx') || name.endsWith('.xls') || name.endsWith('.csv')) {
      return this._readSheet(file);
    } else if (name.endsWith('.pdf')) {
      return this._readPdf(file);
    }
    throw new Error('不支援的檔案格式，請上傳 .xlsx / .xls / .csv 或 .pdf');
  },

  async _readSheet(file) {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: 'array', cellDates: true });
    // Prefer a sheet that looks like a data table (has >1 row); default to first sheet
    let best = null;
    for (const sheetName of wb.SheetNames) {
      const ws = wb.Sheets[sheetName];
      const json = XLSX.utils.sheet_to_json(ws, { defval: '', raw: false });
      if (json.length > 0 && (!best || json.length > best.rows.length)) {
        best = { sheetName, rows: json };
      }
    }
    if (!best) return { headers: [], rows: [], sheetNames: wb.SheetNames };
    const headers = Object.keys(best.rows[0]);
    return { headers, rows: best.rows, sheetNames: wb.SheetNames, usedSheet: best.sheetName };
  },

  /**
   * Read every sheet of an xlsx/xls workbook as a raw grid (array of rows,
   * blanks preserved as ''), for the smart form-parsers in smartparse.js
   * which need to scan label/value positions rather than a header row.
   */
  async readWorkbookGrids(file) {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: 'array', cellDates: false });
    const grids = {};
    wb.SheetNames.forEach(name => {
      grids[name] = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: '', raw: true });
    });
    return grids;
  },

  async _readPdf(file) {
    const buf = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: buf }).promise;
    let allLines = [];
    for (let i = 1; i <= pdf.numPages; i++) {
      const page = await pdf.getPage(i);
      const content = await page.getTextContent();
      // group text items by approximate Y position to reconstruct rows
      const rowsByY = {};
      content.items.forEach(item => {
        const y = Math.round(item.transform[5]);
        if (!rowsByY[y]) rowsByY[y] = [];
        rowsByY[y].push(item);
      });
      const ys = Object.keys(rowsByY).map(Number).sort((a, b) => b - a);
      ys.forEach(y => {
        const items = rowsByY[y].sort((a, b) => a.transform[4] - b.transform[4]);
        const line = items.map(it => it.str).join(' ').trim();
        if (line) allLines.push(line);
      });
    }
    // Best-effort: try to find a header-like line (many short tokens) then
    // split subsequent lines by 2+ spaces / tabs into columns.
    if (allLines.length === 0) return { headers: [], rows: [], pdfLines: [] };
    const splitLine = (l) => l.split(/\s{2,}|\t/).map(s => s.trim()).filter(Boolean);
    let headerIdx = 0;
    let headerCols = splitLine(allLines[0]);
    for (let i = 0; i < Math.min(allLines.length, 5); i++) {
      const cols = splitLine(allLines[i]);
      if (cols.length > headerCols.length) { headerCols = cols; headerIdx = i; }
    }
    const rows = [];
    for (let i = headerIdx + 1; i < allLines.length; i++) {
      const cols = splitLine(allLines[i]);
      if (cols.length === 0) continue;
      const row = {};
      headerCols.forEach((h, idx) => { row[h] = cols[idx] !== undefined ? cols[idx] : ''; });
      rows.push(row);
    }
    return { headers: headerCols, rows, pdfLines: allLines, isPdfBestEffort: true };
  },

  /** Suggest a mapping from source headers to target field keys for a category */
  suggestMapping(sourceHeaders, categoryFields) {
    const mapping = {}; // fieldKey -> sourceHeader | null
    const usedHeaders = new Set();
    categoryFields.forEach(f => {
      let match = sourceHeaders.find(h => !usedHeaders.has(h) && this.normalize(h) === this.normalize(f.key));
      if (!match) {
        match = sourceHeaders.find(h => !usedHeaders.has(h) && this.normalize(h) === this.normalize(f.label));
      }
      if (!match && this.ALIASES[f.key]) {
        for (const alias of this.ALIASES[f.key]) {
          match = sourceHeaders.find(h => !usedHeaders.has(h) && this.normalize(h) === this.normalize(alias));
          if (match) break;
        }
      }
      if (!match) {
        // loose contains-match as last resort
        match = sourceHeaders.find(h => !usedHeaders.has(h) &&
          (this.normalize(h).includes(this.normalize(f.key)) || this.normalize(f.key).includes(this.normalize(h))) &&
          this.normalize(h).length > 1);
      }
      mapping[f.key] = match || null;
      if (match) usedHeaders.add(match);
    });
    return mapping;
  },

  /** Apply a confirmed mapping to raw rows, producing schema-shaped row objects */
  applyMapping(rawRows, mapping, categoryFields) {
    return rawRows.map(raw => {
      const row = {};
      categoryFields.forEach(f => {
        const src = mapping[f.key];
        let val = src ? raw[src] : '';
        if (val === undefined || val === null) val = '';
        row[f.key] = this._coerceValue(val, f.type);
      });
      return row;
    });
  },

  _coerceValue(val, type) {
    if (val instanceof Date) {
      if (type === 'time') return this._fmtTime(val);
      return this._fmtDate(val);
    }
    return String(val).trim();
  },

  _fmtDate(d) {
    const pad = n => String(n).padStart(2, '0');
    return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  },
  _fmtTime(d) {
    const pad = n => String(n).padStart(2, '0');
    return `${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
  },
};
