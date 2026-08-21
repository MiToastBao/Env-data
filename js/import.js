// import.js
// Reads an uploaded file (xlsx/xls/csv/pdf), turns it into raw rows of
// {header: value}, then auto-maps headers onto the target category's
// field keys using exact match + alias table + fuzzy match. The user
// confirms/corrects the mapping in the UI before rows are committed.

const ImportEngine = {

  // Common aliases seen in lab-report exports that don't match our exact field names.
  ALIASES: {
    // A report that states one sampling moment writes just "採樣日期"/"採樣時間".
    // Those belong on the (起) fields — without them the REQUIRED 日期(起) came out
    // unmapped and rows committed with a blank sampling date.
    '日期(起)': ['起始日期', '開始日期', '採樣日期(起)', '監測日期(起)', '日期起', '起日期',
      '採樣日期', '監測日期', '調查日期', '檢測日期', '檢驗日期', '取樣日期', '日期'],
    '時間(起)': ['起始時間', '開始時間', '採樣時間(起)', '監測時間(起)', '時間起', '起時間',
      '採樣時間', '監測時間', '調查時間', '檢測時間', '測定時間', '時間'],
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
    '檢測數值': ['監測數值', '數值', '結果', '檢測結果', '分析結果', '測值', '檢測值', '量測值', '監測值'],
    '監測數值': ['檢測數值', '數值', '結果', '測值', '檢測值', '量測值', '監測值'],
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
  async readFile(file, preferredKeys = []) {
    const name = file.name.toLowerCase();
    if (name.endsWith('.xlsx') || name.endsWith('.xls') || name.endsWith('.csv')) {
      return this._readSheet(file, preferredKeys);
    } else if (name.endsWith('.pdf')) {
      return this._readPdf(file);
    }
    throw new Error('不支援的檔案格式，請上傳 .xlsx / .xls / .csv 或 .pdf');
  },

  cellStr(v) { return v === undefined || v === null ? '' : String(v).trim(); },

  /** An Excel serial number straight from the file -> "YYYY-MM-DD", "HH:MM:SS", or
   *  "YYYY-MM-DD HH:MM:SS". Pure arithmetic, so the result is identical in every
   *  timezone — this is the path date cells take now, and the reason the reader asks
   *  SheetJS for the workbook a second time with cellDates:false. */
  _excelSerialToString(serial) {
    const p = DateTimeUtil.parseAny(serial);
    if (!p) return String(serial);
    if (p.date && p.time) return `${p.date} ${p.time}`;
    return p.date || p.time || String(serial);
  },

  /** Fallback for a Date object when no serial is available (e.g. a CSV, where
   *  SheetJS parses the text itself). Reads the Date through LOCAL getters inside
   *  DateTimeUtil — never getUTC*, which was the original cause of the "10點變成2點"
   *  8-hour shift and the off-by-one day. */
  _excelDateObjToString(d) {
    const p = DateTimeUtil.parseAny(d);
    if (!p) return '';
    if (p.date && p.time) return `${p.date} ${p.time}`;
    return p.date || p.time || '';
  },

  /**
   * Reads a workbook into { headers, rows } for the column-mapping importer.
   *
   * Unlike the old version, the header row does NOT have to be the first row of the
   * sheet: real report exports routinely carry a title, a logo row, or a couple of
   * blank spacer rows above the actual column names, and assuming row 1 meant those
   * files read as "0 usable columns". `_detectHeaderRow` scores each of the first
   * rows for how much it looks like a header (short label-ish cells, recognizable
   * field names, real data underneath) and uses the best one.
   *
   * `preferredKeys` (the target category's field keys) are passed in so a workbook
   * that already matches the official template — e.g. a previous season's completed
   * filing — locks onto the right sheet and the right row instead of whichever sheet
   * merely happens to have the most rows.
   */
  async _readSheet(file, preferredKeys = []) {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: 'array', cellDates: true });
    // the same workbook parsed with cellDates:false — see gridSerial below
    let wbNoDates = null;
    try { wbNoDates = XLSX.read(buf, { type: 'array', cellDates: false }); } catch (e) { wbNoDates = null; }
    let best = null;
    for (const sheetName of wb.SheetNames) {
      const ws = wb.Sheets[sheetName];
      const wsNoDates = wbNoDates ? wbNoDates.Sheets[sheetName] : null;
      // Two passes: raw:false gives human-readable display text for ordinary cells
      // (numbers with their formatting, plain text); raw:true is needed to get the
      // actual Date object back for real date/time cells rather than a formatted
      // string this app might not be able to parse (see _excelDateObjToString).
      const range = this._contentRange(ws);
      const optsD = { header: 1, defval: '', raw: false };
      const optsR = { header: 1, defval: '', raw: true };
      if (range) { optsD.range = range; optsR.range = range; }
      const gridDisplay = XLSX.utils.sheet_to_json(ws, optsD);
      const gridRaw = XLSX.utils.sheet_to_json(ws, optsR);
      // A THIRD read, of the same sheet parsed WITHOUT cellDates, so date cells come
      // back as their raw Excel serial number instead of a JS Date. SheetJS builds
      // those Date objects by calling the Date constructor with local calendar parts,
      // and in a zone whose historical offset is not a whole number of minutes
      // (UTC+5:30 India, +5:45 Nepal, +3:30 Iran) the result lands seconds before
      // midnight on the PREVIOUS day. Reading the serial and doing the arithmetic
      // ourselves (DateTimeUtil.serialToParts) has no timezone in the path at all.
      const gridSerial = wsNoDates ? XLSX.utils.sheet_to_json(wsNoDates, optsR) : null;
      if (gridDisplay.length === 0) continue;
      const detected = this._detectHeaderRow(gridDisplay, preferredKeys);
      if (!detected) continue;
      const built = this._buildRowsFromGrid(gridDisplay, gridRaw, detected.rowIndex, gridSerial);
      if (built.rows.length === 0) continue;
      const score = detected.score * 1000 + Math.min(built.rows.length, 999);
      if (!best || score > best.score) {
        best = { sheetName, score, headers: built.headers, rows: built.rows, headerRowIndex: detected.rowIndex };
      }
    }
    if (!best) return { headers: [], rows: [], sheetNames: wb.SheetNames };
    return {
      headers: best.headers, rows: best.rows, sheetNames: wb.SheetNames,
      usedSheet: best.sheetName, headerRowIndex: best.headerRowIndex,
    };
  },

  /** True for a cell that plausibly IS a column heading: short, has text, isn't a
   *  "label：value" form field, isn't a bare number. */
  _looksLikeHeaderCell(v) {
    const s = this.cellStr(v).replace(/\s+/g, ' ');
    if (s === '' || s.length > 30) return false;
    if (/[:：]\s*\S/.test(s)) return false;   // "採樣地點：荷包嶼大排" is a form field, not a header
    if (/^-?[\d.,%]+$/.test(s)) return false; // a bare number is data, not a heading
    return true;
  },

  /** Scores each of the first rows of a grid and returns the most header-like one. */
  _detectHeaderRow(grid, preferredKeys = []) {
    const preferred = new Set(preferredKeys.map(k => this.normalize(k)));
    const limit = Math.min(grid.length, 40);
    let best = null;
    for (let r = 0; r < limit; r++) {
      const row = grid[r] || [];
      let headerish = 0, known = 0;
      const seen = new Set();
      row.forEach(cell => {
        if (!this._looksLikeHeaderCell(cell)) return;
        const n = this.normalize(cell);
        if (n === '' || seen.has(n)) return;
        seen.add(n);
        headerish++;
        if (preferred.has(n)) known += 3;
        else if (typeof AutoDetect !== 'undefined' && AutoDetect.matchHeader(cell)) known += 1;
      });
      if (headerish < 2) continue;
      // there has to be at least one row of actual content underneath
      let dataRows = 0;
      for (let rr = r + 1; rr < grid.length && dataRows < 2; rr++) {
        if ((grid[rr] || []).some(c => this.cellStr(c) !== '')) dataRows++;
      }
      if (dataRows === 0) continue;
      const score = known * 4 + headerish - r * 0.1; // earlier rows win ties
      if (!best || score > best.score) best = { rowIndex: r, score };
    }
    return best;
  },

  /** Turns a grid + a header row index into {headers, rows} of {header: value}. */
  _buildRowsFromGrid(gridDisplay, gridRaw, headerIdx, gridSerial) {
    const headerRow = gridDisplay[headerIdx] || [];
    // How far right the sheet actually has content — a column whose heading cell is
    // blank still gets a slot, because a merged heading spanning two columns (採樣時間
    // over 起/迄) and unlabeled 單位/序號 columns are both normal, and dropping them
    // meant their data could never be mapped to a field at all.
    let lastCol = -1;
    for (let r = headerIdx; r < gridDisplay.length; r++) {
      const row = gridDisplay[r] || [];
      for (let c = row.length - 1; c > lastCol; c--) { if (this.cellStr(row[c]) !== '') { lastCol = c; break; } }
    }
    const cols = [];
    for (let i = 0; i <= lastCol; i++) {
      let name = this.cellStr(headerRow[i]).replace(/\s+/g, ' ');
      if (name === '') {
        // is there anything below this column worth offering? if not, skip it
        let hasData = false;
        for (let r = headerIdx + 1; r < gridDisplay.length && !hasData; r++) {
          if (this.cellStr((gridDisplay[r] || [])[i]) !== '') hasData = true;
        }
        if (!hasData) continue;
        name = `（未命名欄位 ${i + 1}）`;
      }
      let unique = name, n = 2;
      while (cols.some(c => c.name === unique)) unique = `${name} (${n++})`;
      cols.push({ index: i, name: unique });
    }
    if (cols.length === 0) return { headers: [], rows: [] };
    const rows = [];
    for (let r = headerIdx + 1; r < gridDisplay.length; r++) {
      const dRow = gridDisplay[r] || [];
      const rRow = gridRaw[r] || [];
      const sRow = gridSerial ? (gridSerial[r] || []) : null;
      const out = {};
      let any = false;
      cols.forEach(c => {
        const rawVal = rRow[c.index];
        const display = this.cellStr(dRow[c.index]);
        let v;
        if (rawVal instanceof Date) {
          // prefer the untouched serial number over the (timezone-sensitive) Date
          const serial = sRow ? sRow[c.index] : undefined;
          v = (typeof serial === 'number')
            ? this._excelSerialToString(serial)
            : this._excelDateObjToString(rawVal);
        } else if (typeof rawVal === 'number' && !/%\s*$/.test(display)) {
          // A cell's DISPLAYED text is whatever its number format rounds it to —
          // a coordinate stored as 120.26323411 can show as "120.26323". For a
          // filing, the stored value is the truthful one, so when the two disagree
          // numerically the underlying number wins. When they agree, the display
          // text is kept (it may carry thousands separators or trailing zeros the
          // person expects to see). Percent-formatted cells keep their display,
          // since their underlying value is a different quantity entirely.
          // Thousands separators must not survive into a numeric filing field —
          // "12,000" is text to the receiving system. Keep the display only when it
          // adds nothing but formatting we can safely drop.
          const dispNum = parseFloat(display.replace(/,/g, ''));
          v = (isFinite(dispNum) && Math.abs(dispNum - rawVal) < 1e-9)
            ? display.replace(/,/g, '')
            : String(rawVal);
        } else {
          v = display;
        }
        out[c.name] = v;
        if (v !== '') any = true;
      });
      if (any) rows.push(out);
    }
    return { headers: cols.map(c => c.name), rows };
  },

  /** How much of a category's own field list appears verbatim in a set of headers —
   *  used to recognize "this file is already in the official template format"
   *  (a completed filing, or a blank template someone filled in by hand), which must
   *  go straight to the column-mapping importer instead of a report-form parser. */
  gridSchemaMatchRatio(grid, categoryFields) {
    const norm = new Set(categoryFields.map(f => this.normalize(f.key)));
    let best = 0;
    for (let r = 0; r < Math.min(grid.length, 15); r++) {
      const row = grid[r] || [];
      const hits = new Set();
      row.forEach(c => { const n = this.normalize(c); if (n && norm.has(n)) hits.add(n); });
      best = Math.max(best, hits.size / categoryFields.length);
    }
    return best;
  },

  schemaMatchRatio(headers, categoryFields) {
    if (!headers || headers.length === 0) return 0;
    const norm = new Set(headers.map(h => this.normalize(h)));
    const hits = categoryFields.filter(f => norm.has(this.normalize(f.key))).length;
    return hits / categoryFields.length;
  },

  /**
   * Read every sheet of an xlsx/xls workbook as a raw grid (array of rows,
   * blanks preserved as ''), for the smart form-parsers in smartparse.js
   * which need to scan label/value positions rather than a header row.
   *
   * Also attaches a `_hiddenCols` Set (0-indexed column numbers) to each grid.
   * This matters for reports like the 24hr air-quality table: a lab may reuse
   * one shared spreadsheet template across sites that don't all monitor the
   * same pollutants, and simply hide the columns for items that site doesn't
   * measure — the cells still contain placeholder text, but the column itself
   * is marked hidden. `cellStyles: true` is required for SheetJS to populate
   * `!cols[i].hidden` at all; without it the hidden flag is silently dropped.
   */
  /**
   * The range a sheet's REAL content occupies, as a SheetJS range string, or null.
   *
   * These report workbooks declare absurd used ranges — one 97-row noise sheet
   * reports A1:WWQ97, i.e. 16,163 columns, because stray formatting reaches out
   * that far. sheet_to_json then materializes 1.57M array slots per sheet, and a
   * 16-sheet workbook spent ~5 seconds building padding before a single value was
   * read. Narrowing to the cells that actually hold a value first makes the whole
   * import roughly an order of magnitude faster, with identical output: column
   * indices still start at 0, so hidden-column flags and every parser's column
   * arithmetic are unaffected.
   */
  _contentRange(ws) {
    if (!ws || !ws['!ref']) return null;
    let maxR = -1, maxC = -1;
    for (const key of Object.keys(ws)) {
      if (key[0] === '!') continue;
      const cell = ws[key];
      if (!cell || cell.v === undefined || cell.v === null || cell.v === '') continue;
      const a = XLSX.utils.decode_cell(key);
      if (a.r > maxR) maxR = a.r;
      if (a.c > maxC) maxC = a.c;
    }
    if (maxR < 0 || maxC < 0) return null;
    const declared = XLSX.utils.decode_range(ws['!ref']);
    if (declared.e.c <= maxC && declared.e.r <= maxR) return null; // already tight
    return XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: maxR, c: maxC } });
  },

  async readWorkbookGrids(file) {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: 'array', cellDates: false, cellStyles: true });
    const grids = {};
    wb.SheetNames.forEach(name => {
      const ws = wb.Sheets[name];
      const range = this._contentRange(ws);
      const opts = { header: 1, defval: '', raw: true };
      if (range) opts.range = range;
      const grid = XLSX.utils.sheet_to_json(ws, opts);
      const hiddenCols = new Set();
      if (ws['!cols']) {
        ws['!cols'].forEach((c, i) => { if (c && c.hidden) hiddenCols.add(i); });
      }
      grid._hiddenCols = hiddenCols;
      grids[name] = grid;
    });
    return grids;
  },

  /**
   * Extract per-page text lines from a PDF, each line reconstructed by grouping
   * text fragments at roughly the same vertical position and ordering left-to-right.
   * Returns an array of pages, each an array of line strings — shared by both the
   * generic table reader and readPdfAsGrids below.
   */
  async _readPdfPages(file) {
    const buf = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: buf }).promise;
    const pages = [];
    for (let i = 1; i <= pdf.numPages; i++) {
      const page = await pdf.getPage(i);
      const content = await page.getTextContent();
      const rowsByY = {};
      content.items.forEach(item => {
        const y = Math.round(item.transform[5]);
        if (!rowsByY[y]) rowsByY[y] = [];
        rowsByY[y].push(item);
      });
      const ys = Object.keys(rowsByY).map(Number).sort((a, b) => b - a);
      const lines = [];
      ys.forEach(y => {
        const items = rowsByY[y].sort((a, b) => a.transform[4] - b.transform[4]);
        // Reconstruct spacing from the actual gap between text items, not just a flat
        // single space — otherwise a "label： value" pair sitting close together on
        // the page collapses into one un-splittable string, and downstream code that
        // splits columns on runs of 2+ spaces (both here and in readPdfAsGrids) can
        // never see the label and value as separate cells.
        let line = '';
        let prevEnd = null;
        items.forEach(it => {
          const x = it.transform[4];
          const fontSize = Math.abs(it.transform[3]) || Math.abs(it.transform[0]) || 10;
          const spaceUnit = fontSize * 0.5; // rough width of one space character at this size
          if (prevEnd !== null && spaceUnit > 0) {
            const gap = x - prevEnd;
            if (gap > spaceUnit * 1.2) {
              // a real gap on the page: treat as a column break (always 2+ spaces so
              // "split on 2+ spaces" logic downstream picks it up as a boundary)
              line += ' '.repeat(Math.min(Math.max(Math.round(gap / spaceUnit), 2), 20));
            } else if (gap > 0) {
              line += ' ';
            }
          }
          line += it.str;
          prevEnd = x + (typeof it.width === 'number' ? it.width : it.str.length * fontSize * 0.5);
        });
        line = line.trim();
        if (line) lines.push(line);
      });
      pages.push(lines);
    }
    return pages;
  },

  /**
   * Turn a PDF into per-page "grids" (array-of-arrays of cell strings) compatible
   * with the smart form-parsers in smartparse.js, so a combined multi-category PDF
   * report can potentially be auto-detected page by page the same way a multi-sheet
   * Excel workbook is. This is inherently less reliable than the Excel path: PDF text
   * extraction only approximates column positions from character spacing, so parsers
   * that depend on values sitting in a specific column (not just "the next non-empty
   * cell after a label") can misalign. Rows recovered this way should get extra
   * scrutiny — the item-selection and site-profile review steps still apply, but
   * treat PDF-sourced values as a first draft rather than a verified transcription.
   */
  async readPdfAsGrids(file) {
    const pages = await this._readPdfPages(file);
    const grids = {};
    pages.forEach((lines, i) => {
      // split each line into cells on runs of 2+ spaces (mirrors the generic PDF
      // table reader), padded so every row in the page has the same column count
      const splitLine = (l) => l.split(/\s{2,}|\t/).map(s => s.trim());
      const rows = lines.map(splitLine);
      const maxCols = rows.reduce((m, r) => Math.max(m, r.length), 0);
      const padded = rows.map(r => { const out = r.slice(); while (out.length < maxCols) out.push(''); return out; });
      grids[`page${i + 1}`] = padded;
    });
    return grids;
  },

  async _readPdf(file) {
    const pages = await this._readPdfPages(file);
    const allLines = pages.flat();
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

  /** Everything that lands in a schema field goes through here, so a date field can
   *  only ever end up holding a date and a time field can only ever end up holding a
   *  time — regardless of whether the source cell was a real Excel date, an Excel
   *  serial number, a 民國 string like "115年06月25日12時00分", or free text. */
  _coerceValue(val, type) {
    if (type === 'date') return DateTimeUtil.toISODate(val);
    if (type === 'time') return DateTimeUtil.toHMS(val);
    if (val instanceof Date) return this._excelDateObjToString(val);
    return String(val ?? '').trim();
  },
};
