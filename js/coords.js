// coords.js — 座標換算（v4.54，修正清單 #88、#89）
//
// 經緯度（WGS84）⇄ TWD97 二度分帶（TM2），以及十進位（DD）、度分（DMM）、度分秒（DMS）三種寫法。
//
// 參數（使用者 11/4 裁示：中央經線一律東經 121 度，和環境部 wdms 座標轉換網頁相同）：
//   橢球 GRS80：a = 6378137、f = 1/298.257222101（TWD97 用 GRS80；和 WGS84 的差距在公釐以下）
//   中央經線 121°E、尺度比 k0 = 0.9999、東偏移 250000 m、北偏移 0
//
// 準確度：2026/11/4 和環境部網頁逐點比對——
//   120.26947, 23.42211 → X 175345.1938（相同）、Y 2591215.0449（網頁 2591215.0728，差 2.8 公分）
//   121, 23              → X 250000.0000（相同）、Y 2544283.1248（網頁 2544283.1542，差 2.9 公分）
//   118.3186, 24.4366    → X -21947.3942（相同）、Y 2706010.4819（網頁 2706010.5048，差 2.3 公分）
// 申報要的是整數，四捨五入後兩邊相同；只有小數剛好在 .5 附近時可能差 1 公尺。
//
// ⚠️ 金門用 121 度算出來的 X 是負數（例如金門 -21947；馬祖約 145000～200000，是正的 6 位數），不符合「經度 X 為 6 位數字」，
// 這是環境部網頁本身的算法就會出現的結果；轉換工具會提醒，表格的座標檢查也會照常紅框。

const CoordUtil = {
  A: 6378137.0,
  F: 1 / 298.257222101,
  K0: 0.9999,
  LON0: 121,
  FE: 250000,

  /**
   * 臺灣、澎湖、金門、馬祖大致範圍：經度 116～125、緯度 20～27。
   * 官方（座標系統 2）寫的是緯度 20～26，但馬祖南竿約 26.15、東引約 26.37 會超出，
   * 所以這裡放寬到 27（v4.65 #114，使用者 11/11 裁示）。表格的「座標系統 2 緯度 20~26」官方檢查不受影響。
   */
  inTaiwan(lat, lon) { return lat >= 20 && lat <= 27 && lon >= 116 && lon <= 125; },

  _e2() { return this.F * (2 - this.F); },
  _M(phi) {
    const e2 = this._e2();
    return this.A * ((1 - e2 / 4 - 3 * e2 ** 2 / 64 - 5 * e2 ** 3 / 256) * phi
      - (3 * e2 / 8 + 3 * e2 ** 2 / 32 + 45 * e2 ** 3 / 1024) * Math.sin(2 * phi)
      + (15 * e2 ** 2 / 256 + 45 * e2 ** 3 / 1024) * Math.sin(4 * phi)
      - (35 * e2 ** 3 / 3072) * Math.sin(6 * phi));
  },

  /** 經緯度（十進位度）→ TWD97 二度分帶 { x, y }（公尺，未四捨五入）。 */
  toTM2(lat, lon) {
    const e2 = this._e2(), ep2 = e2 / (1 - e2);
    const phi = lat * Math.PI / 180, lam = (lon - this.LON0) * Math.PI / 180;
    const N = this.A / Math.sqrt(1 - e2 * Math.sin(phi) ** 2);
    const T = Math.tan(phi) ** 2, C = ep2 * Math.cos(phi) ** 2, A = lam * Math.cos(phi);
    const x = this.FE + this.K0 * N * (A + (1 - T + C) * A ** 3 / 6 + (5 - 18 * T + T * T + 72 * C - 58 * ep2) * A ** 5 / 120);
    const y = this.K0 * (this._M(phi) + N * Math.tan(phi) * (A * A / 2 + (5 - T + 9 * C + 4 * C * C) * A ** 4 / 24
      + (61 - 58 * T + T * T + 600 * C - 330 * ep2) * A ** 6 / 720));
    return { x, y };
  },

  /** TWD97 二度分帶 → 經緯度 { lat, lon }（十進位度）。 */
  fromTM2(x, y) {
    const e2 = this._e2(), ep2 = e2 / (1 - e2);
    const M = y / this.K0;
    const mu = M / (this.A * (1 - e2 / 4 - 3 * e2 ** 2 / 64 - 5 * e2 ** 3 / 256));
    const e1 = (1 - Math.sqrt(1 - e2)) / (1 + Math.sqrt(1 - e2));
    const phi1 = mu + (3 * e1 / 2 - 27 * e1 ** 3 / 32) * Math.sin(2 * mu)
      + (21 * e1 ** 2 / 16 - 55 * e1 ** 4 / 32) * Math.sin(4 * mu)
      + (151 * e1 ** 3 / 96) * Math.sin(6 * mu) + (1097 * e1 ** 4 / 512) * Math.sin(8 * mu);
    const C1 = ep2 * Math.cos(phi1) ** 2, T1 = Math.tan(phi1) ** 2;
    const N1 = this.A / Math.sqrt(1 - e2 * Math.sin(phi1) ** 2);
    const R1 = this.A * (1 - e2) / Math.pow(1 - e2 * Math.sin(phi1) ** 2, 1.5);
    const D = (x - this.FE) / (N1 * this.K0);
    const lat = phi1 - (N1 * Math.tan(phi1) / R1) * (D * D / 2 - (5 + 3 * T1 + 10 * C1 - 4 * C1 * C1 - 9 * ep2) * D ** 4 / 24
      + (61 + 90 * T1 + 298 * C1 + 45 * T1 * T1 - 252 * ep2 - 3 * C1 * C1) * D ** 6 / 720);
    const lon = (D - (1 + 2 * T1 + C1) * D ** 3 / 6 + (5 - 2 * C1 + 28 * T1 - 3 * C1 * C1 + 8 * ep2 + 24 * T1 * T1) * D ** 5 / 120) / Math.cos(phi1);
    return { lat: lat * 180 / Math.PI, lon: this.LON0 + lon * 180 / Math.PI };
  },

  /** 全形、度分秒符號、方位字母統一。 */
  _clean(s) {
    return String(s ?? '')
      .replace(/[０-９]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0xFEE0))
      .replace(/[．。]/g, '.').replace(/[－−]/g, '-').replace(/[，、﹐﹑]/g, ',').replace(/[；﹔]/g, ';')
      .replace(/[　]/g, ' ').trim();
  },

  /**
   * 一個角度的文字 → { deg, fmt }，讀不懂回傳 null。
   *   十進位   24.502389、N24.502389、24.502389°
   *   度分     24 30.1433、24°30.1433'
   *   度分秒   24 30 08.6、24°30'08.6"、N24°30′8.6″、24度30分8.6秒
   * 方位 S、W 會變成負數。
   */
  parseAngle(text) {
    let s = this._clean(text).toUpperCase();
    if (s === '') return null;
    let sign = 1;
    const hemi = s.match(/^([NSEW])\s*|\s*([NSEW])$/);
    if (hemi) { if (/[SW]/.test(hemi[1] || hemi[2])) sign = -1; s = s.replace(/^[NSEW]\s*|\s*[NSEW]$/, ''); }
    s = s.replace(/[°º度]/g, ' ').replace(/['′’分]/g, ' ').replace(/["″”秒]|''/g, ' ').trim();
    if (s.startsWith('-')) { sign = -sign; s = s.slice(1).trim(); }
    const parts = s.split(/\s+/).filter(Boolean);
    if (parts.length === 0 || parts.length > 3 || !parts.every((p) => /^\d+(\.\d+)?$/.test(p))) return null;
    const [d, m = '0', sec = '0'] = parts;
    if (parts.length >= 2 && (+m >= 60 || d.includes('.'))) return null;
    if (parts.length === 3 && (+sec >= 60 || m.includes('.'))) return null;
    const deg = sign * (+d + +m / 60 + +sec / 3600);
    return { deg, fmt: ['DD', 'DMM', 'DMS'][parts.length - 1] };
  },

  /** 十進位度 → 度分 { d, m }／度分秒 { d, m, s }（正數部分；正負號另外處理）。 */
  toDMM(deg) {
    const a = Math.abs(deg); let d = Math.floor(a); let m = +((a - d) * 60).toFixed(4);
    if (m >= 60) { d += 1; m = 0; }
    return { d, m };
  },
  /** 度分的文字寫法，分固定兩位整數（23°05.5243′；v4.62 #107）。 */
  fmtDMM(v) { return `${v.d}°${v.m.toFixed(4).padStart(7, '0')}′`; },
  toDMS(deg) {
    const a = Math.abs(deg); let d = Math.floor(a); const mf = (a - d) * 60; let m = Math.floor(mf); let s = +((mf - m) * 60).toFixed(2);
    if (s >= 60) { s = 0; m += 1; }
    if (m >= 60) { m = 0; d += 1; }
    return { d, m, s };
  },

  /**
   * 一組座標文字（例如報告上複製來的「24 30 08.6, 120 57 14.8」）→
   * { lat, lon, fmt } 或 { x, y, fmt:'TM2' }；讀不懂回傳 null。
   * 緯度、經度順序寫反（120.9, 24.5）會自動對調。
   */
  parsePair(text) {
    const s = this._clean(text);
    let parts = s.split(/\s*[,;\/]\s*|\s{2,}|\t+/).filter(Boolean);
    if (parts.length !== 2) {
      // 「N24 30 08.6 E120 57 14.8」這種沒有逗號、用方位字母分開的
      const m = s.toUpperCase().match(/^([NS][^EW]+)\s*([EW].+)$/) || s.toUpperCase().match(/^([EW][^NS]+)\s*([NS].+)$/);
      if (m) parts = [m[1], m[2]];
    }
    if (parts.length !== 2) return null;
    const a = parts[0].trim(), b = parts[1].trim();
    // 連寫度分秒（v4.66 #119）：緯度 DDMMSS.ss（整數 6 位）＋經度 DDDMMSS.ss（整數 7 位），順序不拘。
    // 只有 7 位數小於 1,800,000、分秒都小於 60 才算；臺灣的 TWD97 Y 都在 2,420,000 以上，不會被改判。
    const cpt = this._compactPair(a, b);
    if (cpt) return cpt;
    // TWD97：兩個都是 5 位以上的數字；或一個是 7 位的 Y、另一個是 1～6 位（可為負）的 X——
    // 金門東邊用 121 度中央經線算出的 X 可能只有 4 位數（例如 -6558），v4.66 #120
    const tmY = (t) => /^\d{7}(\.\d+)?$/.test(t), tmX = (t) => /^-?\d{1,6}(\.\d+)?$/.test(t);
    const degLike = (t) => /^\d{1,3}\.\d+$/.test(t) && +t <= 180;   // 0～180 帶小數：多半是經緯度度數，不當 X
    if ((/^-?\d{5,7}(\.\d+)?$/.test(a) && /^-?\d{5,7}(\.\d+)?$/.test(b)) || (tmY(a) && tmX(b) && !degLike(b)) || (tmY(b) && tmX(a) && !degLike(a))) {
      let x = +a, y = +b;
      if (Math.abs(x) > 1000000 && Math.abs(y) < 1000000) [x, y] = [y, x];
      return { x, y, fmt: 'TM2' };
    }
    const pa = this.parseAngle(a), pb = this.parseAngle(b);
    if (!pa || !pb) return null;
    if (Math.abs(pa.deg) >= 1000 || Math.abs(pb.deg) >= 1000) return null;   // 不可能是度數（例如 TWD97 多打一碼），交給「看不懂」說明位數
    let lat = pa.deg, lon = pb.deg;
    const upA = this._clean(a).toUpperCase(), upB = this._clean(b).toUpperCase();
    if (/[EW]/.test(upA) || /[NS]/.test(upB) || (Math.abs(lat) > 90 && Math.abs(lon) <= 90)) [lat, lon] = [lon, lat];
    return { lat, lon, fmt: pa.fmt === pb.fmt ? pa.fmt : 'DD' };
  },
  // ───────────── 多筆轉換（v4.61，修正清單 #106） ─────────────

  /** 判讀結果的格式名稱。 */
  FMT_NAME: { DD: '十進位', DMM: '度分', DMS: '度分秒', TM2: 'TWD97' },

  /** 連寫度分秒的一個數 → 十進位度；不合規則回傳 null。intDigits：緯度 6、經度 7。 */
  _compactAngle(t, intDigits) {
    const m = String(t).match(new RegExp(`^(\\d{${intDigits}})(\\.\\d+)?$`));
    if (!m) return null;
    const i = m[1], frac = m[2] || '';
    const d = +i.slice(0, intDigits - 4), mm = +i.slice(-4, -2), ss = +(i.slice(-2) + frac);
    if (mm >= 60 || ss >= 60) return null;
    return d + mm / 60 + ss / 3600;
  },
  _compactPair(a, b) {
    const isLat = (t) => /^\d{6}(\.\d+)?$/.test(t), isLon = (t) => /^\d{7}(\.\d+)?$/.test(t) && +t < 1800000;
    let la, lo;
    if (isLat(a) && isLon(b)) { la = a; lo = b; } else if (isLon(a) && isLat(b)) { la = b; lo = a; } else return null;
    const lat = this._compactAngle(la, 6), lon = this._compactAngle(lo, 7);
    if (lat === null || lon === null || lat > 90 || lon > 180) return null;
    return { lat, lon, fmt: 'DMS' };
  },

  /**
   * 前面帶字樣的寫法（v4.66 #117）：經度／緯度、X／Y、lon／lat… 後面可接 = : ： 或空格。
   * 例：「經度120.954111,緯度24.502389」「X=175345 Y=2591215」「緯度24°30′08.6″ 經度120°57′14.8″」
   * 字樣只用來排順序（緯度＝Y 放前面、經度＝X 放後面），格式一律看數字本身；
   * 字樣寫錯（例如 TWD97 寫成「緯度223037,經度2591215」）時，後面照數字大小自動對調。
   * 回傳「緯度那段, 經度那段」的文字，不是這種寫法就回傳 null。
   */
  _labeledText(text) {
    const s = this._clean(text);
    const re = /(?:^|[\s,;])\s*(經度|緯度|longitude|latitude|long|lng|lon|lat|[XY](?=\s*[=:：]|\s+[-\d]|[-\d]))\s*[=:：]?\s*/gi;
    const hits = [];
    let m;
    while ((m = re.exec(s))) hits.push({ role: /^(緯度|lat|latitude|y)$/i.test(m[1]) ? 'lat' : 'lon', start: m.index, end: re.lastIndex });
    if (hits.length !== 2 || hits[0].role === hits[1].role) return null;
    // 字樣前面不能有別的東西（名稱另外處理）；只允許「TWD97」「WGS84」「TM2」這類座標系統名稱
    if (s.slice(0, hits[0].start).replace(/TWD\s*97|WGS\s*84|TM\s*2|座標|[\s,;:：]/gi, '') !== '') return null;
    const seg = (i) => s.slice(hits[i].end, i === 0 ? hits[1].start : s.length).replace(/^[\s,;]+|[\s,;]+$/g, '');
    const v0 = seg(0), v1 = seg(1);
    if (!/\d/.test(v0) || !/\d/.test(v1)) return null;
    return hits[0].role === 'lat' ? `${v0}, ${v1}` : `${v1}, ${v0}`;
  },

  /** 只有數字、空白隔開的一串（例如「24.5 120.9」「24 30 08.6 120 57 14.8」）→ 對半切成一組；不是就回傳 null。 */
  _pairBySpaces(text) {
    const t = this._clean(text).split(/\s+/).filter(Boolean);
    if (![2, 4, 6].includes(t.length) || !t.every((x) => /^-?\d+(\.\d+)?$/.test(x))) return null;
    const h = t.length / 2;
    return this.parsePair(t.slice(0, h).join(' ') + ', ' + t.slice(h).join(' '));
  },
  _tryPair(text) {
    if (this._clean(text) === '') return null;
    const lab = this._labeledText(text);
    if (lab) return this.parsePair(lab) || this._pairBySpaces(lab.replace(',', ' '));
    return this.parsePair(text) || this._pairBySpaces(text);
  },
  /** 一組判讀結果 → { lat, lon }（TWD97 先換回經緯度），經緯度超出範圍回傳 null。 */
  _latLon(p) {
    const ll = p.fmt === 'TM2' ? this.fromTM2(p.x, p.y) : { lat: p.lat, lon: p.lon };
    return (Math.abs(ll.lat) <= 90 && Math.abs(ll.lon) <= 180) ? ll : null;
  },

  /**
   * 多筆轉換的一行（或 Excel 的一列，各格以 Tab 連起來）→ { name, pair } 或 null。
   * 可以寫：「座標」、「名稱, 座標」、「名稱<Tab>緯度<Tab>經度」、「名稱 座標」。
   * 名稱也可能長得像方位字母（例如 S1、N2），所以「整行都當座標」和「第一段當名稱」兩種都試，
   * 取落在臺灣範圍內的那一種；都不在臺灣時以整行當座標為先。
   */
  parseBatchLine(line) {
    const s = String(line ?? '').replace(/\r/g, '');
    if (s.trim() === '') return null;
    const cands = [];
    const push = (name, text) => { const p = this._tryPair(text); if (p) cands.push({ name: name.trim(), pair: p }); };
    const cells = s.split('\t').map((c) => c.trim()).filter((c) => c !== '');
    if (cells.length >= 2) {
      push('', cells.join(', '));
      push(cells[0], cells.slice(1).join(', '));
      if (cells.length === 2) push(cells[0], cells[1]);
    } else {
      const one = cells[0] || '';
      push('', one);
      const m = one.match(/^(.+?)\s*[,，、﹐﹑;；﹔:：]\s*(.+)$/);
      if (m) push(m[1], m[2]);
      const sp = one.match(/^(\S+)\s+(.+)$/);
      if (sp) push(sp[1], sp[2]);
    }
    if (!cands.length) return null;
    const ok = cands.map((c) => ({ c, ll: this._latLon(c.pair) })).filter((o) => o.ll);
    const tw = ok.find((o) => this.inTaiwan(o.ll.lat, o.ll.lon));
    return (tw || ok[0] || { c: cands[0] }).c;
  },

  /** 看起來像表頭（名稱、緯度、經度、X、Y、座標…）而且讀不出座標的一行。 */
  isHeaderLine(line) {
    // v4.66 #118：只要還有數字（TWD97、WGS84 這類名稱裡的數字不算）就不是表頭，交給判讀、看不懂就照實說
    const t = String(line ?? '');
    if (/\d/.test(t.replace(/TWD\s*97|WGS\s*84|TM\s*2/gi, ''))) return false;
    return !this.parseBatchLine(line) && /名稱|測站|地點|編號|緯度|經度|座標|lat|lon|^\s*x\s*[\t,]|[\t,]\s*y\s*$/i.test(t);
  },

  /**
   * 一組判讀結果 → 各種寫法（和單筆轉換的格式一樣）。TWD97 輸入時 X、Y 用原值，不再繞一圈。
   * 回傳 { lat, lon, x, y, xi, yi, dd, dmm, dms, warn: [] }，經緯度超出範圍時回傳 null。
   */
  describe(p) {
    const ll = this._latLon(p);
    if (!ll) return null;
    const { lat, lon } = ll;
    const tm = p.fmt === 'TM2' ? { x: p.x, y: p.y } : this.toTM2(lat, lon);
    const xi = Math.round(tm.x), yi = Math.round(tm.y);
    const a = this.toDMM(lat), o = this.toDMM(lon), A = this.toDMS(lat), O = this.toDMS(lon);
    const dms = (v) => `${v.d}°${String(v.m).padStart(2, '0')}′${v.s.toFixed(2).padStart(5, '0')}″`;
    const warn = [];
    if (!this.inTaiwan(lat, lon)) warn.push('不在臺灣、澎湖、金門、馬祖範圍內（經度 116～125、緯度 20～27），請確認有沒有寫錯或經緯度寫反');
    if (!/^\d{6}$/.test(String(xi)) || !/^\d{7}$/.test(String(yi))) warn.push(`TWD97 的 X 應為 6 位、Y 應為 7 位數字（這個點是 ${xi}／${yi}）`);
    return {
      lat, lon, x: tm.x, y: tm.y, xi, yi,
      latDD: lat.toFixed(6), lonDD: lon.toFixed(6),
      latDMM: this.fmtDMM(a), lonDMM: this.fmtDMM(o),
      latDMS: dms(A), lonDMS: dms(O), warn,
    };
  },

  /** 看不懂時的說明：有 8 位以上整數的數字，多半是 TWD97 多打一碼（v4.66）。 */
  failHint(text) {
    const big = (this._clean(text).match(/\d+(?:\.\d+)?/g) || []).filter((n) => n.split('.')[0].length >= 8);
    return big.length ? `${big.join('、')} 有 ${big[0].split('.')[0].length} 位整數；TWD97 的 X 應為 6 位、Y 應為 7 位，請確認是不是多打了一碼` : '';
  },

  /** 多行文字或 Excel 各列 → 轉換結果陣列 [{ no, name, input, fmt, r, err }]；空白行略過，第一行是表頭時略過。 */
  convertBatch(lines) {
    const out = [];
    lines.forEach((line, i) => {
      if (String(line ?? '').trim() === '') return;
      if (out.length === 0 && this.isHeaderLine(line)) return;
      const input = String(line).replace(/\t+/g, '　').trim();
      const b = this.parseBatchLine(line);
      const rec = { no: out.length + 1, src: i + 1, name: b ? b.name : '', input, fmt: b ? b.pair.fmt : '', r: null, err: '' };
      if (!b) rec.err = this.failHint(line) || '看不懂這組座標';
      else if (!(rec.r = this.describe(b.pair))) rec.err = '緯度要在 -90～90、經度要在 -180～180 之間';
      out.push(rec);
    });
    return out;
  },
};
