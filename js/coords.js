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
// ⚠️ 金門、馬祖用 121 度算出來的 X 是負數（例如金門 -21947），不符合「經度 X 為 6 位數字」，
// 這是環境部網頁本身的算法就會出現的結果；轉換工具會提醒，表格的座標檢查也會照常紅框。

const CoordUtil = {
  A: 6378137.0,
  F: 1 / 298.257222101,
  K0: 0.9999,
  LON0: 121,
  FE: 250000,

  /** 臺灣、澎湖、金門、馬祖大致範圍（官方：經度 116～125、緯度 20～26）。 */
  inTaiwan(lat, lon) { return lat >= 20 && lat <= 26 && lon >= 116 && lon <= 125; },

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
      .replace(/[．。]/g, '.').replace(/[－−]/g, '-').replace(/[，、]/g, ',')
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
    // TWD97：兩個都是 5 位以上的數字
    if (/^-?\d{5,7}(\.\d+)?$/.test(a) && /^-?\d{5,7}(\.\d+)?$/.test(b)) {
      let x = +a, y = +b;
      if (Math.abs(x) > 1000000 && Math.abs(y) < 1000000) [x, y] = [y, x];
      return { x, y, fmt: 'TM2' };
    }
    const pa = this.parseAngle(a), pb = this.parseAngle(b);
    if (!pa || !pb) return null;
    let lat = pa.deg, lon = pb.deg;
    const upA = this._clean(a).toUpperCase(), upB = this._clean(b).toUpperCase();
    if (/[EW]/.test(upA) || /[NS]/.test(upB) || (Math.abs(lat) > 90 && Math.abs(lon) <= 90)) [lat, lon] = [lon, lat];
    return { lat, lon, fmt: pa.fmt === pb.fmt ? pa.fmt : 'DD' };
  },
};
