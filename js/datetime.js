// datetime.js
// One single place that decides how a date or a time is read, stored and shown.
//
// WHY THIS FILE EXISTS
// --------------------
// Dates and times used to be converted in three or four different places, each
// with its own rules, and one of them (the Excel-cell reader in import.js) read a
// JavaScript Date object through its **UTC** getters. SheetJS builds those Date
// objects in **local** time, so in Taiwan (UTC+8) every imported value came out
// shifted by 8 hours: a 10:00 sampling time was stored as 02:00, and a plain date
// with no time part slid backwards onto the previous day. That is exactly the
// "日期多或少一天、時間少了一大段" symptom this module fixes.
//
// THE RULES, IN ONE PLACE
// -----------------------
//  - Internal storage / export format for a date is ISO "YYYY-MM-DD".
//  - Internal storage / export format for a time is "HH:MM:00" — seconds are
//    deliberately dropped, because the official filing template's own time format
//    is "h:mm" and no lab report reports sampling times to the second.
//  - What the person SEES is "YYYY/MM/DD" for a date and "HH:MM" for a time.
//  - Anything that can't be understood is returned unchanged rather than silently
//    thrown away, so a wrong value stays visible and fixable instead of vanishing.

const DateTimeUtil = {
  pad(n) { return String(n).padStart(2, '0'); },

  /** Lowest/highest plausible Excel date serial numbers (1950-01-01 .. 2099-12-31),
   *  used only when a cell holds a bare number where a date was expected. */
  SERIAL_MIN: 18264,
  SERIAL_MAX: 73050,

  /** A JS Date -> parts, always read through LOCAL getters (never getUTC*), because
   *  that is how SheetJS constructs the Date in the first place. */
  _parts(d) {
    return {
      y: d.getFullYear(), mo: d.getMonth() + 1, d: d.getDate(),
      h: d.getHours(), mi: d.getMinutes(), s: d.getSeconds(),
    };
  },

  /** An Excel serial number (days since 1899-12-30, fraction = time of day) -> parts.
   *  Pure arithmetic — no Date object, therefore no timezone can interfere. */
  serialToParts(serial) {
    const n = Number(serial);
    if (!isFinite(n)) return null;
    const days = Math.floor(n);
    // round to the nearest whole second so 0.7708333333 comes out as 18:30:00
    // rather than 18:29:59 through floating-point truncation
    let secs = Math.round((n - days) * 86400);
    let dayShift = 0;
    if (secs >= 86400) { secs -= 86400; dayShift = 1; }
    // Excel day 0 is 1899-12-30 under its (Lotus-compatible) 1900 date system.
    const ms = Date.UTC(1899, 11, 30) + (days + dayShift) * 86400000;
    const dt = new Date(ms);
    return {
      y: dt.getUTCFullYear(), mo: dt.getUTCMonth() + 1, d: dt.getUTCDate(),
      h: Math.floor(secs / 3600), mi: Math.floor((secs % 3600) / 60), s: secs % 60,
    };
  },

  /** Reads whatever a cell/field can plausibly hold and returns { date, time } where
   *  date is "YYYY-MM-DD" (or '') and time is "HH:MM:SS" (or ''). Returns null when
   *  the input isn't recognizable as a date/time at all. */
  parseAny(raw) {
    if (raw === null || raw === undefined || raw === '') return null;
    if (raw instanceof Date) {
      if (isNaN(raw.getTime())) return null;
      const p = this._parts(raw);
      // A time-only Excel cell still arrives as a Date whose date part is just
      // Excel's epoch (1899/1900) and carries no meaning — read it as a time.
      if (p.y <= 1900) return { date: '', time: `${this.pad(p.h)}:${this.pad(p.mi)}:${this.pad(p.s)}` };
      const date = `${p.y}-${this.pad(p.mo)}-${this.pad(p.d)}`;
      const time = (p.h || p.mi || p.s) ? `${this.pad(p.h)}:${this.pad(p.mi)}:${this.pad(p.s)}` : '';
      return { date, time };
    }

    if (typeof raw === 'number') {
      if (raw >= 0 && raw <= 1) { // pure fraction of a day = a time-only value
        const p = this.serialToParts(raw);
        return { date: '', time: `${this.pad(p.h)}:${this.pad(p.mi)}:${this.pad(p.s)}` };
      }
      if (raw >= this.SERIAL_MIN && raw <= this.SERIAL_MAX) {
        const p = this.serialToParts(raw);
        const date = `${p.y}-${this.pad(p.mo)}-${this.pad(p.d)}`;
        const time = (p.h || p.mi || p.s) ? `${this.pad(p.h)}:${this.pad(p.mi)}:${this.pad(p.s)}` : '';
        return { date, time };
      }
      return null;
    }

    const s = String(raw).trim();
    if (s === '') return null;

    // A bare number that arrived as text (some exports quote everything)
    if (/^\d+(\.\d+)?$/.test(s) && !/^\d{6,8}$/.test(s)) {
      const n = parseFloat(s);
      if ((n >= 0 && n <= 1) || (n >= this.SERIAL_MIN && n <= this.SERIAL_MAX)) return this.parseAny(n);
    }

    const date = this._extractDate(s);
    const time = this._extractTime(s);
    if (!date && !time) return null;
    return { date: date || '', time: time || '' };
  },

  /** Days in a given month, leap years included. */
  daysInMonth(y, mo) {
    return [31, (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0 ? 29 : 28,
      31, 30, 31, 30, 31, 31, 30, 31, 30, 31][mo - 1];
  },

  /** A y/m/d triple -> "YYYY-MM-DD", or '' when it isn't a real calendar date.
   *  Rejecting 2026-02-30 here matters: without it the value reaches the exporter,
   *  which builds an Excel serial with Date.UTC and quietly rolls it over to
   *  2026-03-02 — a DIFFERENT real date, filed without anyone seeing it change. */
  _ymd(y, mo, d) {
    if (!(y >= 1911 && y <= 2200)) return '';
    if (!(mo >= 1 && mo <= 12)) return '';
    if (!(d >= 1 && d <= this.daysInMonth(y, mo))) return '';
    return `${y}-${this.pad(mo)}-${this.pad(d)}`;
  },

  /** Pulls a calendar date out of free text, understanding both 西元 and 民國 years.
   *
   *  Every branch validates the calendar and returns '' rather than a rolled-over
   *  date. The 民國 branch is also anchored: it used to match ANY "n.n.n" run
   *  anywhere inside a sentence, which turned real report text into dates — the pH
   *  range "6.0-9.0" became 1917-00-09, and a sample id "0606525-W26-01.1" became
   *  1937-01-01. It now only fires when the number really looks like a date, i.e.
   *  written with 年/月/日, or standing on its own rather than embedded in a longer
   *  token. */
  _extractDate(s) {
    let m;
    // 西元 with 年月日 markers, anywhere in the text
    m = s.match(/(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日?/);
    if (m) { const r = this._ymd(+m[1], +m[2], +m[3]); if (r) return r; }
    // 民國 with 年月日 markers, anywhere in the text
    m = s.match(/(\d{1,3})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日?/);
    if (m) { const ry = +m[1]; if (ry >= 1 && ry <= 200) { const r = this._ymd(ry + 1911, +m[2], +m[3]); if (r) return r; } }
    // 西元 with separators — must not be glued to more digits/letters on either side
    m = s.match(/(?:^|[^0-9A-Za-z.\-/])(\d{4})\s*[-/.]\s*(\d{1,2})\s*[-/.]\s*(\d{1,2})(?![0-9])/);
    if (m) { const r = this._ymd(+m[1], +m[2], +m[3]); if (r) return r; }
    // 民國 with separators — same anchoring
    m = s.match(/(?:^|[^0-9A-Za-z.\-/])(\d{2,3})\s*[-/.]\s*(\d{1,2})\s*[-/.]\s*(\d{1,2})(?![0-9])/);
    if (m) { const ry = +m[1]; if (ry >= 1 && ry <= 200) { const r = this._ymd(ry + 1911, +m[2], +m[3]); if (r) return r; } }
    // compact 西元 8 digits: 20260625
    m = s.match(/^(\d{4})(\d{2})(\d{2})$/);
    if (m) { const r = this._ymd(+m[1], +m[2], +m[3]); if (r) return r; }
    // compact 民國 7 digits: 1150625
    m = s.match(/^(\d{3})(\d{2})(\d{2})$/);
    if (m) { const r = this._ymd(+m[1] + 1911, +m[2], +m[3]); if (r) return r; }
    return '';
  },

  /** Pulls a clock time out of free text.
   *
   *  Handles the Chinese half-day markers as well as AM/PM. Excel's own zh-TW
   *  date-time format renders as "2026/3/15 下午 02:30"; reading that as 02:30
   *  filed every afternoon measurement twelve hours early. */
  _extractTime(s) {
    const str = String(s);
    // Whether a half-day marker sits IMMEDIATELY beside the clock value that matched.
    // Testing the whole string meant "PM10 08:00" — the commonest token in this
    // entire domain — was read as 20:00 because "PM" appeared somewhere in the cell.
    const halfDayAround = (idx, len) => {
      const before = str.slice(Math.max(0, idx - 8), idx);
      const after = str.slice(idx + len, idx + len + 6);
      const pm = /(下午|午後|[Pp]\.?[Mm]\.?)[\s　]*$/.test(before) || /^[\s　]*(下午|午後|[Pp]\.?[Mm]\.?)/.test(after);
      const am = /(上午|凌晨|清晨|[Aa]\.?[Mm]\.?)[\s　]*$/.test(before) || /^[\s　]*(上午|凌晨|清晨|[Aa]\.?[Mm]\.?)/.test(after);
      return { pm, am };
    };
    const applyHalfDay = (h, idx, len) => {
      const { pm, am } = halfDayAround(idx, len);
      if (pm && h < 12) return h + 12;
      if (am && h === 12) return 0;
      return h;
    };

    // 12時00分 / 12時30 / 12時 — but never the 時 of 時間, which is a caption, not
    // an hour ("檢測項目:PM10 時間:08:00" was being read as 10 o'clock).
    let m = str.match(/(\d{1,2})\s*時(?!間)\s*(\d{1,2})?\s*分?/);
    if (m) {
      const h = applyHalfDay(parseInt(m[1], 10), m.index, m[0].length);
      const mi = parseInt(m[2] || 0, 10);
      if (h <= 23 && mi <= 59) return `${this.pad(h)}:${this.pad(mi)}:00`;
      return '';
    }
    // 14:30 / 14:30:15 / 2:30 PM / 下午 02:30
    m = str.match(/(\d{1,2})[:：](\d{2})(?:[:：](\d{2}))?/);
    if (m) {
      const h = applyHalfDay(parseInt(m[1], 10), m.index, m[0].length);
      if (h <= 23 && +m[2] <= 59 && +(m[3] || 0) <= 59) return `${this.pad(h)}:${m[2]}:${m[3] || '00'}`;
    }
    // 143000 / 1430 / 930 typed free-hand (only when the whole string is that number)
    m = str.match(/^(\d{2})(\d{2})(\d{2})$/);
    if (m && +m[1] <= 23 && +m[2] <= 59) return `${m[1]}:${m[2]}:${m[3]}`;
    m = str.match(/^(\d{1,2})(\d{2})$/);
    if (m && +m[1] <= 23 && +m[2] <= 59) return `${this.pad(m[1])}:${m[2]}:00`;
    return '';
  },

  // ---------- the four functions the rest of the app actually calls ----------

  /** Anything -> canonical stored date "YYYY-MM-DD". Unrecognized input is returned
   *  as-is (trimmed) so the person can see and correct it. */
  toISODate(raw) {
    if (raw === null || raw === undefined) return '';
    const parsed = this.parseAny(raw);
    if (parsed && parsed.date) return parsed.date;
    return String(raw).trim();
  },

  /** Anything -> canonical stored time "HH:MM:00". Seconds are intentionally
   *  dropped (the official template shows h:mm and reports never give seconds). */
  toHMS(raw) {
    if (raw === null || raw === undefined) return '';
    const parsed = this.parseAny(raw);
    if (parsed && parsed.time) return parsed.time.slice(0, 5) + ':00';
    // a value that is only a date can't be a time — don't leave a date string
    // sitting in a time field, but don't invent 00:00 either
    if (parsed && parsed.date && !parsed.time) return '';
    return String(raw).trim();
  },

  /** Stored "YYYY-MM-DD" -> what the person sees, "YYYY/MM/DD". */
  toDisplayDate(v) {
    const s = String(v ?? '').trim();
    if (s === '') return '';
    const m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    return m ? `${m[1]}/${m[2]}/${m[3]}` : s;
  },

  /** Stored "HH:MM:SS" -> what the person sees, "HH:MM" (never any seconds). */
  toDisplayTime(v) {
    const s = String(v ?? '').trim();
    if (s === '') return '';
    const m = s.match(/^(\d{1,2}):(\d{2})/);
    return m ? `${this.pad(m[1])}:${m[2]}` : s;
  },
};
