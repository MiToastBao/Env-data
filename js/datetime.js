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
      if (raw > 0 && raw < 1) { // pure fraction of a day = a time-only value
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
      if ((n > 0 && n < 1) || (n >= this.SERIAL_MIN && n <= this.SERIAL_MAX)) return this.parseAny(n);
    }

    const date = this._extractDate(s);
    const time = this._extractTime(s);
    if (!date && !time) return null;
    return { date: date || '', time: time || '' };
  },

  /** Pulls a calendar date out of free text, understanding both 西元 and 民國 years. */
  _extractDate(s) {
    let m;
    // 西元: 2026-06-25 / 2026/6/25 / 2026.06.25 / 2026年6月25日
    m = s.match(/(\d{4})\s*[-/.年]\s*(\d{1,2})\s*[-/.月]\s*(\d{1,2})/);
    if (m) return `${m[1]}-${this.pad(m[2])}-${this.pad(m[3])}`;
    // 民國: 115年06月25日 / 115.06.25 / 115/6/25 (year 1..200 only, so it can't
    // swallow a 西元 date — those were already handled above)
    m = s.match(/(?:^|[^\d])(\d{1,3})\s*[-/.年]\s*(\d{1,2})\s*[-/.月]\s*(\d{1,2})/);
    if (m) {
      const ry = parseInt(m[1], 10);
      if (ry >= 1 && ry <= 200) return `${ry + 1911}-${this.pad(m[2])}-${this.pad(m[3])}`;
    }
    // compact 西元 8 digits: 20260625
    m = s.match(/^(\d{4})(\d{2})(\d{2})$/);
    if (m && +m[2] >= 1 && +m[2] <= 12 && +m[3] >= 1 && +m[3] <= 31) return `${m[1]}-${m[2]}-${m[3]}`;
    // compact 民國 7 digits: 1150625
    m = s.match(/^(\d{3})(\d{2})(\d{2})$/);
    if (m && +m[2] >= 1 && +m[2] <= 12 && +m[3] >= 1 && +m[3] <= 31) return `${+m[1] + 1911}-${m[2]}-${m[3]}`;
    return '';
  },

  /** Pulls a clock time out of free text. */
  _extractTime(s) {
    let m;
    // 12時00分 / 12時
    m = s.match(/(\d{1,2})\s*時\s*(\d{1,2})?\s*分?/);
    if (m) return `${this.pad(m[1])}:${this.pad(m[2] || 0)}:00`;
    // 14:30 / 14:30:15 / 2:30 PM
    m = s.match(/(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM|am|pm)?/);
    if (m) {
      let h = parseInt(m[1], 10);
      const ap = (m[4] || '').toUpperCase();
      if (ap === 'PM' && h < 12) h += 12;
      if (ap === 'AM' && h === 12) h = 0;
      if (h <= 23) return `${this.pad(h)}:${m[2]}:${m[3] || '00'}`;
    }
    // 143000 / 1430 / 930 typed free-hand (only when the whole string is that number)
    m = s.match(/^(\d{2})(\d{2})(\d{2})$/);
    if (m && +m[1] <= 23 && +m[2] <= 59) return `${m[1]}:${m[2]}:${m[3]}`;
    m = s.match(/^(\d{1,2})(\d{2})$/);
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
