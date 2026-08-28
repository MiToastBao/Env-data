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
    /*
     * ⚠️ 英文的 am/pm 還要看**它前面是不是別的字的一部分**。
     * 「ppm 08:00」——這個領域最常見的字之一——的結尾正好是 pm，
     * 舊寫法只檢查「pm 後面接空白再接時間」，於是 08:00 被讀成 20:00。
     * 同理 ppm 之外還有 sccm、Nm、ppb…；中文的「下午／上午」不會有這個問題，
     * 所以只有英文那一組要加前置字元的檢查。
     */
    const enPm = /(^|[^0-9A-Za-z])[Pp]\.?[Mm]\.?[\s　]*$/;
    const enAm = /(^|[^0-9A-Za-z])[Aa]\.?[Mm]\.?[\s　]*$/;
    const halfDayAround = (idx, len) => {
      const before = str.slice(Math.max(0, idx - 8), idx);
      const after = str.slice(idx + len, idx + len + 6);
      const pm = /(下午|午後)[\s　]*$/.test(before) || enPm.test(before)
        || /^[\s　]*(下午|午後)/.test(after) || /^[\s　]*[Pp]\.?[Mm]\.?([^0-9A-Za-z]|$)/.test(after);
      const am = /(上午|凌晨|清晨)[\s　]*$/.test(before) || enAm.test(before)
        || /^[\s　]*(上午|凌晨|清晨)/.test(after) || /^[\s　]*[Aa]\.?[Mm]\.?([^0-9A-Za-z]|$)/.test(after);
      return { pm, am };
    };
    const applyHalfDay = (h, idx, len) => {
      const { pm, am } = halfDayAround(idx, len);
      if (pm && h < 12) return h + 12;
      if (am && h === 12) return 0;
      return h;
    };

    /*
     * 12時00分 / 12時30 / 12時。
     *
     * 但「時」也可能是別的詞的頭一個字，那時候前面的數字不是幾點：
     *   ・時間  「檢測項目:PM10 時間:08:00」→ 舊版讀成 10 點
     *   ・時段  「第2時段 08:00」          → 舊版讀成 02:00
     *   ・時制／時程／時候／時區 同理
     * 所以這裡把整組排除掉，不是只排除「間」。
     *
     * ⚠️ 而且**讀不出合法的時／分就往下走**，不能直接 return ''。
     * 舊版一旦這一條命中卻超出範圍就整個放棄，後面的 14:30 那一條根本沒機會跑——
     * 同一格裡先出現一個假的「時」，真正的時間就被吃掉了。
     */
    let m = str.match(/(\d{1,2})\s*時(?!間|段|制|程|候|區|期)\s*(\d{1,2})?\s*分?/);
    if (m) {
      const h = applyHalfDay(parseInt(m[1], 10), m.index, m[0].length);
      const mi = parseInt(m[2] || 0, 10);
      if (h <= 23 && mi <= 59) return `${this.pad(h)}:${this.pad(mi)}:00`;
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

  /*
   * 官方規定：時間限 24 小時制的 00:00~23:59，**勿輸入 24:00**。
   * 五份資料辭典（空品／水質／地質／噪音／生態）都逐字重申這一句。
   *
   * 這裡回傳「為什麼超出範圍」，沒問題就回傳空字串。
   *
   * ⚠️ 只認**看起來就是時鐘值**的字串（HH:MM 或 HH:MM:SS）。
   * 其他讀不懂的內容（例如報告上寫的「連續24小時」）不在這裡處理——
   * 那些本來就會原樣留著讓人看得見要改，硬把它們也標成「時間超出範圍」
   * 只會讓真正該修的 24:00 淹沒在雜訊裡。
   */
  outOfRangeTimeReason(v) {
    const s = String(v ?? '').trim();
    if (s === '') return '';
    let h, mi, sec;
    let m = s.match(/^(\d{1,2})[:：](\d{2})(?:[:：](\d{2}))?$/);
    if (m) {
      h = +m[1]; mi = +m[2]; sec = +(m[3] || 0);
    } else if ((m = s.match(/^(\d{2})(\d{2})(\d{2})$/))) {
      /*
       * 沒有冒號的寫法也要認。_extractTime 收 143000／1430／930 這幾種，
       * 但**只在數字合法時**才收；25:30 打成 `2530`、12:65 打成 `1265` 的時候
       * 它讀不懂，就把原字串原樣留下——而這裡舊版只認冒號，於是那些值
       * 一路過關：沒有紅框、匯出前不清點，就這樣送出去。
       * 打錯的人不會因為少打一個冒號就比較不需要被提醒。
       */
      h = +m[1]; mi = +m[2]; sec = +m[3];
    } else if ((m = s.match(/^(\d{1,2})(\d{2})$/))) {
      h = +m[1]; mi = +m[2]; sec = 0;
    } else {
      return '';
    }
    if (h > 23) return '官方規定時間只能填 00:00~23:59，不可以填 24:00 以上';
    if (mi > 59 || sec > 59) return '分或秒超過 59';
    return '';
  },

  /**
   * 這個值是不是「24:00 整」——唯一一個換成 23:59 有明確道理的寫法。
   * 24:00 就是一天的最後一刻，而 25:30、99:00 不是任何時刻，只是打錯。
   */
  canClampToDayEnd(v) {
    return /^24[:：]00(?:[:：]00)?$/.test(String(v ?? '').trim())
      || /^2400(?:00)?$/.test(String(v ?? '').trim());
  },

  /**
   * 24:00 / 24:00:00 / 2400 → 23:59:00。其餘一律原樣回傳。
   *
   * ⚠️ 舊版是「小時 ≤ 23 就原樣回傳，否則一律回 23:59」，
   * 於是 25:30、99:00 這種明顯打錯的值也會被改成 23:59。
   * 那不是把一天的結尾寫成合法值，是**替使用者挑一個他沒說過的時間**，
   * 而且挑完之後畫面上再也看不出原本打的是什麼——正是 v4.37 要修掉的事。
   * 24:00 有唯一合理的解釋，25:30 沒有，所以只有 24:00 會被換掉。
   */
  clampToDayEnd(v) {
    const s = String(v ?? '').trim();
    return this.canClampToDayEnd(s) ? '23:59:00' : s;
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
