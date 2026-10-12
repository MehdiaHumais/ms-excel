/**
 * Excel-style date/time serial numbers (days since 1899-12-30, with the
 * fictitious leap-day quirk preserved for compatibility with Excel and XLSX).
 */

const EXCEL_EPOCH = Date.UTC(1899, 11, 30);

/** Convert a JS Date to an Excel serial number (time preserved). */
export function dateToSerial(d: Date): number {
  const ms = d.getTime() - EXCEL_EPOCH;
  return ms / 86_400_000;
}

/** Convert an Excel serial number to a JS Date (UTC-based). */
export function serialToDate(serial: number): Date {
  return new Date(EXCEL_EPOCH + serial * 86_400_000);
}

export function isSerialNumber(serial: number): boolean {
  return Number.isFinite(serial);
}

/** Whole-day portion (date) of a serial. */
export function serialDatePart(serial: number): number {
  return Math.floor(serial);
}

/** Fractional portion (time of day) of a serial. */
export function serialTimePart(serial: number): number {
  return serial - Math.floor(serial);
}

const WEEKDAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
const WEEKDAYS_SHORT = WEEKDAYS.map((d) => d.slice(0, 3));
const MONTHS = [
  'January',
  'February',
  'March',
  'April',
  'May',
  'June',
  'July',
  'August',
  'September',
  'October',
  'November',
  'December'
];
const MONTHS_SHORT = MONTHS.map((m) => m.slice(0, 3));

export function getParts(serial: number): {
  year: number;
  month: number;
  day: number;
  weekday: number;
  hour: number;
  minute: number;
  second: number;
} {
  const d = serialToDate(serial);
  return {
    year: d.getUTCFullYear(),
    month: d.getUTCMonth() + 1,
    day: d.getUTCDate(),
    weekday: d.getUTCDay(),
    hour: d.getUTCHours(),
    minute: d.getUTCMinutes(),
    second: d.getUTCSeconds()
  };
}

export function weekdayName(serial: number, short = false): string {
  const { weekday } = getParts(serial);
  return short ? WEEKDAYS_SHORT[weekday] : WEEKDAYS[weekday];
}

export function monthName(serial: number, short = false): string {
  const { month } = getParts(serial);
  return short ? MONTHS_SHORT[month - 1] : MONTHS[month - 1];
}

export function monthNameByNumber(month: number, short = false): string {
  if (short) return MONTHS_SHORT[(month - 1 + 12) % 12];
  return MONTHS[(month - 1 + 12) % 12];
}

export function timedToSerial(hours: number, minutes: number, seconds = 0): number {
  return hours / 24 + minutes / 1440 + seconds / 86400;
}

/** Try to parse common date strings to an Excel serial. Returns null on failure. */
export function parseStringToSerial(text: string): number | null {
  const s = text.trim();
  const iso = /^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T ](\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?$/.exec(s);
  if (iso) {
    const d = Date.UTC(+iso[1], +iso[2] - 1, +iso[3], +(iso[4] ?? 0), +(iso[5] ?? 0), +(iso[6] ?? 0));
    if (!Number.isNaN(d)) return dateToSerial(new Date(d));
  }
  const dmy = /^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})$/.exec(s);
  if (dmy) {
    const year = +dmy[3] < 100 ? 2000 + (+dmy[3]) : +dmy[3];
    const d = new Date(Date.UTC(year, +dmy[2] - 1, +dmy[1]));
    if (!Number.isNaN(d.getTime()) && d.getUTCFullYear() === year) return dateToSerial(d);
  }
  const timeOnly = /^(\d{1,2}):(\d{2})(?::(\d{2}))?(?:\s*(AM|PM))?$/i.exec(s);
  if (timeOnly) {
    let h = +timeOnly[1];
    const m = +timeOnly[2];
    const sec = +(timeOnly[3] ?? 0);
    const ap = timeOnly[4]?.toUpperCase();
    if (ap === 'PM' && h < 12) h += 12;
    if (ap === 'AM' && h === 12) h = 0;
    return timedToSerial(h, m, sec);
  }
  return null;
}