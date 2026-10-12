import { serialDatePart, serialTimePart, getParts, monthNameByNumber, weekdayName } from '../model/serial';

/**
 * A pragmatic number/date format engine supporting the common Excel codes:
 *   General, 0, 0.00, #,##0, #,##0.00, 0%, #,##0.00E+00, $#,##0.00, [$...]...
 *   dates: yyyy, mmmm/mmm/mm/m, dd/d, dddd/ddd, hh/h, mm, ss, AM/PM, [h], [m], [s]
 *   plus semicolon-separated sections: positive;negative;zero;text
 */

export const BUILTIN_FORMATS: Record<string, string> = {
  '0': 'General',
  '1': '0',
  '2': '0.00',
  '3': '#,##0',
  '4': '#,##0.00',
  '9': '0%',
  '10': '0.00%',
  '11': '0.00E+00',
  '12': '# ?/?',
  '13': '# ??/??',
  '14': 'yyyy-mm-dd',
  '15': 'd-mmm-yy',
  '16': 'd-mmm',
  '17': 'mmm-yy',
  '18': 'h:mm AM/PM',
  '19': 'h:mm:ss AM/PM',
  '20': 'h:mm',
  '21': 'h:mm:ss',
  '22': 'm/d/yyyy h:mm',
  '37': '#,##0 ;(#,##0)',
  '38': '#,##0 ;[Red](#,##0)',
  '39': '#,##0.00;(#,##0.00)',
  '40': '#,##0.00;[Red](#,##0.00)',
  '45': 'mm:ss',
  '46': '[h]:mm:ss',
  '47': 'mmss.0',
  '48': '##0.0E+0',
  '49': '@'
};

export function isDateFormat(code: string): boolean {
  return /[ymdhs]/.test(code) && !/E[+-]|#\??\/\??|\*/.test(code);
}

export function isPercentFormat(code: string): boolean {
  return /%/.test(code);
}

export function isTimeOnlyFormat(code: string): boolean {
  return /[hsam]/.test(code) && !/[ymd]/.test(code);
}

type Prim = number | string | boolean | null | undefined;

function roundTo(n: number, digits: number): number {
  const f = Math.pow(10, digits);
  return Math.round(n * f) / f;
}

function generalNumber(n: number): string {
  if (Number.isInteger(n) && Math.abs(n) < 1e15) return String(n);
  if (Math.abs(n) >= 1e15 || (n !== 0 && Math.abs(n) < 1e-9)) {
    return n.toExponential(6).replace(/(\.0+)(?=e)/, '').replace(/e([+-])0+(\d)/, 'e$1$2');
  }
  return String(roundTo(n, 10));
}

function splitSections(code: string): string[] {
  const parts: string[] = [];
  let cur = '';
  let inQuote = false;
  let depth = 0;
  for (let i = 0; i < code.length; i++) {
    const ch = code[i];
    if (ch === '"') inQuote = !inQuote;
    else if (ch === '[' && !inQuote) depth++;
    else if (ch === ']' && !inQuote) depth = Math.max(0, depth - 1);
    if (ch === ';' && !inQuote && depth === 0) {
      parts.push(cur);
      cur = '';
    } else {
      cur += ch;
    }
  }
  parts.push(cur);
  return parts;
}

/** Extract literals outside quotes/brackets and strip format meta. */
function extractPrefix(fmt: string): string {
  let out = '';
  let i = 0;
  while (i < fmt.length) {
    const ch = fmt[i];
    if (ch === '"') {
      let j = i + 1;
      while (j < fmt.length && fmt[j] !== '"') j++;
      out += fmt.slice(i + 1, j);
      i = j + 1;
      continue;
    }
    if (ch === '[') {
      if (fmt.startsWith('[$', i)) {
        let j = i + 2;
        while (j < fmt.length && fmt[j] !== ']') j++;
        let sym = fmt.slice(i + 2, j);
        const dash = sym.indexOf('-');
        if (dash >= 0) sym = sym.slice(0, dash);
        out += sym;
        i = j + 1;
        continue;
      }
      let j = i + 1;
      while (j < fmt.length && fmt[j] !== ']') j++;
      i = j + 1;
      continue;
    }
    if (ch === '\\') {
      if (i + 1 < fmt.length) out += fmt[i + 1];
      i += 2;
      continue;
    }
    if ('0#?.,%E+-()/'.includes(ch) || ch === ' ') break;
    out += ch;
    i++;
  }
  return out;
}

/** Extract the numeric core pattern (first run of 0#? and separators). */
function extractPattern(fmt: string): { pattern: string; before: string; after: string } {
  let core = '';
  let after = '';
  let i = 0;
  while (i < fmt.length) {
    const ch = fmt[i];
    if (ch === ']' || ch === '[') {
      let j = i + 1;
      while (j < fmt.length && fmt[j] !== ']') j++;
      i = j + 1;
      continue;
    }
    if (ch === '"') {
      let j = i + 1;
      while (j < fmt.length && fmt[j] !== '"') j++;
      i = j + 1;
      continue;
    }
    if ('0#?,%.E+-()/'.includes(ch)) {
      while (i < fmt.length) {
        const c2 = fmt[i];
        if (c2 === ']' || c2 === '"') break;
        if (!'0#?,%.E+-()/ '.includes(c2)) {
          break;
        }
        if (c2 !== ' ') core += c2;
        i++;
      }
    } else {
      i++;
    }
  }
  return { pattern: core, before: '', after };
}

interface NumMeta {
  intZeros: number;
  decZeros: number;
  decMax: number;
  group: boolean;
  percent: boolean;
  sci: number;
}

function analyzePattern(pattern: string): NumMeta {
  const sciM = /E[+-]?(\d*)/.exec(pattern);
  const sci = sciM ? (sciM[1] ? parseInt(sciM[1], 10) : 1) : 0;
  const p = pattern.replace(/E[+-]?\d*/gi, '');
  const dot = p.indexOf('.');
  const intPart = dot < 0 ? p : p.slice(0, dot);
  const decPart = dot < 0 ? '' : p.slice(dot + 1);
  return {
    intZeros: (intPart.match(/0/g) ?? []).length,
    decZeros: (decPart.match(/0/g) ?? []).length,
    decMax: decPart.replace(/[0#]/g, '').length + (decPart.match(/[0#]/g) ?? []).length,
    group: intPart.includes(','),
    percent: p.includes('%'),
    sci
  };
}

function groupDigits(s: string): string {
  const neg = s.startsWith('-');
  if (neg) s = s.slice(1);
  const [int, dec] = s.split('.');
  const gi = int.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return (neg ? '-' : '') + (dec !== undefined ? gi + '.' + dec : gi);
}

function formatNumeric(n: number, fmt: string): string {
  const { pattern } = extractPattern(fmt);
  const meta = analyzePattern(pattern);
  void meta.intZeros;
  void meta.decMax;

  let abs = Math.abs(n);
  const neg = n < 0;
  let prefix = extractPrefix(fmt);

  if (meta.percent) abs = abs * 100;
  let mantissa = abs;
  let expSign = '';
  let expPart = '';
  if (meta.sci > 0) {
    if (abs !== 0) {
      const e = Math.floor(Math.log10(abs));
      mantissa = abs / Math.pow(10, e);
      expSign = e < 0 ? '-' : '+';
      const eAbs = Math.abs(e);
      expPart = eAbs < 10 && meta.sci + 1 > 1 && false ? '0' + eAbs : String(eAbs);
      if (meta.sci > 1) expPart = eAbs < 10 ? '0' + eAbs : String(eAbs);
    } else {
      mantissa = 0;
      expPart = '0'.padStart(meta.sci + String(0).length, '0');
    }
  }

  const intDigits = Math.floor(mantissa);
  let decDigits = meta.decZeros;
  let scaled = roundTo(mantissa, decDigits + 2);
  if (meta.sci > 0) {
    decDigits = Math.max(0, meta.sci);
    scaled = roundTo(mantissa, decDigits);
  }
  const intS = String(Math.floor(scaled));
  const fracFull = Math.round((scaled - Math.floor(scaled)) * Math.pow(10, decDigits + 1))
    .toString();
  let frac = Math.round((scaled - Math.floor(scaled)) * Math.pow(10, decDigits))
    .toString()
    .padStart(decDigits, '0');
  void fracFull;

  let out = (meta.group ? groupDigits(intS) : intS);
  if (decDigits > 0) {
    let f = frac;
    // trim trailing zeros if pattern uses # (no forced zeros)
    if (meta.decZeros === 0 && meta.sci === 0) {
      f = f.replace(/0+$/, '');
    }
    if (f.length > 0) out += '.' + f;
  }
  if (meta.percent) out += '%';
  if (meta.sci > 0) {
    let mant = (Math.floor(scaled) + (meta.sci > 1 ? '.' + String(scaled).split('.')[1] : '')) || String(Math.floor(scaled));
    const mantFull = String(scaled);
    mant = Number(mantFull).toFixed(decDigits);
    out = mant + 'E' + expSign + expPart;
  }

  const hasParens = /\(.*\)/.test(fmt) && !/0/.test(fmt.slice(0, fmt.indexOf('(')));
  const paren = hasParens && neg;
  const sign = neg && !paren ? '-' : '';
  return `${prefix}${paren ? '(' : ''}${sign}${out}${paren ? ')' : ''}`;
}

export function formatValue(value: Prim, code: string | undefined): string {
  if (value === null || value === undefined || value === '') return '';
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  const effective = code === undefined || code === '' || code === 'General' ? 'General' : code;

  if (typeof value === 'string') {
    if (effective === 'General' || effective === '@') return value;
    // Text section: splitSections index 3.
    const parts = splitSections(effective);
    if (parts.length === 4 && parts[3]) return parts[3].replace(/@/g, value).replace(/"/g, '');
    return value;
  }

  // number
  if (Number.isNaN(value)) return '###';
  if (effective === 'General') return generalNumber(value);

  const sections = splitSections(effective);
  const sec = value < 0 ? (sections[1] ?? sections[0]) : value === 0 && sections[2] ? sections[2] : sections[0];
  let fmt = sec;

  // percentage & date detection on the effective section
  if (isDateFormat(fmt)) {
    if (/\[h\]|\[m\]|\[s\]/.test(fmt)) return formatElapsed(value, fmt);
    return formatDateSerial(value, fmt);
  }
  if (/^@@?/.test(fmt)) return formatValue(value, '@') as string;
  return formatNumeric(value, fmt);
}

/* ---------------- Date/time formatting ---------------- */

function formatDateSerial(serial: number, fmt: string): string {
  const p = getParts(serial);
  const apStyle = /AM\/PM|A\/P/i.test(fmt);
  const tokens: Array<() => string> = [];
  const push = (fn: () => string) => tokens.push(fn);
  let i = 0;

  const lit = (s: string) => () => s;
  const hAmPm = () => {
    let h = p.hour % 12;
    if (h === 0) h = 12;
    return String(h).padStart(2, '0');
  };

  while (i < fmt.length) {
    const ch = fmt[i];
    let run = 1;
    while (i + run < fmt.length && fmt[i + run] === ch) run++;
    i += run;
    if (ch === '"') {
      let j = i;
      while (j < fmt.length && fmt[j] !== '"') j++;
      push(lit(fmt.slice(i, j)));
      i = j + 1;
      continue;
    }
    if (ch === '[') {
      let j = i + 1;
      while (j < fmt.length && fmt[j] !== ']') j++;
      const inner = fmt.slice(i, j + 1);
      const m = /^\[\$([^\]]*)\]$/.exec(inner);
      if (m) {
        const sym = m[1].split('-')[0];
        push(lit(sym));
      }
      i = j + 1;
      continue;
    }
    if (ch === '\\') {
      push(lit(fmt[i] ?? ''));
      i++;
      continue;
    }
    switch (ch) {
      case 'y':
        if (run === 2) push(() => String(p.year % 100).padStart(2, '0'));
        else push(() => String(p.year).padStart(Math.min(run, 4), '0'));
        break;
      case 'm':
        if (run >= 4) push(() => monthNameByNumber(p.month, false));
        else if (run === 3) push(() => monthNameByNumber(p.month, true));
        else if (run === 2) push(() => String(p.month).padStart(2, '0'));
        else push(() => String(p.month));
        break;
      case 'd':
        if (run >= 4) push(() => weekdayName(serial, false));
        else if (run === 3) push(() => weekdayName(serial, true));
        else if (run === 2) push(() => String(p.day).padStart(2, '0'));
        else push(() => String(p.day));
        break;
      case 'h':
        if (run === 2) push(() => (apStyle ? hAmPm() : String(p.hour).padStart(2, '0')));
        else push(() => (apStyle ? hAmPm() : String(p.hour)));
        break;
      case 's':
        if (run === 2) push(() => String(p.second).padStart(2, '0'));
        else push(() => String(p.second));
        break;
      case 'A':
        if (/^M\/PM/.test(fmt.slice(i))) {
          push(() => (p.hour >= 12 ? 'PM' : 'AM'));
        } else if (fmt.slice(i, i + 2).toUpperCase() === 'M/') {
          push(() => String(p.month).padStart(2, '0'));
        } else {
          push(() => String(p.month).padStart(2, '0'));
        }
        break;
      default: {
        // literal
        let litStr = '';
        let j = i - run;
        // gather the literal run we may have skipped; simplest: emit char
        push(lit(ch));
        break;
      }
    }
  }
  void serialDatePart;
  void serialTimePart;
  return tokens.map((t) => t()).join('');
}

/** Elapsed duration formats like [h]:mm:ss */
function formatElapsed(serial: number, fmt: string): string {
  const totalSeconds = Math.max(0, Math.round(Math.abs(serial) * 86400));
  const h = Math.floor(totalSeconds / 3600);
  const m = Math.floor((totalSeconds % 3600) / 60);
  const s = totalSeconds % 60;
  const unitM = /\[([hms])\]/.exec(fmt);
  const unit = unitM ? unitM[1] : 'h';
  let out = fmt.replace(/\[[hms]\]/, unit === 'h' ? String(h) : unit === 'm' ? String(Math.floor(totalSeconds / 60)) : String(totalSeconds));
  out = out.replace(/:mm/, (unit === 'h' || unit === 'm') ? ':' + String(m).padStart(2, '0') : ':00');
  out = out.replace(/:ss/, ':' + String(s).padStart(2, '0'));
  return out;
}