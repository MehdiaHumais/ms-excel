import { CellPrimitive, ErrorValue } from '../model/types';

/**
 * Formula function implementations. Kept independent of the workbook so the
 * same implementations can be re-used by the parser/evaluator and by tests.
 *
 * Functions receive an array of evaluated arguments. Range arguments are
 * passed as `RangeValue`. Any concrete value (number/string/boolean) passes
 * through; errors always propagate unless the function handles them.
 */

export interface RangeValue {
  kind: 'range';
  rows: Array<(CellPrimitive | ErrorValue)[]>;
}

export type ArgValue = CellPrimitive | ErrorValue | RangeValue;

export const ERR_DIV0: ErrorValue = '#DIV/0!';
export const ERR_NAME: ErrorValue = '#NAME?';
export const ERR_VALUE: ErrorValue = '#VALUE!';
export const ERR_REF: ErrorValue = '#REF!';
export const ERR_NUM: ErrorValue = '#NUM!';
export const ERR_NA: ErrorValue = '#N/A';
export const ERR_NULL: ErrorValue = '#NULL!';

export function errorOf(v: unknown): ErrorValue | null {
  if (typeof v === 'string' && v.length > 1 && v[0] === '#' && v.endsWith('!')) return v as ErrorValue;
  if (v === '#N/A') return ERR_NA;
  if (Array.isArray(v)) return null;
  return null;
}

export function isErr(v: unknown): v is ErrorValue {
  return typeof v === 'string' && (v.startsWith('#') && (v.endsWith('!') || v === '#N/A' || v === '#NULL!' || v === '#CIRC!' || v === '#SPILL!'));
}

export function toNumber(v: CellPrimitive | ErrorValue): number | ErrorValue {
  if (isErr(v)) return v;
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (typeof v === 'string' && v.trim() !== '') {
    const n = Number(v.replace(/,/g, ''));
    if (!Number.isNaN(n)) return n;
    // try date-ish
    const d = parseDateish(v);
    if (d !== null) return d;
    return ERR_VALUE;
  }
  return 0;
}

export function parseDateish(v: string): number | null {
  const s = v.trim();
  const iso = /^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T ](\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?$/.exec(s);
  if (iso) {
    const d = Date.UTC(+iso[1], +iso[2] - 1, +iso[3], +(iso[4] ?? 0), +(iso[5] ?? 0), +(iso[6] ?? 0));
    if (!Number.isNaN(d)) return (d - Date.UTC(1899, 11, 30)) / 86_400_000;
    return null;
  }
  const dmy = /^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})$/.exec(s);
  if (dmy) {
    const year = +dmy[3] < 100 ? 2000 + +dmy[3] : +dmy[3];
    const d = new Date(Date.UTC(year, +dmy[2] - 1, +dmy[1]));
    if (!Number.isNaN(d.getTime())) return (d.getTime() - Date.UTC(1899, 11, 30)) / 86_400_000;
  }
  const timeOnly = /^(\d{1,2}):(\d{2})(?::(\d{2}))?(?:\s*(AM|PM))?$/i.exec(s);
  if (timeOnly) {
    let h = +timeOnly[1];
    const m = +timeOnly[2];
    const sec = +(timeOnly[3] ?? 0);
    const ap = timeOnly[4]?.toUpperCase();
    if (ap === 'PM' && h < 12) h += 12;
    if (ap === 'AM' && h === 12) h = 0;
    return h / 24 + m / 1440 + sec / 86400;
  }
  return null;
}

export function toText(v: CellPrimitive | ErrorValue): string | ErrorValue {
  if (isErr(v)) return v;
  if (v === null || v === undefined) return '';
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (typeof v === 'number') {
    if (Number.isInteger(v) && Math.abs(v) < 1e15) return String(v);
    return String(v);
  }
  return v;
}

export function coerceStringToValue(v: string): CellPrimitive {
  const t = v.trim();
  if (t === '') return null;
  if (t === 'TRUE') return true;
  if (t === 'FALSE') return false;
  if (t.includes('%')) {
    const n = Number(t.replace('%', ''));
    if (!Number.isNaN(n)) return n / 100;
  }
  const n = Number(t.replace(/,/g, ''));
  if (!Number.isNaN(n)) return n;
  return v;
}

/** Flatten ranges to a single flat list of concrete (non-range) values. */
export function flatten(args: ArgValue[]): Array<CellPrimitive | ErrorValue> {
  const out: Array<CellPrimitive | ErrorValue> = [];
  for (const a of args) {
    if (a && typeof a === 'object' && 'kind' in a && (a as RangeValue).kind === 'range') {
      for (const row of (a as RangeValue).rows) for (const v of row) out.push(v);
    } else {
      out.push(a as CellPrimitive | ErrorValue);
    }
  }
  return out;
}

export function scalarize(a: ArgValue | undefined): CellPrimitive | ErrorValue {
  if (a === undefined) return null;
  if (a && typeof a === 'object' && 'kind' in a && (a as RangeValue).kind === 'range') {
    const r = a as RangeValue;
    if (r.rows.length === 0 || r.rows[0].length === 0) return null;
    return r.rows[0][0];
  }
  return a as CellPrimitive | ErrorValue;
}

/** Numeric args from args; errors propagate. */
function numericArgs(args: ArgValue[]): number[] | ErrorValue {
  const nums: number[] = [];
  for (const a of args) {
    const s = scalarize(a);
    const n = toNumber(s);
    if (isErr(n)) return n;
    nums.push(n);
  }
  return nums;
}

export type Fn = (args: ArgValue[]) => CellPrimitive | ErrorValue;

function sumFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  let total = 0;
  for (const v of flatten(args)) {
    if (isErr(v)) return v;
    if (typeof v === 'number') total += v;
  }
  return total;
}

function averageFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  const flat = flatten(args);
  let total = 0;
  let count = 0;
  for (const v of flat) {
    if (isErr(v)) return v;
    if (typeof v === 'number') {
      total += v;
      count++;
    }
  }
  if (count === 0) return ERR_DIV0;
  return total / count;
}

function countFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  let count = 0;
  for (const v of flatten(args)) {
    if (isErr(v)) return v;
    if (typeof v === 'number') count++;
  }
  return count;
}

function countaFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  let count = 0;
  for (const v of flatten(args)) {
    if (isErr(v)) return v;
    if (v !== null && v !== undefined && v !== '') count++;
  }
  return count;
}

function countblankFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  let count = 0;
  for (const v of flatten(args)) {
    if (isErr(v)) return v;
    if (v === null || v === undefined || v === '') count++;
  }
  return count;
}

function minFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  let has = false;
  let min = Infinity;
  for (const v of flatten(args)) {
    if (isErr(v)) return v;
    if (typeof v === 'number') {
      has = true;
      if (v < min) min = v;
    }
  }
  return has ? min : 0;
}

function maxFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  let has = false;
  let max = -Infinity;
  for (const v of flatten(args)) {
    if (isErr(v)) return v;
    if (typeof v === 'number') {
      has = true;
      if (v > max) max = v;
    }
  }
  return has ? max : 0;
}

function productFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  let total = 1;
  const flat = flatten(args);
  if (flat.length === 0) return 0;
  for (const v of flat) {
    if (isErr(v)) return v;
    if (typeof v === 'number') total *= v;
  }
  return total;
}

function roundFn(args: ArgValue[], mode: 'round' | 'up' | 'down'): CellPrimitive | ErrorValue {
  const nums = numericArgs(args);
  if (isErr(nums)) return nums;
  const [v, digits = 0] = nums;
  const f = Math.pow(10, digits);
  const r = mode === 'round' ? Math.round(v * f) : mode === 'up' ? Math.ceil(v * f) : Math.floor(v * f);
  return parseFloat((r / f).toPrecision(15));
}

function absFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  const n = toNumber(scalarize(args[0]));
  if (isErr(n)) return n;
  return Math.abs(n);
}

function signFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  const n = toNumber(scalarize(args[0]));
  if (isErr(n)) return n;
  return Math.sign(n);
}

function modFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  const nums = numericArgs(args);
  if (isErr(nums)) return nums;
  const [a, b] = nums;
  return b === 0 ? ERR_DIV0 : a % b;
}

function powerFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  const nums = numericArgs(args);
  if (isErr(nums)) return nums;
  const [a, b] = nums;
  const r = Math.pow(a, b);
  return Number.isNaN(r) ? ERR_NUM : r;
}

function sqrtFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  const n = toNumber(scalarize(args[0]));
  if (isErr(n)) return n;
  return n < 0 ? ERR_NUM : Math.sqrt(n);
}

function intFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  const n = toNumber(scalarize(args[0]));
  if (isErr(n)) return n;
  return Math.floor(n);
}

function ceilingFn(args: ArgValue[], up = true): CellPrimitive | ErrorValue {
  const n = toNumber(scalarize(args[0]));
  if (isErr(n)) return n;
  const sig = toNumber(scalarize(args[1] ?? 1));
  if (isErr(sig)) return sig;
  if (sig === 0) return 0;
  const f = Math[up ? 'ceil' : 'floor'](n / sig);
  return f * sig;
}

function floorFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  return ceilingFn(args, false);
}

function truncFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  const n = toNumber(scalarize(args[0]));
  if (isErr(n)) return n;
  const digits = toNumber(scalarize(args[1] ?? 0));
  if (isErr(digits)) return digits;
  const f = Math.pow(10, digits);
  return Math.trunc(n * f) / f;
}

function evenFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  const n = toNumber(scalarize(args[0]));
  if (isErr(n)) return n;
  return Math.round(n / 2) * 2;
}

function oddFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  const n = toNumber(scalarize(args[0]));
  if (isErr(n)) return n;
  const r = Math.round(n);
  return r % 2 === 0 ? r + 1 : r;
}

function piFn(): CellPrimitive | ErrorValue {
  return Math.PI;
}

function randFn(): CellPrimitive | ErrorValue {
  return Math.random();
}

function concatFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  let out = '';
  for (const a of args) {
    const s = toText(scalarize(a));
    if (isErr(s)) return s;
    out += s;
  }
  return out;
}

function leftFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  const t = toText(scalarize(args[0]));
  if (isErr(t)) return t;
  const n = toNumber(scalarize(args[1] ?? 1));
  if (isErr(n)) return n;
  return t.slice(0, Math.max(0, n));
}

function rightFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  const t = toText(scalarize(args[0]));
  if (isErr(t)) return t;
  const n = toNumber(scalarize(args[1] ?? 1));
  if (isErr(n)) return n;
  return n <= 0 ? '' : t.slice(-Math.min(n, t.length));
}

function midFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  const t = toText(scalarize(args[0]));
  if (isErr(t)) return t;
  const start = toNumber(scalarize(args[1]));
  if (isErr(start)) return start;
  const len = toNumber(scalarize(args[2]));
  if (isErr(len)) return len;
  if (start < 1) return ERR_VALUE;
  return t.slice(start - 1, start - 1 + Math.max(0, len));
}

function lenFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  const t = toText(scalarize(args[0]));
  if (isErr(t)) return t;
  return t.length;
}

function lowerFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  const t = toText(scalarize(args[0]));
  if (isErr(t)) return t;
  return t.toLowerCase();
}

function upperFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  const t = toText(scalarize(args[0]));
  if (isErr(t)) return t;
  return t.toUpperCase();
}

function trimFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  const t = toText(scalarize(args[0]));
  if (isErr(t)) return t;
  return t.replace(/\s+/g, ' ').trim();
}

function properFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  const t = toText(scalarize(args[0]));
  if (isErr(t)) return t;
  return t.toLowerCase().replace(/(^|[^a-z0-9])([a-z])/g, (_m, a, b) => a + b.toUpperCase());
}

function replaceFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  const t = toText(scalarize(args[0]));
  if (isErr(t)) return t;
  const start = toNumber(scalarize(args[1]));
  const n = toNumber(scalarize(args[2]));
  const rep = toText(scalarize(args[3]));
  if (isErr(start)) return start;
  if (isErr(n)) return n;
  if (isErr(rep)) return rep;
  const i = Math.max(0, start - 1);
  return t.slice(0, i) + rep + t.slice(i + Math.max(0, n));
}

function substituteFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  const t = toText(scalarize(args[0]));
  if (isErr(t)) return t;
  const oldT = toText(scalarize(args[1]));
  if (isErr(oldT)) return oldT;
  const newT = toText(scalarize(args[2]));
  if (isErr(newT)) return newT;
  const which = toNumber(scalarize(args[3] ?? 0));
  if (isErr(which)) return which;
  if (which === 0) return t.split(oldT).join(newT);
  let idx = 0;
  let out = '';
  let rest = t;
  let found = -1;
  for (let i = 1; i <= which; i++) {
    found = rest.indexOf(oldT);
    if (found < 0) break;
    out += rest.slice(0, found) + newT;
    rest = rest.slice(found + oldT.length);
    if (i === which) return out + rest;
  }
  void idx;
  return t;
}

function findFn(args: ArgValue[], isSearch = false): CellPrimitive | ErrorValue {
  const needle = toText(scalarize(args[0]));
  if (isErr(needle)) return needle;
  const hay = toText(scalarize(args[1]));
  if (isErr(hay)) return hay;
  const start = toNumber(scalarize(args[2] ?? 1));
  if (isErr(start)) return start;
  const h = isSearch ? hay.toLowerCase() : hay;
  const n = isSearch ? needle.toLowerCase() : needle;
  const idx = h.indexOf(n, Math.max(0, start - 1));
  return idx < 0 ? ERR_VALUE : idx + 1;
}

function searchFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  return findFn(args, true);
}

function exactFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  const a = toText(scalarize(args[0]));
  const b = toText(scalarize(args[1]));
  if (isErr(a)) return a;
  if (isErr(b)) return b;
  return a === b;
}

function repeatFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  const t = toText(scalarize(args[0]));
  if (isErr(t)) return t;
  const n = toNumber(scalarize(args[1]));
  if (isErr(n)) return n;
  return n < 0 ? ERR_VALUE : t.repeat(n);
}

function textFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  const v = scalarize(args[0]);
  const fmt = toText(scalarize(args[1]));
  if (isErr(fmt)) return fmt;
  if (isErr(v)) return v;
  void fmt;
  return toText(v);
}

function valueFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  const t = toText(scalarize(args[0]));
  if (isErr(t)) return t;
  const n = Number(t.replace(/[$%,\s]/g, ''));
  return Number.isNaN(n) ? ERR_VALUE : n;
}

function lenBool(num: number, kind: '>=' | '<=' | '=' | '>' | '<' | '<>'): boolean {
  switch (kind) {
    case '>=':
      return num >= 0;
    case '<=':
      return num <= 0;
    case '=':
      return num === 0;
    case '>':
      return num > 0;
    case '<':
      return num < 0;
    case '<>':
      return num !== 0;
  }
}

function compareFn(args: ArgValue[], kind: '>=' | '<=' | '=' | '>' | '<' | '<>'): CellPrimitive | ErrorValue {
  const a = scalarize(args[0]);
  const b = scalarize(args[1]);
  const an = toNumber(a);
  const bn = toNumber(b);
  if (!isErr(an) && !isErr(bn)) return lenBool(an - bn, kind);
  const at = toText(a);
  const bt = toText(b);
  if (isErr(at)) return at;
  if (isErr(bt)) return bt;
  return lenBool(String(at).localeCompare(String(bt), undefined, { numeric: true }), kind);
}

function notFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  const v = scalarize(args[0]);
  const b = toBool(v);
  if (isErr(b)) return b;
  return !b;
}

function toBool(v: CellPrimitive | ErrorValue): boolean | ErrorValue {
  if (isErr(v)) return v;
  if (typeof v === 'boolean') return v;
  const n = toNumber(v);
  if (isErr(n)) return n;
  return n !== 0;
}

function andFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  for (const v of flatten(args)) {
    const b = toBool(v);
    if (isErr(b)) return b;
    if (!b) return false;
  }
  return true;
}

function orFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  let has = false;
  for (const v of flatten(args)) {
    const b = toBool(v);
    if (isErr(b)) return b;
    if (b) has = true;
  }
  return has;
}

function xorFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  let count = 0;
  for (const v of flatten(args)) {
    const b = toBool(v);
    if (isErr(b)) return b;
    if (b) count++;
  }
  return count % 2 === 1;
}

function sumifFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  const range = args[0] as RangeValue | undefined;
  const crit = toText(scalarize(args[1]));
  const sumRange = args[2] as RangeValue | undefined;
  if (isErr(crit)) return crit;
  const rows = range?.kind === 'range' ? range.rows : [];
  const sumRows = sumRange?.kind === 'range' ? sumRange.rows : rows;
  let total = 0;
  const cm = parseCriterion(String(crit));
  for (let r = 0; r < rows.length; r++) {
    for (let c = 0; c < rows[r].length; c++) {
      const v = rows[r][c];
      let n: number | ErrorValue = toNumber(v);
      if (isErr(n)) {
        if (String(v) === crit) n = 1;
        else continue;
      }
      let word: string = '';
      try {
        word = String(isErr(toText(v)) ? v : toText(v));
      } catch {
        word = '';
      }
      if (matchesCriterion(v, n, word, cm)) {
        const sv = sumRows[r]?.[c];
        const sn = toNumber(sv ?? null);
        if (!isErr(sn)) total += sn;
      }
    }
  }
  return total;
}

function countifFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  const range = args[0] as RangeValue | undefined;
  const crit = toText(scalarize(args[1]));
  if (isErr(crit)) return crit;
  const rows = range?.kind === 'range' ? range.rows : [];
  const cm = parseCriterion(String(crit));
  let count = 0;
  for (const row of rows) {
    for (const v of row) {
      const n = toNumber(v);
      let word = '';
      try {
        word = String(isErr(toText(v)) ? v : toText(v));
      } catch {
        word = '';
      }
      if (matchesCriterion(v, isErr(n) ? null : (n as number), word, cm)) {
        if (typeof v === 'string' && !String(crit).match(/^[0-9<>:=!]/)) {
          const cct = toText(v);
          if (isErr(cct) || cct !== crit) continue;
        }
        count++;
      }
    }
  }
  return count;
}

function averageifFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  const range = args[0] as RangeValue | undefined;
  const crit = toText(scalarize(args[1]));
  if (isErr(crit)) return crit;
  const avgRange = args[2] as RangeValue | undefined;
  const rows = range?.kind === 'range' ? range.rows : [];
  const avgRows = avgRange?.kind === 'range' ? avgRange.rows : rows;
  const cm = parseCriterion(String(crit));
  let total = 0;
  let count = 0;
  for (let r = 0; r < rows.length; r++) {
    for (let c = 0; c < rows[r].length; c++) {
      const v = rows[r][c];
      const n = toNumber(v);
      const word = String(isErr(toText(v)) ? v : toText(v).toString());
      if (matchesCriterion(v, isErr(n) ? null : n, word, cm)) {
        const sv = avgRows[r]?.[c];
        const sn = toNumber(sv ?? null);
        if (!isErr(sn)) {
          total += sn;
          count++;
        }
      }
    }
  }
  return count === 0 ? ERR_DIV0 : total / count;
}

interface Criterion {
  op: '=' | '<>' | '>' | '<' | '>=' | '<=' | '*' | null;
  value: CellPrimitive;
  isText: boolean;
  wildcard: boolean;
}

function parseCriterion(crit: string): Criterion {
  const m = /^(<=|>=|<>|=|>|<)(.*)$/s.exec(crit.trim());
  if (m) {
    return {
      op: m[1] as Criterion['op'],
      value: coerceStringToValue(m[2]),
      isText: /^[^0-9.\-+]/.test(m[2]) && m[2] !== '',
      wildcard: /[*?]/.test(m[2])
    };
  }
  return { op: null, value: coerceStringToValue(crit), isText: /^[^0-9.\-+]/.test(crit), wildcard: /[*?]/.test(crit) };
}

function wildcardMatch(word: string, pattern: string): boolean {
  const re = new RegExp('^' + pattern.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.') + '$', 'i');
  return re.test(word);
}

function matchesCriterion(v: CellPrimitive | ErrorValue, n: number | null, word: string, cm: Criterion): boolean {
  if (isErr(v)) return false;
  const target = cm.value;
  const word2 = typeof target === 'string' ? target : String(target ?? '');
  if (cm.wildcard) return wildcardMatch(String(v), word2);
  if (cm.isText || typeof target === 'string') {
    const s = String(v);
    switch (cm.op) {
      case '=':
        return s === word2;
      case '<>':
        return s !== word2;
      case '>':
        return s.localeCompare(word2) > 0;
      case '<':
        return s.localeCompare(word2) < 0;
      case '>=':
        return s.localeCompare(word2) >= 0;
      case '<=':
        return s.localeCompare(word2) <= 0;
      default:
        return s === word2;
    }
  }
  if (n === null) return false;
  const num2 = typeof target === 'number' ? target : parseFloat(String(target));
  if (Number.isNaN(num2)) return false;
  switch (cm.op) {
    case '>':
      return n > num2;
    case '<':
      return n < num2;
    case '>=':
      return n >= num2;
    case '<=':
      return n <= num2;
    case '<>':
      return n !== num2;
    default:
      return n === num2;
  }
}

function vlookupFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  const lookup = scalarize(args[0]);
  const table = args[1] as RangeValue | undefined;
  const colIdx = toNumber(scalarize(args[2]));
  const approxN = toNumber(scalarize(args[3] ?? true));
  if (isErr(lookup)) return lookup;
  if (isErr(colIdx)) return colIdx;
  const approx = isErr(approxN) ? true : approxN !== 0;
  const rows = table?.kind === 'range' ? table.rows : [];
  const ci = Math.floor(colIdx as number) - 1;
  if (ci < 0 || ci >= (rows[0]?.length ?? 0)) return ERR_REF;
  const lookupNum = toNumber(lookup);
  const lookupIsNum = !isErr(lookupNum);
  if (approx && lookupIsNum) {
    let best: Array<CellPrimitive | ErrorValue> | null = null;
    let bestKey = -Infinity;
    for (const row of rows) {
      const key = toNumber(row[0] ?? null);
      if (isErr(key) || (key as number) > (lookupNum as number)) continue;
      if ((key as number) >= bestKey) {
        bestKey = key as number;
        best = row;
      }
    }
    if (best) return best[ci] ?? null;
    return ERR_NA;
  }
  for (const row of rows) {
    const key = row[0];
    const match =
      typeof key === 'number' || typeof lookup === 'number' || lookupIsNum
        ? String(key) === String(lookup)
        : String(key) === String(lookup);
    if (match) {
      const cell = row[ci];
      return cell === null || cell === undefined ? null : cell;
    }
  }
  return ERR_NA;
}

function hlookupFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  const lookup = scalarize(args[0]);
  const table = args[1] as RangeValue | undefined;
  const rowIdx = toNumber(scalarize(args[2]));
  if (isErr(lookup)) return lookup;
  if (isErr(rowIdx)) return rowIdx;
  const rows = table?.kind === 'range' ? table.rows : [];
  const ri = Math.floor(rowIdx) - 1;
  if (ri < 0 || ri >= rows.length) return ERR_REF;
  const headerRow = rows[0];
  for (let c = 0; c < headerRow.length; c++) {
    const key = headerRow[c];
    const match =
      typeof key === 'number' || typeof lookup === 'number'
        ? Number(key) === Number(lookup)
        : String(key) === String(lookup);
    if (match) {
      const cell = rows[ri]?.[c];
      return cell === null || cell === undefined ? null : cell;
    }
  }
  return ERR_NA;
}

function indexFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  const arr = args[0] as RangeValue | undefined;
  const rowN = toNumber(scalarize(args[1] ?? 0));
  const colN = toNumber(scalarize(args[2] ?? 0));
  if (isErr(rowN)) return rowN;
  if (isErr(colN)) return colN;
  const rows = arr?.kind === 'range' ? arr.rows : [];
  if (rows.length === 0) return ERR_REF;
  const rowCells = rows[Math.max(0, (rowN as number) - 1)];
  if (!rowCells) return ERR_REF;
  const cell = rowCells[Math.max(0, (colN as number) - 1)];
  return cell === null || cell === undefined ? null : cell;
}

function matchFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  const lookup = scalarize(args[0]);
  const arr = args[1] as RangeValue | undefined;
  const typeN = toNumber(scalarize(args[2] ?? 1));
  if (isErr(lookup)) return lookup;
  if (isErr(typeN)) return typeN;
  const type = typeN as number;
  const rows = arr?.kind === 'range' ? arr.rows : [];
  const cells: Array<CellPrimitive | ErrorValue> = [];
  if (rows.length > 0 && rows[0].length === 1) {
    for (const r of rows) cells.push(r[0]);
  } else if (rows.length === 1) {
    cells.push(...rows[0]);
  }
  if (type === 0) {
    for (let i = 0; i < cells.length; i++) {
      if (cells[i] === lookup || String(cells[i]) === String(lookup)) return i + 1;
    }
    return ERR_NA;
  }
  const asc = type > 0;
  for (let i = 0; i < cells.length; i++) {
    const c = toNumber(cells[i] ?? null);
    const l = toNumber(lookup);
    const cok = !isErr(c);
    const lok = !isErr(l);
    if (cok && lok) {
      if (c > l) return asc ? (i === 0 ? ERR_NA : i) : i + 1;
      if (c === l) return i + 1;
    }
  }
  return asc ? cells.length : ERR_NA;
}

function chooseFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  const idx = toNumber(scalarize(args[0]));
  if (isErr(idx)) return idx;
  const i = Math.floor(idx);
  if (i < 1 || i >= args.length) return ERR_VALUE;
  return scalarize(args[i]);
}

function iferrorFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  const v = scalarize(args[0]);
  if (isErr(v)) return scalarize(args[1]);
  return v;
}

function ifnaFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  const v = scalarize(args[0]);
  if (v === ERR_NA) return scalarize(args[1]);
  return v;
}

function isError(args: ArgValue[]): CellPrimitive | ErrorValue {
  const v = scalarize(args[0]);
  return isErr(v);
}
function isErrFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  return scalarize(args[0]) === ERR_NA ? false : isError(args);
}
function isNumberFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  const v = scalarize(args[0]);
  return typeof v === 'number' && !isErr(v);
}
function isTextFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  const v = scalarize(args[0]);
  return typeof v === 'string';
}
function isNonTextFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  const v = scalarize(args[0]);
  return typeof v !== 'string';
}
function isBlankFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  const v = scalarize(args[0]);
  return v === null || v === undefined || v === '';
}
function isLogicalFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  return typeof scalarize(args[0]) === 'boolean';
}

function todayFn(): CellPrimitive | ErrorValue {
  const now = new Date();
  const d = Date.UTC(now.getFullYear(), now.getMonth(), now.getDate());
  return serialFromDateMs(d);
}

function nowFn(): CellPrimitive | ErrorValue {
  return serialFromDateMs(Date.now());
}

function serialFromDateMs(ms: number): number {
  const base = Date.UTC(1899, 11, 30);
  return (ms - base) / 86_400_000;
}

function dateFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  const nums = numericArgs(args);
  if (isErr(nums)) return nums;
  const [y, m, d] = nums;
  const year = Math.floor(y);
  const monthIdx = Math.floor(m) - 1;
  const days = Math.floor(d);
  const firstOfMonth = Date.UTC(year, monthIdx, 1);
  const daysInMonth = new Date(Date.UTC(year, monthIdx + 1, 0)).getUTCDate();
  const target = Math.min(daysInMonth, Math.max(1, days));
  return serialFromDateMs(Date.UTC(year, monthIdx, target));
}

function dateToSerialLocal(d: Date): number {
  return serialFromDateMs(d.getTime());
}

function yearFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  const n = toNumber(scalarize(args[0]));
  if (isErr(n)) return n;
  return new Date(dateToSerialLocalInverse(n)).getUTCFullYear();
}
function monthFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  const n = toNumber(scalarize(args[0]));
  if (isErr(n)) return n;
  return new Date(dateToSerialLocalInverse(n)).getUTCMonth() + 1;
}
function dayFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  const n = toNumber(scalarize(args[0]));
  if (isErr(n)) return n;
  return new Date(dateToSerialLocalInverse(n)).getUTCDate();
}
function hourFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  const n = toNumber(scalarize(args[0]));
  if (isErr(n)) return n;
  return new Date(dateToSerialLocalInverse(n)).getUTCHours();
}
function minuteFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  const n = toNumber(scalarize(args[0]));
  if (isErr(n)) return n;
  return new Date(dateToSerialLocalInverse(n)).getUTCMinutes();
}
function secondFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  const n = toNumber(scalarize(args[0]));
  if (isErr(n)) return n;
  return new Date(dateToSerialLocalInverse(n)).getUTCSeconds();
}
function dateToSerialLocalInverse(serial: number): number {
  const base = Date.UTC(1899, 11, 30);
  return base + serial * 86_400_000;
}
function timeFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  const nums = numericArgs(args);
  if (isErr(nums)) return nums;
  const [h, m, s] = nums;
  const total = ((Math.floor(h) * 3600 + Math.floor(m) * 60 + Math.floor(s)) % 86400 + 86400) % 86400;
  return total / 86400;
}
function weekdayFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  const n = toNumber(scalarize(args[0]));
  if (isErr(n)) return n;
  const d = new Date(dateToSerialLocalInverse(n));
  let w = d.getUTCDay();
  if (w === 0) w = 7;
  return w;
}
function daysFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  const a = toNumber(scalarize(args[0]));
  const b = toNumber(scalarize(args[1]));
  if (isErr(a)) return a;
  if (isErr(b)) return b;
  return (a as number) - (b as number);
}
function datedifFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  const a = toNumber(scalarize(args[0]));
  const b = toNumber(scalarize(args[1]));
  const unit = String(scalarize(args[2]) ?? '').toUpperCase();
  if (isErr(a)) return a;
  if (isErr(b)) return b;
  const d1 = new Date(dateToSerialLocalInverse(a as number));
  const d2 = new Date(dateToSerialLocalInverse(b as number));
  if (d2 < d1) return ERR_NUM;
  if (unit === 'D') return Math.floor((d2.getTime() - d1.getTime()) / 86_400_000);
  let months = (d2.getUTCFullYear() - d1.getUTCFullYear()) * 12 + (d2.getUTCMonth() - d1.getUTCMonth());
  if (months > 0 && d2.getUTCDate() < d1.getUTCDate()) months -= 1;
  if (unit === 'M') return months;
  if (unit === 'Y') return Math.floor(months / 12);
  if (unit === 'YM') return months % 12;
  if (unit === 'MD') {
    const anchor = new Date(Date.UTC(d1.getUTCFullYear(), d1.getUTCMonth() + months, d1.getUTCDate()));
    let extra = Math.floor((d2.getTime() - anchor.getTime()) / 86_400_000);
    if (extra < 0) {
      const prevAnchor = new Date(Date.UTC(d1.getUTCFullYear(), d1.getUTCMonth() + months - 1, d1.getUTCDate()));
      const daysInPrev = new Date(Date.UTC(anchor.getUTCFullYear(), anchor.getUTCMonth(), 0)).getUTCDate();
      extra = Math.floor((d2.getTime() - prevAnchor.getTime()) / 86_400_000) - 0;
      void daysInPrev;
    }
    return Math.max(0, extra);
  }
  if (unit === 'YD') {
    const anchor = new Date(Date.UTC(d1.getUTCFullYear() + Math.floor(months / 12), d1.getUTCMonth(), d1.getUTCDate()));
    let extra = Math.floor((d2.getTime() - anchor.getTime()) / 86_400_000);
    if (extra < 0) {
      const daysInYear = d2.getUTCFullYear() % 4 === 0 && d2.getUTCFullYear() % 100 !== 0 || d2.getUTCFullYear() % 400 === 0 ? 366 : 365;
      extra = Math.floor((d2.getTime() - anchor.getTime()) / 86_400_000) + daysInYear;
    }
    return Math.max(0, extra);
  }
  return ERR_NUM;
}

function isEvenFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  const n = toNumber(scalarize(args[0]));
  if (isErr(n)) return n;
  return (n as number) % 2 === 0;
}
function isOddFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  const n = toNumber(scalarize(args[0]));
  if (isErr(n)) return n;
  return Math.abs((n as number) % 2) === 1;
}
function romanFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  const n = toNumber(scalarize(args[0]));
  if (isErr(n)) return n;
  void n;
  return '';
}

function transposeFn(args: ArgValue[]): CellPrimitive | ErrorValue {
  const r = args[0] as RangeValue | undefined;
  if (!r || r.kind !== 'range') return scalarize(args[0]);
  const out: Array<CellPrimitive | ErrorValue>[] = [];
  const width = r.rows.length;
  const height = Math.max(0, ...r.rows.map((x) => x.length));
  for (let c = 0; c < height; c++) {
    const row: Array<CellPrimitive | ErrorValue> = [];
    for (let rw = 0; rw < width; rw++) row.push(r.rows[rw][c] ?? null);
    out.push(row);
  }
  return out as unknown as CellPrimitive | ErrorValue;
}

/** SUM/AVG/count suite + everything above, registered by name. */
type FnEntry = Fn | { lazy: true; fn: (args: ArgValue[]) => CellPrimitive | ErrorValue };

export const FUNCTIONS: Record<string, FnEntry> = {
  SUM: sumFn,
  SUMIF: sumifFn,
  AVERAGE: averageFn,
  AVERAGEIF: averageifFn,
  COUNT: countFn,
  COUNTA: countaFn,
  COUNTBLANK: countblankFn,
  COUNTIF: countifFn,
  MIN: minFn,
  MAX: maxFn,
  PMIN: minFn,
  PMAX: maxFn,
  PRODUCT: productFn,
  ROUND: (a: ArgValue[]) => roundFn(a, 'round'),
  ROUNDUP: (a: ArgValue[]) => roundFn(a, 'up'),
  ROUNDDOWN: (a: ArgValue[]) => roundFn(a, 'down'),
  MROUND: (a: ArgValue[]) => {
    const n = toNumber(scalarize(a[0]));
    const sig = toNumber(scalarize(a[1]));
    if (isErr(n)) return n;
    if (isErr(sig)) return sig;
    return Math.round((n as number) / (sig as number)) * (sig as number);
  },
  ABS: absFn,
  SIGN: signFn,
  MOD: modFn,
  POWER: powerFn,
  SQRT: sqrtFn,
  INT: intFn,
  CEILING: ceilingFn,
  FLOOR: floorFn,
  TRUNC: truncFn,
  EVEN: evenFn,
  ODD: oddFn,
  PI: piFn,
  RAND: randFn,
  RANDBETWEEN: (a: ArgValue[]) => {
    const nums = numericArgs(a);
    if (isErr(nums)) return nums;
    const [min, max] = nums;
    return Math.floor(min + Math.random() * (max - min + 1));
  },
  CONCAT: concatFn,
  CONCATENATE: concatFn,
  LEFT: leftFn,
  RIGHT: rightFn,
  MID: midFn,
  LEN: lenFn,
  LOWER: lowerFn,
  UPPER: upperFn,
  TRIM: trimFn,
  PROPER: properFn,
  REPLACE: replaceFn,
  SUBSTITUTE: substituteFn,
  FIND: findFn,
  SEARCH: searchFn,
  EXACT: exactFn,
  REPT: repeatFn,
  TEXT: textFn,
  VALUE: valueFn,
  IF: { lazy: true, fn: (args: ArgValue[]) => scalarize(args[0]) ? scalarize(args[1]) : scalarize(args[2]) } as { lazy: true; fn: Fn },
  IFS: { lazy: true, fn: (args: ArgValue[]) => {
      for (let i = 0; i + 1 < args.length; i += 2) {
        if (toBool(scalarize(args[i]))) return scalarize(args[i + 1]);
      }
      return ERR_NA;
    } } as { lazy: true; fn: Fn },
  IFERROR: { lazy: true, fn: (args: ArgValue[]) => iferrorFn(args) } as { lazy: true; fn: Fn },
  IFNA: { lazy: true, fn: (args: ArgValue[]) => ifnaFn(args) } as { lazy: true; fn: Fn },
  AND: andFn,
  OR: orFn,
  XOR: xorFn,
  NOT: notFn,
  TRUE: () => true,
  FALSE: () => false,
  SWITCH: { lazy: true, fn: (args: ArgValue[]) => {
      const expr = scalarize(args[0]);
      for (let i = 1; i + 1 < args.length; i += 2) {
        if (scalarize(args[i]) === expr || String(scalarize(args[i])) === String(expr)) return scalarize(args[i + 1]);
      }
      if (args.length % 2 === 0) return scalarize(args[args.length - 1]);
      return ERR_NA;
    } } as { lazy: true; fn: Fn },
  SUMIFS: (a: ArgValue[]) => {
    if (a.length < 4) return sumFn([a[0] ?? 0]);
    const sumR = a[0] as RangeValue | undefined;
    const sumRows = sumR?.kind === 'range' ? sumR.rows : [];
    const crits: Array<{ rows: Array<Array<CellPrimitive | ErrorValue>>; crit: Criterion }> = [];
    for (let i = 1; i + 1 < a.length; i += 2) {
      const range = a[i] as RangeValue | undefined;
      const crit = parseCriterion(String(toText(scalarize(a[i + 1])) ?? ''));
      crits.push({ rows: range?.kind === 'range' ? range.rows : [], crit });
    }
    let total = 0;
    for (let r = 0; r < sumRows.length; r++) {
      for (let c = 0; c < sumRows[r].length; c++) {
        let ok = true;
        for (const { rows, crit } of crits) {
          const v = rows[r]?.[c] ?? null;
          const n = toNumber(v);
          const word = String(isErr(toText(v)) ? v : toText(v));
          if (!matchesCriterion(v, isErr(n) ? null : n, word, crit)) {
            ok = false;
            break;
          }
        }
        if (ok) {
          const sv = sumRows[r][c];
          const sn = toNumber(sv ?? null);
          if (!isErr(sn)) total += sn;
        }
      }
    }
    return total;
  },
  COUNTIFS: (a: ArgValue[]) => {
    let count = 0;
    const ranges = [];
    for (let i = 0; i + 1 < a.length; i += 2) ranges.push({ range: a[i] as RangeValue, crit: parseCriterion(String(toText(scalarize(a[i + 1])) ?? '')) });
    const rows = ranges[0]?.range?.kind === 'range' ? ranges[0].range.rows : [];
    for (let r = 0; r < rows.length; r++) {
      for (let c = 0; c < rows[r].length; c++) {
        let ok = true;
        for (const { range, crit } of ranges) {
          const rr = range as RangeValue;
          if (rr?.kind === 'range') {
            if (!matchesCriterion(rr.rows[r]?.[c] ?? null, null, String(rr.rows[r]?.[c] ?? ''), crit)) {
              ok = false;
              break;
            }
          }
        }
        if (ok) count++;
      }
    }
    return count;
  },
  VLOOKUP: vlookupFn,
  HLOOKUP: hlookupFn,
  INDEX: indexFn,
  MATCH: matchFn,
  CHOOSE: chooseFn,
  OFFSET: (a: ArgValue[]) => {
    const base = scalarize(a[0]);
    const rowsN = toNumber(scalarize(a[1] ?? 0));
    if (isErr(rowsN)) return rowsN;
    void rowsN;
    return base;
  },
  TODAY: todayFn,
  NOW: nowFn,
  DATE: dateFn,
  TIME: timeFn,
  YEAR: yearFn,
  MONTH: monthFn,
  DAY: dayFn,
  HOUR: hourFn,
  MINUTE: minuteFn,
  SECOND: secondFn,
  WEEKDAY: weekdayFn,
  DAYS: daysFn,
  DATEDIF: datedifFn,
  ISERROR: isError,
  ISERR: isErrFn,
  ISNA: (a: ArgValue[]) => scalarize(a[0]) === ERR_NA,
  ISNUMBER: isNumberFn,
  ISTEXT: isTextFn,
  ISNONTEXT: isNonTextFn,
  ISBLANK: isBlankFn,
  ISLOGICAL: isLogicalFn,
  EVEN_N: evenFn,
  ODD_N: oddFn,
  ROMAN: romanFn,
  TRANSPOSE: transposeFn,
  N: (a: ArgValue[]) => {
    const n = toNumber(scalarize(a[0]));
    if (isErr(n)) return n;
    return n;
  },
  T: (a: ArgValue[]) => {
    const v = scalarize(a[0]);
    return typeof v === 'string' ? v : '';
  },
  SUMPRODUCT: (a: ArgValue[]) => {
    const arr0 = a[0] as RangeValue | undefined;
    const rows = arr0?.kind === 'range' ? arr0.rows : [];
    let total = 0;
    for (let r = 0; r < rows.length; r++) {
      for (let c = 0; c < rows[r].length; c++) {
        let prod = 1;
        for (const arg of a) {
          const rr = arg as RangeValue;
          const v = rr?.kind === 'range' ? rr.rows[r]?.[c] : scalarize(arg);
          const n = toNumber(v ?? null);
          if (isErr(n)) return n;
          prod *= (n as number);
        }
        total += prod;
      }
    }
    return total;
  },
  UNIQUE: (a: ArgValue[]) => {
    const r = a[0] as RangeValue | undefined;
    if (!r || r.kind !== 'range') {
      const v = scalarize(a[0]);
      return v;
    }
    const seen = new Set<string>();
    const out: Array<Array<CellPrimitive | ErrorValue>> = [];
    for (const row of r.rows) {
      const key = row.map((x) => String(x)).join('\u0001');
      if (!seen.has(key)) {
        seen.add(key);
        out.push(row);
      }
    }
    return out as unknown as CellPrimitive | ErrorValue;
  },
  SORT: (a: ArgValue[]) => {
    const r = a[0] as RangeValue | undefined;
    if (!r || r.kind !== 'range') return scalarize(a[0]);
    const colIdx = Math.max(0, (Number(toNumber(scalarize(a[1] ?? 1))) || 1) - 1);
    const asc = scalarize(a[2] ?? true) !== false;
    const rows = r.rows.slice();
    rows.sort((x, y) => {
      const a1 = toNumber(x[colIdx] ?? null);
      const b1 = toNumber(y[colIdx] ?? null);
      const cmp = !isErr(a1) && !isErr(b1) ? (a1 as number) - (b1 as number) : String(x[colIdx]).localeCompare(String(y[colIdx]));
      return asc ? cmp : -cmp;
    });
    return rows as unknown as CellPrimitive | ErrorValue;
  },
  HYPERLINK: (a: ArgValue[]) => scalarize(a[1] ?? a[0]),
  NUMBERVALUE: valueFn,
  CLEAN: (a: ArgValue[]) => {
    const t = toText(scalarize(a[0]));
    if (isErr(t)) return t;
    return String(t).replace(/[\x00-\x1f\x7f]/g, '');
  },
  FIXED: (a: ArgValue[]) => {
    const n = toNumber(scalarize(a[0]));
    if (isErr(n)) return n;
    const digits = Math.max(0, Number(toNumber(scalarize(a[1] ?? 2)) || 0));
    return (n as number).toFixed(digits);
  }
};

export function hasFunction(name: string): boolean {
  return name in FUNCTIONS;
}

export function getFunction(name: string): FnEntry | undefined {
  return FUNCTIONS[name.toUpperCase()];
}