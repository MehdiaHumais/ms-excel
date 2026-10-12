import { CellPrimitive, ErrorValue } from '../model/types';
import { getFunction, ArgValue, RangeValue, scalarize, toNumber, toText } from './functions';

/**
 * Formula engine core: tokenizer -> parser -> evaluator.
 *
 * The engine is deliberately decoupled from the workbook model through a
 * narrow `FormulaSource` interface, so tests can run against fakes and the
 * UI can bind the live workbook.
 */

export interface FormulaSource {
  /** Read a single cell. 1-based row/col. */
  getCell(sheet: string, row: number, col: number): CellPrimitive | ErrorValue | null;
  /** Resolve sheet id from name. */
  resolveSheet(name: string): string | null;
  /** Resolve a named range to a concrete sheet+range. */
  resolveNamedRange(name: string): { sheet: string; range: string } | null;
  /** True when the workbook / source contains this sheet name. */
  sheetExists(name: string): boolean;
}

export type FormulaResult = { kind: 'v'; value: CellPrimitive } | { kind: 'e'; value: ErrorValue };

export const ERR_CIRC: ErrorValue = '#CIRC!';

/* ----------------------------- Tokenizer ----------------------------- */

type TokenType =
  | 'number'
  | 'string'
  | 'ref'
  | 'fn'
  | 'op'
  | 'lp'
  | 'rp'
  | 'sep'
  | 'colon'
  | 'error'
  | 'eof';

interface Token {
  type: TokenType;
  value: string;
  pos: number;
}

const OPS = ['<>', '<=', '>=', '=', '&', '+', '-', '*', '/', '^', '%', '>', '<'];

export function tokenize(input: string): Token[] {
  const tokens: Token[] = [];
  let i = 0;
  const n = input.length;

  while (i < n) {
    const ch = input[i];

    if (ch === ' ' || ch === '\t' || ch === '\n' || ch === '\r') {
      i++;
      continue;
    }
    if (ch === '(') {
      tokens.push({ type: 'lp', value: '(', pos: i });
      i++;
      continue;
    }
    if (ch === ')') {
      tokens.push({ type: 'rp', value: ')', pos: i });
      i++;
      continue;
    }
    if (ch === ',' || ch === ';') {
      tokens.push({ type: 'sep', value: ch, pos: i });
      i++;
      continue;
    }
    if (ch === ':') {
      tokens.push({ type: 'colon', value: ':', pos: i });
      i++;
      continue;
    }
    if (ch === '"') {
      let j = i + 1;
      let out = '';
      while (j < n) {
        if (input[j] === '"') {
          if (input[j + 1] === '"') {
            out += '"';
            j += 2;
            continue;
          }
          break;
        }
        out += input[j];
        j++;
      }
      tokens.push({ type: 'string', value: out, pos: i });
      i = j + 1;
      continue;
    }

    // Numbers (incl. decimals and exponents)
    const numM = /^\d+(\.\d+)?([eE][+-]?\d+)?/.exec(input.slice(i));
    if (numM) {
      tokens.push({ type: 'number', value: numM[0], pos: i });
      i += numM[0].length;
      continue;
    }

    // Operators
    const op = OPS.find((o) => input.startsWith(o, i));
    if (op) {
      tokens.push({ type: 'op', value: op, pos: i });
      i += op.length;
      continue;
    }

    // Sheet-qualified reference or name (a-z, letters, digits, _, ., quotes)
    const nameM = /^('(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)/.exec(input.slice(i));
    if (nameM) {
      let name = nameM[1];
      const hasBang = input[i + name.length] === '!';
      if (hasBang) {
        let j = i + name.length + 1;
        while (j < n && /[A-Za-z0-9$]/.test(input[j])) j++;
        const refPart = input.slice(i + name.length + 1, j);
        // refPart like A1 or A1:B2
        const rgx = /^(\$?[A-Za-z]{1,4}\$?[0-9]{1,7})(?::(\$?[A-Za-z]{1,4}\$?[0-9]{1,7}))?$/.exec(refPart);
        if (rgx) {
          const rangeTok = name + '!' + refPart;
          tokens.push({ type: 'ref', value: rangeTok, pos: i });
          i = j;
          continue;
        }
        // else sheet-qualified function? Not valid; push as name
        tokens.push({ type: 'fn', value: name, pos: i });
        i += name.length;
        continue;
      }
      // could be ref (A1), named range, or function
      if (/^[A-Za-z]{1,4}\d{1,7}$/.test(name)) {
        tokens.push({ type: 'ref', value: name, pos: i });
        i += name.length;
        continue;
      }
      // function name if followed by '('
      let j = i + name.length;
      while (j < n && input[j] === ' ') j++;
      if (input[j] === '(') {
        tokens.push({ type: 'fn', value: name.toUpperCase(), pos: i });
        i += name.length;
        continue;
      }
      tokens.push({ type: 'fn', value: name.toUpperCase(), pos: i });
      i += name.length;
      continue;
    }

    // Error literals
    const errM = /^#[A-Z0-9]+\??!?/.exec(input.slice(i));
    if (errM) {
      tokens.push({ type: 'error', value: errM[0], pos: i });
      i += errM[0].length;
      continue;
    }

    // Range colon fallback already handled; anything else: skip silently
    i++;
  }

  tokens.push({ type: 'eof', value: '', pos: n });
  return tokens;
}

/* ------------------------------ Parser ------------------------------- */

export type Node =
  | { kind: 'num'; value: number }
  | { kind: 'str'; value: string }
  | { kind: 'bool'; value: boolean }
  | { kind: 'err'; value: ErrorValue }
  | { kind: 'ref'; sheet?: string; col: number; row: number; absCol?: boolean; absRow?: boolean }
  | { kind: 'range'; sheet?: string; sCol: number; sRow: number; eCol: number; eRow: number }
  | { kind: 'named'; name: string }
  | { kind: 'fname'; name: string; args: Node[] }
  | { kind: 'bin'; op: string; left: Node; right: Node }
  | { kind: 'un'; op: string; operand: Node }
  | { kind: 'paren'; expr: Node };

class Parser {
  private tokens: Token[];
  private pos = 0;

  constructor(input: string) {
    // Strip a leading '=' if present.
    if (input.startsWith('=')) input = input.slice(1);
    this.tokens = tokenize(input);
  }

  private peek(): Token {
    return this.tokens[this.pos] ?? { type: 'eof', value: '', pos: -1 };
  }
  private next(): Token {
    return this.tokens[this.pos++] ?? { type: 'eof', value: '', pos: -1 };
  }
  private atOp(op: string): boolean {
    const t = this.peek();
    return t.type === 'op' && t.value === op;
  }

  parse(): Node {
    const expr = this.parseComparison();
    return expr;
  }

  private parseComparison(): Node {
    let left = this.parseAdditive();
    while (true) {
      const t = this.peek();
      if (t.type === 'op' && ['=', '<>', '<', '>', '<=', '>='].includes(t.value)) {
        this.next();
        const right = this.parseAdditive();
        left = { kind: 'bin', op: t.value, left, right };
      } else break;
    }
    return left;
  }

  private parseAdditive(): Node {
    let left = this.parseMultiplicative();
    while (true) {
      const t = this.peek();
      if (t.type === 'op' && (t.value === '+' || t.value === '-')) {
        this.next();
        const right = this.parseMultiplicative();
        left = { kind: 'bin', op: t.value, left, right };
      } else if (t.type === 'op' && t.value === '&') {
        this.next();
        const right = this.parseMultiplicative();
        left = { kind: 'bin', op: '&', left, right };
      } else break;
    }
    return left;
  }

  private parseMultiplicative(): Node {
    let left = this.parseUnary();
    while (true) {
      const t = this.peek();
      if (t.type === 'op' && (t.value === '*' || t.value === '/')) {
        this.next();
        const right = this.parseUnary();
        left = { kind: 'bin', op: t.value, left, right };
      } else break;
    }
    return left;
  }

  private parseUnary(): Node {
    const t = this.peek();
    if (t.type === 'op' && (t.value === '-' || t.value === '+')) {
      this.next();
      const operand = this.parseUnary();
      return { kind: 'un', op: t.value, operand };
    }
    return this.parsePower();
  }

  private parsePower(): Node {
    const left = this.parsePrimary();
    const t = this.peek();
    if (t.type === 'op' && t.value === '^') {
      this.next();
      const right = this.parseUnary();
      return { kind: 'bin', op: '^', left, right };
    }
    if (t.type === 'op' && t.value === '%') {
      this.next();
      return { kind: 'un', op: '%', operand: left };
    }
    return left;
  }

  private parsePrimary(): Node {
    const t = this.next();
    switch (t.type) {
      case 'number':
        return { kind: 'num', value: parseFloat(t.value) };
      case 'string':
        return { kind: 'str', value: t.value };
      case 'error':
        return { kind: 'err', value: t.value as ErrorValue };
      case 'lp': {
        const inner = this.parseComparison();
        const close = this.next();
        if (close.type !== 'rp') throw new Error('Expected )');
        return { kind: 'paren', expr: inner };
      }
      case 'ref': {
        const tok = t.value;
        const bang = tok.lastIndexOf('!');
        const sheetConv = (s: string) => {
          if (s.startsWith("'") && s.endsWith("'")) return s.slice(1, -1).replace(/''/g, "'");
          return s;
        };
        if (bang >= 0) {
          const sheet = sheetConv(tok.slice(0, bang));
          const rest = tok.slice(bang + 1);
          return this.parseRefOrRange(rest, sheet);
        }
        return this.parseRefOrRange(tok, undefined);
      }
      case 'fn': {
        if (t.value === 'TRUE' || t.value === 'FALSE') return { kind: 'bool', value: t.value === 'TRUE' };
        // A bare identifier is not a call: it is a defined name (e.g. =SUM(Total)).
        if (this.peek().type !== 'lp') return { kind: 'named', name: t.value };
        const args: Node[] = [];
        const lp = this.next();
        if (lp.type !== 'lp') throw new Error('Expected ( after function ' + t.value);
        if (this.peek().type === 'rp') {
          this.next();
        } else {
          const first = this.parseComparison();
          args.push(first);
          while (this.peek().type === 'sep') {
            this.next();
            args.push(this.parseComparison());
          }
          const close = this.next();
          if (close.type !== 'rp') throw new Error('Expected )');
        }
        return { kind: 'fname', name: t.value, args };
      }
      default:
        throw new Error('Unexpected token ' + t.value);
    }
  }

  private parseRefOrRange(rest: string, sheet?: string): Node {
    const parts = rest.split(':');
    const parseRef = (s: string) => {
      const m = /^(\$?)([A-Za-z]{1,4})(\$?)([0-9]{1,7})$/.exec(s);
      if (!m) throw new Error('Invalid reference ' + s);
      let col = -1;
      for (const ch of m[2].toUpperCase()) col = (col + 1) * 26 + (ch.charCodeAt(0) - 65);
      const row = parseInt(m[4], 10);
      return { col, row, absCol: m[1] === '$', absRow: m[3] === '$' };
    };
    const a = parseRef(parts[0]);
    if (parts.length === 1) {
      // Bare range like A1:B2 arrives tokenized as ref, colon, ref — consume it.
      if (this.peek().type === 'colon') {
        this.next();
        const end = this.next();
        if (end.type !== 'ref') throw new Error('Invalid range end ' + end.value);
        const b = parseRef(end.value);
        return {
          kind: 'range',
          sheet,
          sCol: Math.min(a.col, b.col),
          sRow: Math.min(a.row, b.row),
          eCol: Math.max(a.col, b.col),
          eRow: Math.max(a.row, b.row)
        };
      }
      return { kind: 'ref', sheet, ...a };
    }
    const b = parseRef(parts[1]);
    return {
      kind: 'range',
      sheet,
      sCol: Math.min(a.col, b.col),
      sRow: Math.min(a.row, b.row),
      eCol: Math.max(a.col, b.col),
      eRow: Math.max(a.row, b.row)
    };
  }
}

export function parseFormula(input: string): Node {
  try {
    return new Parser(input).parse();
  } catch (e) {
    throw new Error('Formula parse error: ' + (e as Error).message);
  }
}

/* ----------------------------- Evaluation ---------------------------- */

export class Evaluator {
  private src: FormulaSource;
  private stack: string[] = [];
  private volatileDepth = false;

  constructor(src: FormulaSource) {
    this.src = src;
  }

  evaluateFormula(formula: string, currentSheet: string): FormulaResult {
    this.stack.length = 0;
    return this.evaluateFormulaKeepStack(formula, currentSheet);
  }

  /** Evaluate without clearing the call stack (for nested evaluations). */
  evaluateFormulaKeepStack(formula: string, currentSheet: string): FormulaResult {
    let ast: Node;
    try {
      ast = parseFormula(formula);
    } catch {
      return { kind: 'e', value: '#NAME?' };
    }
    try {
      const v = this.eval(ast, currentSheet);
      if (v instanceof EvalError) {
        return { kind: 'e', value: v.value };
      }
      return { kind: 'v', value: coerceScalar(v) };
    } catch (e) {
      const msg = String((e as Error).message ?? e);
      if (msg.includes('circular') || msg.includes('Circular')) return { kind: 'e', value: ERR_CIRC };
      return { kind: 'e', value: '#VALUE!' };
    }
  }

  private eval(node: Node, sheet: string): CellPrimitive | ErrorValue | RangeValue | EvalErrorMarker {
    switch (node.kind) {
      case 'num':
        return node.value;
      case 'str':
        return node.value;
      case 'bool':
        return node.value;
      case 'err':
        return node.value as ErrorValue;
      case 'paren':
        return this.eval(node.expr, sheet);
      case 'un': {
        if (node.op === '%') {
          const v = this.eval(node.operand, sheet) as any;
          const n = toNumber(scalarize(v));
          if (isErr(n)) return n;
          return (n as number) / 100;
        }
        const v = this.eval(node.operand, sheet) as any;
        const n = toNumber(scalarize(v));
        if (isErr(n)) return n;
        return node.op === '-' ? -(n as number) : n as number;
      }
      case 'bin':
        return this.evalBinary(node, sheet);
      case 'ref': {
        const targetSheet = node.sheet ?? sheet;
        if (node.sheet && !this.src.sheetExists(node.sheet)) return '#REF!';
        const key = targetSheet + '!' + node.row + ',' + node.col;
        if (this.stack.includes(key)) {
          throw new Error('Circular reference at ' + key);
        }
        this.stack.push(key);
        const val = this.src.getCell(targetSheet, node.row, node.col);
        this.stack.pop();
        return val === null || val === undefined ? null : val;
      }
      case 'range': {
        const targetSheet = node.sheet ?? sheet;
        if (node.sheet && !this.src.sheetExists(node.sheet)) return '#REF!';
        const rows: Array<Array<CellPrimitive | ErrorValue>> = [];
        for (let r = node.sRow; r <= node.eRow; r++) {
          const row: Array<CellPrimitive | ErrorValue> = [];
          for (let c = node.sCol; c <= node.eCol; c++) {
            const key = targetSheet + '!' + r + ',' + c;
            if (this.stack.includes(key)) return new EvalError(ERR_CIRC) as any;
            this.stack.push(key);
            const val = this.src.getCell(targetSheet, r, c);
            this.stack.pop();
            row.push(val === null || val === undefined ? null : val);
          }
          rows.push(row);
        }
        return { kind: 'range', rows };
      }
      case 'named': {
        const nr = this.src.resolveNamedRange(node.name);
        if (!nr) return '#NAME?';
        const p = parseRangePart(nr.range);
        if (!p) return '#NAME?';
        const targetSheet = nr.sheet;
        const rows: Array<Array<CellPrimitive | ErrorValue>> = [];
        for (let r = p.sRow; r <= p.eRow; r++) {
          const row: Array<CellPrimitive | ErrorValue> = [];
          for (let c = p.sCol; c <= p.eCol; c++) {
            const key = targetSheet + '!' + r + ',' + c;
            if (this.stack.includes(key)) return new EvalError(ERR_CIRC) as any;
            this.stack.push(key);
            const val = this.src.getCell(targetSheet, r, c);
            this.stack.pop();
            row.push(val === null || val === undefined ? null : val);
          }
          rows.push(row);
        }
        return { kind: 'range', rows };
      }
      case 'fname':
        return this.evalFunction(node, sheet);
    }
  }

  private evalBinary(node: Extract<Node, { kind: 'bin' }>, sheet: string): CellPrimitive | ErrorValue {
    // Short-circuit for comparison of strings handled generically.
    if (node.op === '&') {
      const a = this.eval(node.left, sheet) as any;
      const b = this.eval(node.right, sheet) as any;
      const ta = toText(scalarize(a));
      const tb = toText(scalarize(b));
      if (isErr(ta)) return ta;
      if (isErr(tb)) return tb;
      return String(ta) + String(tb);
    }
    const a = this.eval(node.left, sheet) as any;
    const b = this.eval(node.right, sheet) as any;
    const na = toNumber(scalarize(a));
    const nb = toNumber(scalarize(b));
    if (node.op === '^') {
      const r = Math.pow(na as number, nb as number);
      return Number.isNaN(r) || !Number.isFinite(r) ? '#NUM!' : r;
    }
    if (isErr(na)) return node.op === '=' || node.op === '<>' ? this.compareErrs(node.op, a, b) : na;
    if (isErr(nb)) return node.op === '=' || node.op === '<>' ? this.compareErrs(node.op, a, b) : nb;
    switch (node.op) {
      case '+':
        return (na as number) + (nb as number);
      case '-':
        return (na as number) - (nb as number);
      case '*':
        return (na as number) * (nb as number);
      case '/':
        return (nb as number) === 0 ? '#DIV/0!' : (na as number) / (nb as number);
      case '=':
        return this.cmpEq(a, b);
      case '<>':
        return !this.cmpEq(a, b);
      case '<':
        return this.cmp(a, b) < 0;
      case '>':
        return this.cmp(a, b) > 0;
      case '<=':
        return this.cmp(a, b) <= 0;
      case '>=':
        return this.cmp(a, b) >= 0;
    }
    return '#VALUE!';
  }

  private compareErrs(op: string, a: unknown, b: unknown): boolean {
    const sa = String(a);
    const sb = String(b);
    if (op === '=') return sa === sb;
    return sa !== sb;
  }

  private cmpEq(a: CellPrimitive | ErrorValue, b: CellPrimitive | ErrorValue): boolean {
    if (a === null && b === null) return true;
    if (typeof a === 'number' && typeof b === 'number') return a === b;
    if (typeof a === 'number' || typeof b === 'number') {
      const na = toNumber(a);
      const nb = toNumber(b);
      if (!isErr(na) && !isErr(nb)) return na === nb;
      return String(a) === String(b);
    }
    return String(a) === String(b);
  }

  private cmp(a: CellPrimitive | ErrorValue, b: CellPrimitive | ErrorValue): number {
    const na = toNumber(a);
    const nb = toNumber(b);
    if (!isErr(na) && !isErr(nb)) return (na as number) - (nb as number);
    return String(a).localeCompare(String(b));
  }

  private evalFunction(node: Extract<Node, { kind: 'fname' }>, sheet: string): CellPrimitive | ErrorValue | RangeValue | EvalErrorMarker {
    // Lazy functions must short-circuit argument evaluation.
    switch (node.name) {
      case 'IF': {
        const cond = this.eval(node.args[0], sheet);
        const b = toBoolish(cond);
        if (b instanceof EvalError) return b;
        return this.eval(node.args[b ? 1 : 2], sheet) as any;
      }
      case 'IFS': {
        for (let i = 0; i + 1 < node.args.length; i += 2) {
          const cond = this.eval(node.args[i], sheet);
          const b = toBoolish(cond);
          if (b instanceof EvalError) return b as unknown as EvalErrorMarker;
          if (b) return this.eval(node.args[i + 1], sheet) as any;
        }
        return '#N/A' as ErrorValue;
      }
      case 'IFERROR': {
        const v = this.eval(node.args[0], sheet);
        if (v instanceof EvalError || isErr(v)) return this.eval(node.args[1], sheet) as any;
        return v as any;
      }
      case 'IFNA': {
        const v = this.eval(node.args[0], sheet);
        if (v instanceof EvalError || v === '#N/A') return this.eval(node.args[1], sheet) as any;
        return v as any;
      }
      case 'SWITCH': {
        const expr = this.eval(node.args[0], sheet);
        for (let i = 1; i + 1 < node.args.length; i += 2) {
          const c = this.eval(node.args[i], sheet);
          if (String(c) === String(expr)) return this.eval(node.args[i + 1], sheet) as any;
        }
        if (node.args.length % 2 === 0) return this.eval(node.args[node.args.length - 1], sheet) as any;
        return '#N/A' as ErrorValue;
      }
    }

    const fn = getFunction(node.name);
    if (!fn) return '#NAME?';

    const args: ArgValue[] = [];
    for (const argNode of node.args) {
      const v = this.eval(argNode, sheet) as any;
      if (v instanceof EvalError) return v as unknown as EvalErrorMarker;
      args.push(v);
    }
    return ((fn as unknown as { fn: (a: ArgValue[]) => CellPrimitive | ErrorValue }).fn ?? (fn as (a: ArgValue[]) => CellPrimitive | ErrorValue))(args);
  }
}

class EvalError {
  constructor(public value: ErrorValue) {}
}

type EvalErrorMarker = InstanceType<typeof EvalError>;

export function isEvalError(v: unknown): v is EvalErrorMarker {
  return v instanceof EvalError;
}

function isErr(v: unknown): v is ErrorValue {
  return typeof v === 'string' && (v.startsWith('#') && (v.endsWith('!') || v === '#N/A' || v === '#NULL!' || v === '#CIRC!' || v === '#SPILL!'));
}

function toBoolish(v: unknown): boolean | EvalError {
  if (v instanceof EvalError) return v;
  if (typeof v === 'boolean') return v;
  const n = toNumber(scalarize(v as any));
  if (isErr(n)) return false;
  return (n as number) !== 0;
}

/** Reduce any evaluation result down to a CellPrimitive for a FormulaResult. */
export function coerceScalar(v: unknown): CellPrimitive {
  if (v === null || v === undefined) return null;
  if (typeof v === 'string' || typeof v === 'boolean' || typeof v === 'number') return v as CellPrimitive;
  if (v instanceof EvalError) return v.value;
  if (typeof v === 'object' && 'rows' in (v as object) && Array.isArray((v as { rows: unknown }).rows)) {
    const rows = (v as { rows: Array<Array<unknown>> }).rows;
    const first = rows[0]?.[0];
    if (first === null || first === undefined) return null;
    if (typeof first === 'string' || typeof first === 'boolean' || typeof first === 'number') return first as CellPrimitive;
  }
  return '#VALUE!';
}

function parseRangePart(range: string): { sRow: number; sCol: number; eRow: number; eCol: number } | null {
  const parts = range.split(':');
  const p = (s: string) => /^(\$?)([A-Za-z]{1,4})(\$?)([0-9]{1,7})$/.exec(s);
  const a = p(parts[0]);
  if (!a) return null;
  let col = -1;
  for (const ch of a[2].toUpperCase()) col = (col + 1) * 26 + (ch.charCodeAt(0) - 65);
  const row = parseInt(a[4], 10);
  if (parts.length === 1) return { sRow: row, sCol: col, eRow: row, eCol: col };
  const b = p(parts[1]);
  if (!b) return null;
  let col2 = -1;
  for (const ch of b[2].toUpperCase()) col2 = (col2 + 1) * 26 + (ch.charCodeAt(0) - 65);
  const row2 = parseInt(b[4], 10);
  return {
    sRow: Math.min(row, row2),
    sCol: Math.min(col, col2),
    eRow: Math.max(row, row2),
    eCol: Math.max(col, col2)
  };
}

/** Convenience: evaluate a formula against a source. */
export function evaluate(formula: string, sheet: string, src: FormulaSource): FormulaResult {
  return new Evaluator(src).evaluateFormula(formula, sheet);
}

/** True when a formula uses TODAY()/NOW()/RAND()/RANDBETWEEN (volatile). */
export function isVolatile(formula: string): boolean {
  const up = formula.toUpperCase();
  return up.includes('TODAY(') || up.includes('NOW(') || up.includes('RAND(') || up.includes('RANDBETWEEN(');
}