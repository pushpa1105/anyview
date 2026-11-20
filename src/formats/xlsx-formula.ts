/**
 * A deliberately small formula evaluator, used only to decide whether a
 * conditional-formatting rule applies to a cell. It understands constants,
 * single-cell references, comparison and arithmetic operators and a handful of
 * functions. Anything else makes `compileFormula` return `null`, and the rule
 * is then skipped rather than guessed at.
 */

export type FormulaValue = number | string | boolean | null;

/** Reads the value of a cell (zero-based row and column); `null` when empty. */
export type CellReader = (row: number, col: number) => FormulaValue;

/** A compiled formula; `rowShift`/`colShift` move its relative references. */
export type CompiledFormula = (
  read: CellReader,
  rowShift: number,
  colShift: number,
  row: number,
  col: number,
) => FormulaValue;

type Node =
  | { kind: 'value'; value: FormulaValue }
  | { kind: 'ref'; row: number; col: number; rowFixed: boolean; colFixed: boolean }
  | { kind: 'unary'; operator: string; operand: Node }
  | { kind: 'binary'; operator: string; left: Node; right: Node }
  | { kind: 'call'; name: string; args: Node[] };

interface Context {
  read: CellReader;
  rowShift: number;
  colShift: number;
  row: number;
  col: number;
}

class Unsupported extends Error {}

const REFERENCE = /^(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})$/;

function tokenize(source: string): string[] {
  const tokens: string[] = [];
  let i = 0;
  while (i < source.length) {
    const char = source[i]!;
    if (/\s/.test(char)) {
      i++;
    } else if (char === '"') {
      let end = i + 1;
      while (end < source.length && (source[end] !== '"' || source[end + 1] === '"')) {
        end += source[end] === '"' ? 2 : 1;
      }
      if (end >= source.length) throw new Unsupported();
      tokens.push(source.slice(i, end + 1));
      i = end + 1;
    } else if (/[0-9.]/.test(char)) {
      const match = /^\d*\.?\d+(?:[eE][+-]?\d+)?/.exec(source.slice(i));
      if (!match) throw new Unsupported();
      tokens.push(match[0]);
      i += match[0].length;
    } else if (/[A-Za-z_$]/.test(char)) {
      const match = /^[A-Za-z_$][A-Za-z0-9_.$]*/.exec(source.slice(i))!;
      tokens.push(match[0]);
      i += match[0].length;
    } else if ('<>='.includes(char)) {
      const two = source.slice(i, i + 2);
      const operator = two === '<=' || two === '>=' || two === '<>' ? two : char;
      tokens.push(operator);
      i += operator.length;
    } else if ('+-*/&(),%^'.includes(char)) {
      tokens.push(char);
      i++;
    } else {
      // Ranges, sheet references, structured references, arrays…
      throw new Unsupported();
    }
  }
  return tokens;
}

function columnIndex(letters: string): number {
  let col = 0;
  for (const letter of letters.toUpperCase()) col = col * 26 + (letter.charCodeAt(0) - 64);
  return col - 1;
}

function parse(tokens: string[]): Node {
  let position = 0;
  const peek = (): string | undefined => tokens[position];
  const next = (): string => {
    const token = tokens[position++];
    if (token === undefined) throw new Unsupported();
    return token;
  };
  const binary = (operators: string[], operand: () => Node): (() => Node) => {
    return () => {
      let left = operand();
      while (operators.includes(peek() ?? '')) {
        const operator = next();
        left = { kind: 'binary', operator, left, right: operand() };
      }
      return left;
    };
  };

  const primary = (): Node => {
    const token = next();
    if (token === '(') {
      const inner = comparison();
      if (next() !== ')') throw new Unsupported();
      return inner;
    }
    if (token === '-' || token === '+')
      return { kind: 'unary', operator: token, operand: primary() };
    if (token.startsWith('"')) {
      return { kind: 'value', value: token.slice(1, -1).replace(/""/g, '"') };
    }
    if (/^[0-9.]/.test(token)) {
      let value = Number(token);
      if (peek() === '%') {
        next();
        value /= 100;
      }
      return { kind: 'value', value };
    }
    if (peek() === '(') {
      next();
      const args: Node[] = [];
      if (peek() === ')') next();
      else {
        for (;;) {
          args.push(comparison());
          const separator = next();
          if (separator === ')') break;
          if (separator !== ',') throw new Unsupported();
        }
      }
      return { kind: 'call', name: token.toUpperCase(), args };
    }
    const upper = token.toUpperCase();
    if (upper === 'TRUE') return { kind: 'value', value: true };
    if (upper === 'FALSE') return { kind: 'value', value: false };
    const reference = REFERENCE.exec(token);
    if (reference) {
      return {
        kind: 'ref',
        col: columnIndex(reference[2]!),
        row: Number(reference[4]) - 1,
        colFixed: reference[1] === '$',
        rowFixed: reference[3] === '$',
      };
    }
    // Defined names and the like.
    throw new Unsupported();
  };
  const power = binary(['^'], primary);
  const product = binary(['*', '/'], power);
  const sum = binary(['+', '-'], product);
  const concat = binary(['&'], sum);
  const comparison: () => Node = binary(['=', '<>', '<', '>', '<=', '>='], concat);

  const tree = comparison();
  if (position !== tokens.length) throw new Unsupported();
  return tree;
}

const toNumber = (value: FormulaValue): number => {
  if (typeof value === 'number') return value;
  if (typeof value === 'boolean') return value ? 1 : 0;
  if (value === null || value === '') return 0;
  const parsed = Number(value);
  if (Number.isNaN(parsed)) throw new Unsupported();
  return parsed;
};
const toText = (value: FormulaValue): string =>
  value === null ? '' : typeof value === 'boolean' ? (value ? 'TRUE' : 'FALSE') : String(value);
const toBoolean = (value: FormulaValue): boolean =>
  typeof value === 'string' ? value.toUpperCase() === 'TRUE' : Boolean(value);

function compare(left: FormulaValue, right: FormulaValue): number {
  // Empty cells compare as 0 or "" depending on the other side.
  const a = left ?? (typeof right === 'string' ? '' : 0);
  const b = right ?? (typeof a === 'string' ? '' : 0);
  if (typeof a === 'number' && typeof b === 'number') return a - b;
  if (typeof a === 'string' && typeof b === 'string') {
    return a.toLowerCase() < b.toLowerCase() ? -1 : a.toLowerCase() > b.toLowerCase() ? 1 : 0;
  }
  if (typeof a === 'boolean' && typeof b === 'boolean') return Number(a) - Number(b);
  // Mixed types order as number < text < boolean.
  const rank = (value: FormulaValue): number =>
    typeof value === 'number' ? 0 : typeof value === 'string' ? 1 : 2;
  return rank(a) - rank(b);
}

const FUNCTIONS: Readonly<Record<string, (args: Node[], context: Context) => FormulaValue>> = {
  AND: (args, context) => args.every((arg) => toBoolean(evaluate(arg, context))),
  OR: (args, context) => args.some((arg) => toBoolean(evaluate(arg, context))),
  NOT: (args, context) => !toBoolean(evaluate(args[0]!, context)),
  IF: (args, context) =>
    toBoolean(evaluate(args[0]!, context))
      ? args[1]
        ? evaluate(args[1], context)
        : true
      : args[2]
        ? evaluate(args[2], context)
        : false,
  LEN: (args, context) => toText(evaluate(args[0]!, context)).length,
  TRIM: (args, context) => toText(evaluate(args[0]!, context)).trim().replace(/ +/g, ' '),
  UPPER: (args, context) => toText(evaluate(args[0]!, context)).toUpperCase(),
  LOWER: (args, context) => toText(evaluate(args[0]!, context)).toLowerCase(),
  LEFT: (args, context) =>
    toText(evaluate(args[0]!, context)).slice(
      0,
      args[1] ? toNumber(evaluate(args[1], context)) : 1,
    ),
  RIGHT: (args, context) => {
    const text = toText(evaluate(args[0]!, context));
    const count = args[1] ? toNumber(evaluate(args[1], context)) : 1;
    return count <= 0 ? '' : text.slice(-count);
  },
  ISBLANK: (args, context) => evaluate(args[0]!, context) === null,
  ISNUMBER: (args, context) => typeof evaluate(args[0]!, context) === 'number',
  ISTEXT: (args, context) => typeof evaluate(args[0]!, context) === 'string',
  ABS: (args, context) => Math.abs(toNumber(evaluate(args[0]!, context))),
  MOD: (args, context) => {
    const divisor = toNumber(evaluate(args[1]!, context));
    if (divisor === 0) throw new Unsupported();
    const value = toNumber(evaluate(args[0]!, context));
    return value - divisor * Math.floor(value / divisor);
  },
  ROW: (args, context) => {
    if (args.length === 0) return context.row + 1;
    const target = args[0]!;
    if (target.kind !== 'ref') throw new Unsupported();
    return target.row + (target.rowFixed ? 0 : context.rowShift) + 1;
  },
  COLUMN: (args, context) => {
    if (args.length === 0) return context.col + 1;
    const target = args[0]!;
    if (target.kind !== 'ref') throw new Unsupported();
    return target.col + (target.colFixed ? 0 : context.colShift) + 1;
  },
  // Days since 1899-12-30, the spreadsheet date serial.
  TODAY: () => Math.floor(Date.now() / 86_400_000) + 25_569,
};

function evaluate(node: Node, context: Context): FormulaValue {
  switch (node.kind) {
    case 'value':
      return node.value;
    case 'ref':
      return context.read(
        node.row + (node.rowFixed ? 0 : context.rowShift),
        node.col + (node.colFixed ? 0 : context.colShift),
      );
    case 'unary': {
      const value = toNumber(evaluate(node.operand, context));
      return node.operator === '-' ? -value : value;
    }
    case 'call': {
      const fn = FUNCTIONS[node.name];
      if (!fn || (node.args.length === 0 && !['ROW', 'COLUMN', 'TODAY'].includes(node.name))) {
        throw new Unsupported();
      }
      return fn(node.args, context);
    }
    case 'binary': {
      const left = evaluate(node.left, context);
      const right = evaluate(node.right, context);
      switch (node.operator) {
        case '&':
          return toText(left) + toText(right);
        case '+':
          return toNumber(left) + toNumber(right);
        case '-':
          return toNumber(left) - toNumber(right);
        case '*':
          return toNumber(left) * toNumber(right);
        case '/': {
          const divisor = toNumber(right);
          if (divisor === 0) throw new Unsupported();
          return toNumber(left) / divisor;
        }
        case '^':
          return toNumber(left) ** toNumber(right);
        case '=':
          return compare(left, right) === 0;
        case '<>':
          return compare(left, right) !== 0;
        case '<':
          return compare(left, right) < 0;
        case '>':
          return compare(left, right) > 0;
        case '<=':
          return compare(left, right) <= 0;
        default:
          return compare(left, right) >= 0;
      }
    }
  }
}

/**
 * Compiles a formula (without the leading `=`). Returns `null` when it uses
 * anything this evaluator does not understand. The compiled function returns
 * `null` when evaluation fails for a particular cell.
 */
export function compileFormula(source: string): CompiledFormula | null {
  let tree: Node;
  try {
    tree = parse(tokenize(source.replace(/^=/, '')));
  } catch {
    return null;
  }
  return (read, rowShift, colShift, row, col) => {
    try {
      return evaluate(tree, { read, rowShift, colShift, row, col });
    } catch {
      return null;
    }
  };
}

/** Truthiness of a formula result, as conditional formatting sees it. */
export function isTruthy(value: FormulaValue): boolean {
  return typeof value === 'number' ? value !== 0 : value === true;
}
