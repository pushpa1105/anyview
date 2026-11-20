/**
 * Spreadsheet number formatting (the "0.00", "#,##0", "yyyy-mm-dd" mini-language
 * shared by Excel and friends). Covers numbers, percentages, thousands
 * separators, scaling commas, scientific notation, currency literals, dates,
 * times and elapsed time. Fractions and conditional sections fall back to the
 * general format.
 */

export const BUILTIN_FORMATS: Readonly<Record<number, string>> = {
  0: 'General',
  1: '0',
  2: '0.00',
  3: '#,##0',
  4: '#,##0.00',
  9: '0%',
  10: '0.00%',
  11: '0.00E+00',
  12: '# ?/?',
  13: '# ??/??',
  14: 'm/d/yyyy',
  15: 'd-mmm-yy',
  16: 'd-mmm',
  17: 'mmm-yy',
  18: 'h:mm AM/PM',
  19: 'h:mm:ss AM/PM',
  20: 'h:mm',
  21: 'h:mm:ss',
  22: 'm/d/yyyy h:mm',
  37: '#,##0;(#,##0)',
  38: '#,##0;[Red](#,##0)',
  39: '#,##0.00;(#,##0.00)',
  40: '#,##0.00;[Red](#,##0.00)',
  45: 'mm:ss',
  46: '[h]:mm:ss',
  47: 'mm:ss.0',
  48: '##0.0E+0',
  49: '@',
};

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
  'December',
];
const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

type Part =
  | { kind: 'literal'; text: string }
  | { kind: 'digit'; char: '0' | '#' | '?' }
  | { kind: 'point' }
  | { kind: 'comma' }
  | { kind: 'percent' }
  | { kind: 'exponent'; sign: string }
  | { kind: 'text' }
  | { kind: 'date'; token: string }
  | { kind: 'elapsed'; unit: string }
  | { kind: 'general' }
  | { kind: 'fraction' };

interface Section {
  parts: Part[];
  isDate: boolean;
  /** Colour named by a code such as [Red], as a CSS colour. */
  color?: string;
}

/** Marks the position of an asterisk (repeat-to-fill) code inside formatted text. */
const FILL_MARK = String.fromCharCode(1);

const FORMAT_COLORS: Readonly<Record<string, string>> = {
  black: '#000000',
  blue: '#0000ff',
  cyan: '#00ffff',
  green: '#008000',
  magenta: '#ff00ff',
  red: '#ff0000',
  white: '#ffffff',
  yellow: '#ffff00',
};

const sectionCache = new Map<string, Section[]>();

function splitSections(format: string): string[] {
  const sections: string[] = [];
  let current = '';
  let quoted = false;
  for (let i = 0; i < format.length; i++) {
    const char = format[i]!;
    if (char === '"') quoted = !quoted;
    if (char === '\\' && !quoted) {
      current += char + (format[i + 1] ?? '');
      i++;
    } else if (char === ';' && !quoted) {
      sections.push(current);
      current = '';
    } else {
      current += char;
    }
  }
  sections.push(current);
  return sections;
}

function parseSection(source: string): Section {
  const parts: Part[] = [];
  let isDate = false;
  let color: string | undefined;
  let i = 0;
  const literal = (text: string): void => {
    const last = parts[parts.length - 1];
    if (last?.kind === 'literal') last.text += text;
    else parts.push({ kind: 'literal', text });
  };
  while (i < source.length) {
    const char = source[i]!;
    const lower = char.toLowerCase();
    if (char === '"') {
      const end = source.indexOf('"', i + 1);
      literal(source.slice(i + 1, end === -1 ? source.length : end));
      i = end === -1 ? source.length : end + 1;
    } else if (char === '\\') {
      literal(source[i + 1] ?? '');
      i += 2;
    } else if (char === '_') {
      literal(' ');
      i += 2;
    } else if (char === '*') {
      // Repeat-to-fill: the viewer pushes what follows to the far edge of the cell.
      if (!parts.some((part) => part.kind === 'literal' && part.text.includes(FILL_MARK)))
        literal(FILL_MARK);
      i += 2;
    } else if (char === '[') {
      const end = source.indexOf(']', i);
      const body = source.slice(i + 1, end === -1 ? source.length : end);
      if (/^(h+|m+|s+)$/i.test(body)) {
        parts.push({ kind: 'elapsed', unit: body.toLowerCase() });
        isDate = true;
      } else if (body.startsWith('$')) {
        // [$€-407] → currency symbol "€".
        literal(body.slice(1).split('-')[0] ?? '');
      } else {
        color ??= FORMAT_COLORS[body.toLowerCase()];
      }
      i = end === -1 ? source.length : end + 1;
    } else if (/^general/i.test(source.slice(i, i + 7))) {
      parts.push({ kind: 'general' });
      i += 7;
    } else if (/^am\/pm/i.test(source.slice(i, i + 5))) {
      parts.push({ kind: 'date', token: 'AM/PM' });
      isDate = true;
      i += 5;
    } else if (/^a\/p/i.test(source.slice(i, i + 3))) {
      parts.push({ kind: 'date', token: 'A/P' });
      isDate = true;
      i += 3;
    } else if ('ymdhs'.includes(lower)) {
      let end = i + 1;
      while (end < source.length && source[end]!.toLowerCase() === lower) end++;
      parts.push({ kind: 'date', token: lower.repeat(end - i) });
      isDate = true;
      i = end;
    } else if (char === '0' || char === '#' || char === '?') {
      parts.push({ kind: 'digit', char });
      i++;
    } else if (char === '.') {
      parts.push({ kind: 'point' });
      i++;
    } else if (char === ',') {
      parts.push({ kind: 'comma' });
      i++;
    } else if (char === '%') {
      parts.push({ kind: 'percent' });
      i++;
    } else if (char === '/' && parts.some((part) => part.kind === 'digit') && !isDate) {
      parts.push({ kind: 'fraction' });
      i++;
    } else if ((char === 'E' || char === 'e') && /[+-]/.test(source[i + 1] ?? '')) {
      parts.push({ kind: 'exponent', sign: source[i + 1]! });
      i += 2;
    } else if (char === '@') {
      parts.push({ kind: 'text' });
      i++;
    } else {
      literal(char);
      i++;
    }
  }
  return color ? { parts, isDate, color } : { parts, isDate };
}

function sectionsOf(format: string): Section[] {
  let sections = sectionCache.get(format);
  if (!sections) {
    sections = splitSections(format).map(parseSection);
    if (sectionCache.size > 500) sectionCache.clear();
    sectionCache.set(format, sections);
  }
  return sections;
}

/** True when a format code displays its value as a date or time. */
export function isDateFormat(format: string): boolean {
  return sectionsOf(format)[0]?.isDate ?? false;
}

/** The "General" format: up to 11 significant digits, exponent for extremes. */
export function formatGeneral(value: number): string {
  if (!Number.isFinite(value)) return String(value);
  if (value === 0) return '0';
  const magnitude = Math.abs(value);
  if (magnitude >= 1e11 || magnitude < 1e-9) {
    const [mantissa, exponent] = value.toExponential(5).split('e') as [string, string];
    const power = Number(exponent);
    const digits = mantissa.includes('.') ? mantissa.replace(/\.?0+$/, '') : mantissa;
    return `${digits}E${power < 0 ? '-' : '+'}${Math.abs(power)}`;
  }
  return String(Number(value.toPrecision(11)));
}

function pad(value: number, width: number): string {
  return String(value).padStart(width, '0');
}

/** Converts a spreadsheet serial date to a UTC `Date`. */
export function serialToDate(serial: number, date1904 = false): Date {
  let days = serial;
  if (date1904) days += 1462;
  // Excel's 1900 system wrongly treats 1900 as a leap year; serials before 61 are off by one.
  else if (days < 61) days += 1;
  const milliseconds = Math.round(days * 86_400_000);
  return new Date(Date.UTC(1899, 11, 30) + milliseconds);
}

function formatDate(serial: number, section: Section, date1904: boolean): string {
  const { parts } = section;
  // Round to the smallest displayed unit so 0.999999 does not show as 23:59:59.
  const fractionDigits = parts.reduce((digits, part, index) => {
    if (part.kind !== 'point') return digits;
    let count = 0;
    while (parts[index + 1 + count]?.kind === 'digit') count++;
    return Math.max(digits, count);
  }, 0);
  const unit = 86_400 * 10 ** fractionDigits;
  const rounded = Math.round(serial * unit) / unit;
  const date = serialToDate(rounded, date1904);
  const totalSeconds = rounded * 86_400;
  const hasMeridiem = parts.some((part) => part.kind === 'date' && part.token.includes('/'));
  const hours = date.getUTCHours();
  let output = '';

  for (let index = 0; index < parts.length; index++) {
    const part = parts[index]!;
    if (part.kind === 'literal') {
      output += part.text;
    } else if (part.kind === 'elapsed') {
      const divisor = part.unit[0] === 'h' ? 3600 : part.unit[0] === 'm' ? 60 : 1;
      output += pad(Math.floor(totalSeconds / divisor + 1e-9), part.unit.length);
    } else if (part.kind === 'point') {
      let digits = 0;
      while (parts[index + 1 + digits]?.kind === 'digit') digits++;
      if (digits > 0) {
        const fraction = totalSeconds - Math.floor(totalSeconds + 1e-9);
        output += '.' + pad(Math.round(fraction * 10 ** digits) % 10 ** digits, digits);
        index += digits;
      } else {
        output += '.';
      }
    } else if (part.kind === 'comma') {
      output += ',';
    } else if (part.kind === 'date') {
      const token = part.token;
      const length = token.length;
      switch (token[0]) {
        case 'y':
          output +=
            length <= 2 ? pad(date.getUTCFullYear() % 100, 2) : String(date.getUTCFullYear());
          break;
        case 'm': {
          // "m" means minutes right after hours or right before seconds, otherwise month.
          const previous = findDateToken(parts, index, -1);
          const next = findDateToken(parts, index, 1);
          const isMinutes = previous === 'h' || next === 's' || (previous === '[' && length <= 2);
          if (isMinutes && length <= 2) output += pad(date.getUTCMinutes(), length);
          else if (length === 1) output += String(date.getUTCMonth() + 1);
          else if (length === 2) output += pad(date.getUTCMonth() + 1, 2);
          else if (length === 3) output += MONTHS[date.getUTCMonth()]!.slice(0, 3);
          else if (length === 5) output += MONTHS[date.getUTCMonth()]![0];
          else output += MONTHS[date.getUTCMonth()];
          break;
        }
        case 'd':
          if (length === 1) output += String(date.getUTCDate());
          else if (length === 2) output += pad(date.getUTCDate(), 2);
          else if (length === 3) output += DAYS[date.getUTCDay()]!.slice(0, 3);
          else output += DAYS[date.getUTCDay()];
          break;
        case 'h': {
          const shown = hasMeridiem ? hours % 12 || 12 : hours;
          output += pad(shown, Math.min(length, 2));
          break;
        }
        case 's':
          output += pad(date.getUTCSeconds(), Math.min(length, 2));
          break;
        case 'A':
          output += token === 'A/P' ? (hours < 12 ? 'A' : 'P') : hours < 12 ? 'AM' : 'PM';
          break;
      }
    }
  }
  return output;
}

/** Nearest date token before/after `index`: its first letter, or `"["` for elapsed time. */
function findDateToken(parts: Part[], index: number, step: 1 | -1): string {
  for (let i = index + step; i >= 0 && i < parts.length; i += step) {
    const part = parts[i]!;
    if (part.kind === 'date') return part.token[0]!;
    if (part.kind === 'elapsed') return '[';
  }
  return '';
}

function groupThousands(digits: string): string {
  return digits.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

function formatNumeric(value: number, section: Section): string {
  const { parts } = section;
  const firstDigit = parts.findIndex((part) => part.kind === 'digit');
  let lastDigit = -1;
  for (let i = parts.length - 1; i >= 0; i--) {
    if (parts[i]!.kind === 'digit') {
      lastDigit = i;
      break;
    }
  }
  if (firstDigit === -1) {
    // No digit placeholders: literals only (possibly with "General").
    return parts
      .map((part) =>
        part.kind === 'literal' ? part.text : part.kind === 'general' ? formatGeneral(value) : '',
      )
      .join('');
  }

  let scaled = value;
  let integerZeros = 0;
  let decimalZeros = 0;
  let decimalPlaces = 0;
  let grouping = false;
  let seenPoint = false;
  let exponent: string | null = null;
  let exponentDigits = 0;
  for (let i = 0; i < parts.length; i++) {
    const part = parts[i]!;
    if (part.kind === 'percent') scaled *= 100;
    else if (part.kind === 'point') seenPoint = true;
    else if (part.kind === 'exponent') exponent = part.sign;
    else if (part.kind === 'digit') {
      if (exponent !== null) exponentDigits++;
      else if (seenPoint) {
        decimalPlaces++;
        if (part.char === '0') decimalZeros = decimalPlaces;
      } else if (part.char === '0') integerZeros++;
    } else if (part.kind === 'comma') {
      // A comma between digits groups thousands; trailing commas scale by 1000.
      if (i > firstDigit && i < lastDigit && !seenPoint) grouping = true;
      else if (i > lastDigit) scaled /= 1000;
    }
  }

  let body: string;
  if (exponent !== null) {
    const [mantissa, power] = Math.abs(scaled).toExponential(decimalPlaces).split('e') as [
      string,
      string,
    ];
    const exponentValue = parseInt(power, 10);
    const sign = exponentValue < 0 ? '-' : exponent === '+' ? '+' : '';
    body = `${mantissa}E${sign}${pad(Math.abs(exponentValue), Math.max(exponentDigits, 1))}`;
  } else {
    const fixed = Math.abs(scaled).toFixed(Math.min(decimalPlaces, 20));
    let [integer, fraction = ''] = fixed.split('.') as [string, string | undefined];
    // Drop optional ("#") decimals that are zero.
    while (fraction.length > decimalZeros && fraction.endsWith('0'))
      fraction = fraction.slice(0, -1);
    if (integer === '0' && integerZeros === 0) integer = '';
    integer = integer.padStart(integerZeros, '0');
    if (grouping) integer = groupThousands(integer);
    body = fraction ? `${integer}.${fraction}` : integer;
  }

  // Literals before the first and after the last placeholder keep their place;
  // percent signs are literals too.
  let prefix = '';
  let suffix = '';
  for (let i = 0; i < parts.length; i++) {
    const part = parts[i]!;
    const text = part.kind === 'literal' ? part.text : part.kind === 'percent' ? '%' : '';
    if (!text) continue;
    if (i < firstDigit) prefix += text;
    else suffix += text;
  }
  return prefix + body + suffix;
}

/**
 * Formats a number with a spreadsheet format code.
 *
 * @example
 * formatNumber(1234.5, '#,##0.00')   // "1,234.50"
 * formatNumber(45000, 'yyyy-mm-dd')  // "2023-03-15"
 */
export function formatNumber(value: number, format: string, date1904 = false): string {
  return formatMarked(value, format, date1904).text.replace(FILL_MARK, '');
}

/** A formatted value split for display in a cell. */
export interface FormattedValue {
  text: string;
  /** Set for formats with a fill code: one part hugs the left edge, the other the right edge. */
  fill?: { left: string; right: string };
  /** Colour requested by the format, e.g. [Red]. */
  color?: string;
}

/**
 * Like {@link formatNumber}, and also reports fill alignment (accounting
 * formats put the currency symbol at the left edge) and colour codes.
 */
export function formatValue(value: number, format: string, date1904 = false): FormattedValue {
  const { text, color } = formatMarked(value, format, date1904);
  const mark = text.indexOf(FILL_MARK);
  const result: FormattedValue = { text: text.replace(FILL_MARK, '') };
  if (mark !== -1) result.fill = { left: text.slice(0, mark), right: text.slice(mark + 1) };
  if (color) result.color = color;
  return result;
}

function formatMarked(
  value: number,
  format: string,
  date1904: boolean,
): { text: string; color?: string } {
  const text = formatSection(value, format, date1904);
  return typeof text === 'string' ? { text } : text;
}

function formatSection(
  value: number,
  format: string,
  date1904: boolean,
): string | { text: string; color?: string } {
  if (!Number.isFinite(value)) return String(value);
  if (!format || /^general$/i.test(format)) return formatGeneral(value);
  const sections = sectionsOf(format);
  let section = sections[0]!;
  let magnitude = value;
  let explicitNegative = false;
  if (sections.length >= 2 && value < 0) {
    section = sections[1]!;
    magnitude = -value;
    explicitNegative = true;
  } else if (sections.length >= 3 && value === 0) {
    section = sections[2]!;
  }
  const colored = (text: string): string | { text: string; color?: string } =>
    section.color ? { text, color: section.color } : text;
  if (section.parts.some((part) => part.kind === 'fraction')) return formatGeneral(value);
  if (section.isDate) {
    return value < 0 ? formatGeneral(value) : colored(formatDate(value, section, date1904));
  }
  if (section.parts.length === 0) return '';
  if (section.parts.every((part) => part.kind === 'general' || part.kind === 'literal')) {
    return colored(formatNumeric(magnitude, section));
  }
  const text = formatNumeric(magnitude, section);
  // With a single section, negative numbers get a leading minus sign.
  const isZero = !/[1-9]/.test(text);
  return colored(value < 0 && !explicitNegative && !isZero ? `-${text}` : text);
}

/** Applies a text section (`@`) of a format to a string value. */
export function formatText(value: string, format: string): string {
  if (!format || !format.includes('@')) return value;
  const sections = sectionsOf(format);
  const section = sections.length >= 4 ? sections[3]! : sections[sections.length - 1]!;
  if (!section.parts.some((part) => part.kind === 'text')) return value;
  return section.parts
    .map((part) => (part.kind === 'text' ? value : part.kind === 'literal' ? part.text : ''))
    .join('')
    .replace(FILL_MARK, '');
}
