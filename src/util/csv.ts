export interface CsvResult {
  rows: string[][];
  /** Widest row. */
  columnCount: number;
  delimiter: string;
  /** Total number of rows in the input, which can exceed `rows.length` when `maxRows` applies. */
  totalRows: number;
}

export interface CsvOptions {
  /** Field separator. Detected from the first lines when omitted. */
  delimiter?: string;
  /** Stop keeping rows after this many (the rest are only counted). */
  maxRows?: number;
}

const CANDIDATES = [',', ';', '\t', '|'];

/**
 * Picks the delimiter that splits the first lines into the most consistent
 * number of fields. Quoted sections are ignored while counting.
 */
export function detectDelimiter(text: string): string {
  const sample = text.slice(0, 16_384);
  const lines: string[] = [];
  let current = '';
  let quoted = false;
  for (const char of sample) {
    if (char === '"') quoted = !quoted;
    if ((char === '\n' || char === '\r') && !quoted) {
      if (current) lines.push(current);
      current = '';
      if (lines.length >= 20) break;
    } else {
      current += char;
    }
  }
  if (current && lines.length < 20) lines.push(current);

  let best = ',';
  let bestScore = 0;
  for (const candidate of CANDIDATES) {
    const counts = lines.map((line) => {
      let count = 0;
      let inQuotes = false;
      for (const char of line) {
        if (char === '"') inQuotes = !inQuotes;
        else if (char === candidate && !inQuotes) count++;
      }
      return count;
    });
    const first = counts[0] ?? 0;
    if (first === 0) continue;
    const consistent = counts.filter((count) => count === first).length;
    const score = consistent * first;
    if (score > bestScore) {
      best = candidate;
      bestScore = score;
    }
  }
  return best;
}

/** Parses RFC 4180 CSV: quoted fields, doubled quotes, embedded newlines, CRLF or LF. */
export function parseCsv(input: string, options: CsvOptions = {}): CsvResult {
  const text = input.charCodeAt(0) === 0xfeff ? input.slice(1) : input;
  const delimiter = options.delimiter ?? detectDelimiter(text);
  const maxRows = options.maxRows ?? Infinity;
  const rows: string[][] = [];
  let columnCount = 0;
  let totalRows = 0;
  let row: string[] = [];
  let field = '';
  let i = 0;
  const length = text.length;

  const endRow = (): void => {
    row.push(field);
    field = '';
    // A lone empty field is a blank line, not a row.
    if (row.length > 1 || row[0] !== '') {
      totalRows++;
      if (rows.length < maxRows) {
        rows.push(row);
        if (row.length > columnCount) columnCount = row.length;
      }
    }
    row = [];
  };

  while (i < length) {
    const char = text[i]!;
    if (char === '"' && field === '') {
      // Quoted field: read up to the closing quote, un-doubling "" pairs.
      let end = i + 1;
      let value = '';
      for (;;) {
        const close = text.indexOf('"', end);
        if (close === -1) {
          value += text.slice(end);
          end = length;
          break;
        }
        value += text.slice(end, close);
        if (text[close + 1] === '"') {
          value += '"';
          end = close + 2;
        } else {
          end = close + 1;
          break;
        }
      }
      field = value;
      i = end;
      // Tolerate stray characters between the closing quote and the delimiter.
      while (i < length && text[i] !== delimiter && text[i] !== '\n' && text[i] !== '\r') {
        field += text[i];
        i++;
      }
    } else if (char === delimiter) {
      row.push(field);
      field = '';
      i++;
    } else if (char === '\n' || char === '\r') {
      if (char === '\r' && text[i + 1] === '\n') i++;
      i++;
      endRow();
    } else {
      // Unquoted run: copy up to the next delimiter or line break in one slice.
      let end = i + 1;
      while (end < length) {
        const next = text[end]!;
        if (next === delimiter || next === '\n' || next === '\r') break;
        end++;
      }
      field += text.slice(i, end);
      i = end;
    }
  }
  if (field !== '' || row.length > 0) endRow();
  return { rows, columnCount, delimiter, totalRows };
}
