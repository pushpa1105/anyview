import { describe, expect, it } from 'vitest';
import { detectDelimiter, parseCsv } from '../../src/util/csv';
import {
  formatGeneral,
  formatNumber,
  formatText,
  isDateFormat,
  serialToDate,
} from '../../src/util/numfmt';

describe('parseCsv', () => {
  it('parses simple rows with LF and CRLF line endings', () => {
    expect(parseCsv('a,b,c\n1,2,3\n').rows).toEqual([
      ['a', 'b', 'c'],
      ['1', '2', '3'],
    ]);
    expect(parseCsv('a,b\r\n1,2\r\n').rows).toEqual([
      ['a', 'b'],
      ['1', '2'],
    ]);
    expect(parseCsv('a,b\r1,2').rows).toEqual([
      ['a', 'b'],
      ['1', '2'],
    ]);
  });

  it('handles quoted fields, doubled quotes and embedded newlines', () => {
    const result = parseCsv('name,quote\n"Smith, John","He said ""hi"""\n"multi\nline",x\n');
    expect(result.rows).toEqual([
      ['name', 'quote'],
      ['Smith, John', 'He said "hi"'],
      ['multi\nline', 'x'],
    ]);
  });

  it('keeps empty fields and counts the widest row', () => {
    const result = parseCsv('a,,c\n,,\n1,2,3,4\n');
    expect(result.rows).toEqual([
      ['a', '', 'c'],
      ['', '', ''],
      ['1', '2', '3', '4'],
    ]);
    expect(result.columnCount).toBe(4);
  });

  it('skips blank lines and a byte-order mark', () => {
    expect(parseCsv('﻿a,b\n\n\n1,2').rows).toEqual([
      ['a', 'b'],
      ['1', '2'],
    ]);
  });

  it('tolerates an unterminated quote and text after a closing quote', () => {
    expect(parseCsv('a,"unterminated').rows).toEqual([['a', 'unterminated']]);
    expect(parseCsv('"a"x,b').rows).toEqual([['ax', 'b']]);
  });

  it('respects maxRows while still counting every row', () => {
    const result = parseCsv('1\n2\n3\n4\n5\n', { maxRows: 2 });
    expect(result.rows).toEqual([['1'], ['2']]);
    expect(result.totalRows).toBe(5);
  });

  it('returns nothing for empty input', () => {
    expect(parseCsv('')).toMatchObject({ rows: [], columnCount: 0, totalRows: 0 });
  });

  it('uses an explicit delimiter', () => {
    expect(parseCsv('a\tb,c\n', { delimiter: '\t' }).rows).toEqual([['a', 'b,c']]);
  });
});

describe('detectDelimiter', () => {
  it.each([
    ['a,b,c\n1,2,3', ','],
    ['a;b;c\n1;2;3', ';'],
    ['a\tb\tc\n1\t2\t3', '\t'],
    ['a|b|c\n1|2|3', '|'],
    ['"x;y",b,c\n"1;2",2,3', ','],
    ['no delimiters here', ','],
    ['', ','],
  ])('detects the delimiter of %j', (text, expected) => {
    expect(detectDelimiter(text)).toBe(expected);
  });

  it('prefers the delimiter that is consistent across lines', () => {
    // European CSV: semicolons separate fields, commas are decimal marks.
    expect(detectDelimiter('name;price\nA;1,50\nB;2,75\nC;3')).toBe(';');
  });
});

describe('formatGeneral', () => {
  it.each([
    [0, '0'],
    [42, '42'],
    [-1.5, '-1.5'],
    [0.1 + 0.2, '0.3'],
    [1234567.891, '1234567.891'],
    [1e21, '1E+21'],
    [1.5e-12, '1.5E-12'],
    [123456789012345, '1.23457E+14'],
  ])('formats %d as %s', (value, expected) => {
    expect(formatGeneral(value)).toBe(expected);
  });

  it('passes through non-finite values', () => {
    expect(formatGeneral(Number.NaN)).toBe('NaN');
    expect(formatNumber(Number.POSITIVE_INFINITY, '0.00')).toBe('Infinity');
  });
});

describe('formatNumber', () => {
  it.each([
    [1234.567, '0', '1235'],
    [1234.567, '0.00', '1234.57'],
    [1234.567, '#,##0', '1,235'],
    [1234567.891, '#,##0.00', '1,234,567.89'],
    [0.5, '#,##0.00', '0.50'],
    [0.5, '#.##', '.5'],
    [5, '000', '005'],
    [0.256, '0%', '26%'],
    [0.256, '0.0%', '25.6%'],
    [1234.5, '"$"#,##0.00', '$1,234.50'],
    [1234.5, '[$€-407] #,##0.00', '€ 1,234.50'],
    [1234.5, '#,##0.00 "kg"', '1,234.50 kg'],
    [1500000, '0.0,,"M"', '1.5M'],
    [12345, '0.00E+00', '1.23E+04'],
    [0.00012, '0.0E+0', '1.2E-4'],
    [-5, '0.00', '-5.00'],
    [-0.001, '0.00', '0.00'],
    [-5, '#,##0;(#,##0)', '(5)'],
    [-5, '#,##0.00;[Red](#,##0.00)', '(5.00)'],
    [0, '0;-0;"zero"', 'zero'],
    [7, '0;-0;"zero"', '7'],
    [3, 'General', '3'],
    [3, '', '3'],
    [3.14159, '0.0#', '3.14'],
    [3, '0.0#', '3.0'],
    [42, '"Total: "0', 'Total: 42'],
    [42, '\\#0', '#42'],
    [42, '0_)', '42 '],
    [42, '*-0', '42'],
    [1.5, '# ?/?', '1.5'],
    [5, '"text only"', 'text only'],
    [5, '"n="General', 'n=5'],
  ])('formats %d with %j as %j', (value, format, expected) => {
    expect(formatNumber(value, format)).toBe(expected);
  });

  it('returns an empty string for an empty section', () => {
    expect(formatNumber(0, '0;-0;')).toBe('');
  });
});

describe('date formats', () => {
  // 45000 is 15 March 2023, a Wednesday.
  it.each([
    [45000, 'yyyy-mm-dd', '2023-03-15'],
    [45000, 'm/d/yyyy', '3/15/2023'],
    [45000, 'd-mmm-yy', '15-Mar-23'],
    [45000, 'dddd, mmmm d, yyyy', 'Wednesday, March 15, 2023'],
    [45000, 'ddd mmmmm', 'Wed M'],
    [45000.5, 'yyyy-mm-dd hh:mm:ss', '2023-03-15 12:00:00'],
    [45000.75, 'h:mm AM/PM', '6:00 PM'],
    [45000.25, 'h:mm:ss A/P', '6:00:00 A'],
    [0.5, 'h:mm', '12:00'],
    [0.000011574, 'mm:ss', '00:01'],
    [1.5, '[h]:mm:ss', '36:00:00'],
    [0.5, '[m]', '720'],
    [0.5, '[s]', '43200'],
    [0.0000115741, 'ss.00', '01.00'],
    [45000.999999, 'yyyy-mm-dd hh:mm:ss', '2023-03-16 00:00:00'],
    [45000, 'dd.mm.yyyy', '15.03.2023'],
  ])('formats serial %d with %j as %j', (value, format, expected) => {
    expect(formatNumber(value, format)).toBe(expected);
  });

  it('supports the 1904 date system', () => {
    expect(formatNumber(43538, 'yyyy-mm-dd', true)).toBe('2023-03-15');
  });

  it('accounts for the 1900 leap-year bug', () => {
    expect(serialToDate(1).toISOString().slice(0, 10)).toBe('1900-01-01');
    expect(serialToDate(61).toISOString().slice(0, 10)).toBe('1900-03-01');
  });

  it('shows negative dates with the general format', () => {
    expect(formatNumber(-1, 'yyyy-mm-dd')).toBe('-1');
  });

  it('recognises date formats', () => {
    expect(isDateFormat('yyyy-mm-dd')).toBe(true);
    expect(isDateFormat('h:mm AM/PM')).toBe(true);
    expect(isDateFormat('[h]:mm')).toBe(true);
    expect(isDateFormat('#,##0.00')).toBe(false);
    expect(isDateFormat('"day"0')).toBe(false);
    expect(isDateFormat('General')).toBe(false);
  });
});

describe('formatText', () => {
  it('applies the text section of a format', () => {
    expect(formatText('abc', '@')).toBe('abc');
    expect(formatText('abc', '"[" @ "]"')).toBe('[ abc ]');
    expect(formatText('abc', '0;-0;0;"<"@">"')).toBe('<abc>');
    expect(formatText('abc', 'General')).toBe('abc');
    expect(formatText('abc', '')).toBe('abc');
    expect(formatText('a@b', '0.00')).toBe('a@b');
  });
});
