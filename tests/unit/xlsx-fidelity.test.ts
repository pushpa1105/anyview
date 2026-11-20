import { afterEach, describe, expect, it } from 'vitest';
import { drawingColor, parseChart, type ChartModel } from '../../src/formats/chart';
import { parseTheme, type OfficeTheme } from '../../src/formats/theme';
import { openWorkbook } from '../../src/formats/xlsx';
import {
  CF_STRIDE,
  evaluateConditionalFormats,
  parseRange,
  parseRanges,
  type CfCell,
} from '../../src/formats/xlsx-cf';
import { compileFormula, isTruthy, type FormulaValue } from '../../src/formats/xlsx-formula';
import {
  applyFont,
  builtinTableStyle,
  parseInlineDxf,
  parseStyles,
  themeColorAt,
} from '../../src/formats/xlsx-styles';
import { renderer as xlsxRenderer } from '../../src/renderers/xlsx';
import { axisScale, niceStep, renderChart } from '../../src/ui/chart';
import { createGrid, createGridIcon, type GridCell, type GridSource } from '../../src/ui/grid';
import { formatNumber, formatText, formatValue } from '../../src/util/numfmt';
import { parseXml } from '../../src/util/xml';
import { openZip } from '../../src/util/zip';
import { instanceOf, mount, renderContext, waitFor } from '../helpers/dom';
import { makeZip, PNG_1X1 } from '../helpers/zip';

afterEach(() => {
  document.body.textContent = '';
});

const THEME_XML = `<a:theme xmlns:a="a"><a:themeElements>
  <a:clrScheme name="t">
    <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
    <a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2>
    <a:accent1><a:srgbClr val="4F81BD"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2>
    <a:accent3><a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4>
    <a:accent5><a:srgbClr val="4BACC6"/></a:accent5><a:accent6><a:srgbClr val="F79646"/></a:accent6>
  </a:clrScheme>
  <a:fontScheme name="t"><a:majorFont><a:latin typeface="Impact"/></a:majorFont><a:minorFont><a:latin typeface="Corbel"/></a:minorFont></a:fontScheme>
</a:themeElements></a:theme>`;
const theme: OfficeTheme = parseTheme(parseXml(THEME_XML));

describe('formatValue', () => {
  const accounting = '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)';

  it('splits accounting formats at the fill code', () => {
    expect(formatValue(55, accounting)).toEqual({
      text: ' $55.00 ',
      fill: { left: ' $', right: '55.00 ' },
    });
    expect(formatValue(-2.5, accounting).fill).toEqual({ left: ' $', right: '(2.50)' });
    expect(formatValue(0, accounting).fill).toEqual({ left: ' $', right: '- ' });
    // The plain formatter never leaks the marker.
    expect(formatNumber(55, accounting)).toBe(' $55.00 ');
    expect(formatText('abc', accounting)).toBe(' abc ');
  });

  it('reports colour codes and leaves other formats alone', () => {
    expect(formatValue(-1.26, '0.0;[Red]-0.0')).toEqual({ text: '-1.3', color: '#ff0000' });
    expect(formatValue(1.26, '[Blue]0.0;[Red]-0.0')).toEqual({ text: '1.3', color: '#0000ff' });
    expect(formatValue(45000, '[Green]yyyy-mm-dd')).toEqual({
      text: '2023-03-15',
      color: '#008000',
    });
    expect(formatValue(3, '[Magenta]"n/a"')).toEqual({ text: 'n/a', color: '#ff00ff' });
    expect(formatValue(1.5, 'General')).toEqual({ text: '1.5' });
    expect(formatValue(Infinity, '0')).toEqual({ text: 'Infinity' });
    // Unknown bracket codes (conditions, locales) are ignored.
    expect(formatValue(5, '[>3]0.0')).toEqual({ text: '5.0' });
  });
});

const STYLES_XML = `<styleSheet>
  <numFmts><numFmt numFmtId="170" formatCode="0.000"/><numFmt numFmtId="bad"/></numFmts>
  <fonts>
    <font><sz val="11"/><color theme="1"/><name val="Calibri"/><scheme val="minor"/></font>
    <font><b/><i val="0"/><strike/><u val="none"/><sz val="14"/><color theme="4" tint="-0.25"/><name val="Arial"/></font>
    <font><u/><color indexed="10"/><name val="bad;font{}"/><scheme val="major"/></font>
    <font><color indexed="64"/><sz val="9999"/></font>
    <font><color indexed="65"/></font>
    <font><color indexed="1"/></font>
    <font><color indexed="999"/></font>
    <font><color auto="1"/></font>
  </fonts>
  <fills>
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor theme="4" tint="0.8"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="mediumGray"><fgColor rgb="FF000000"/><bgColor rgb="FFFFFFFF"/></patternFill></fill>
    <fill><gradientFill degree="90"><stop position="0"><color rgb="FFFF0000"/></stop><stop position="1"><color rgb="FF0000FF"/></stop></gradientFill></fill>
    <fill><gradientFill><stop position="0"><color rgb="FF00FF00"/></stop></gradientFill></fill>
    <fill><gradientFill/></fill>
    <fill/>
    <fill><patternFill patternType="darkTrellis"/></fill>
    <fill><patternFill><fgColor rgb="FF112233"/></patternFill></fill>
  </fills>
  <borders>
    <border><left/><right/><top/><bottom/></border>
    <border>
      <start style="hair"><color rgb="FF111111"/></start><end style="mediumDashed"/>
      <top style="thick"><color theme="2" tint="-0.5"/></top><bottom style="double"><color rgb="FF00FF00"/></bottom>
    </border>
  </borders>
  <cellXfs>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="170" fontId="1" fillId="2" borderId="1" applyFont="1" applyFill="1" applyBorder="1"><alignment horizontal="right" vertical="center" wrapText="1" indent="2"/></xf>
    <xf numFmtId="14" fontId="2" fillId="0" borderId="0"><alignment indent="900"/></xf>
    <xf numFmtId="999" fontId="99" fillId="99" borderId="99"/>
  </cellXfs>
  <dxfs>
    <dxf><font><strike/><color rgb="FF808080"/></font></dxf>
    <dxf><fill><patternFill patternType="solid"><fgColor rgb="FF111111"/><bgColor rgb="FFF2F2F2"/></patternFill></fill></dxf>
    <dxf><fill><patternFill patternType="none"/></fill><border><left/><top style="dotted"><color rgb="FFAA0000"/></top><horizontal style="thin"/></border><numFmt numFmtId="171" formatCode="0%"/></dxf>
    <dxf><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>
  </dxfs>
  <tableStyles defaultTableStyle="Mine">
    <tableStyle name="Mine"><tableStyleElement type="wholeTable" dxfId="2"/><tableStyleElement type="firstRowStripe" size="2" dxfId="1"/><tableStyleElement type="headerRow" dxfId="77"/><tableStyleElement dxfId="0"/></tableStyle>
    <tableStyle/>
  </tableStyles>
  <colors><indexedColors><rgbColor rgb="00000000"/><rgbColor rgb="00ABCDEF"/></indexedColors></colors>
</styleSheet>`;

describe('parseStyles', () => {
  const styles = parseStyles(parseXml(STYLES_XML), theme);

  it('resolves theme, tinted, indexed and custom palette colours', () => {
    expect(themeColorAt(theme, 0)).toBe('#ffffff');
    expect(themeColorAt(theme, 1)).toBe('#000000');
    expect(themeColorAt(theme, 3)).toBe('#1f497d');
    expect(themeColorAt(theme, 99)).toBeUndefined();
    // Without a theme part the Office defaults apply.
    expect(themeColorAt(parseTheme(null), 4)).toBe('#4472c4');
    const fonts = styles.cellXfs.map((xf) => xf.font);
    expect(fonts[0]).toEqual({ color: '#000000', name: 'Corbel', size: 11 });
    expect(fonts[1]).toMatchObject({ bold: true, italic: false, strike: true, underline: false });
    expect(fonts[1]!.size).toBe(14);
    expect(fonts[1]!.name).toBe('Arial');
    // accent1 darkened by a quarter keeps its hue.
    expect(fonts[1]!.color).toMatch(/^#[0-9a-f]{6}$/);
    expect(fonts[1]!.color).not.toBe('#4f81bd');
    expect(fonts[2]).toEqual({ underline: true, color: '#FF0000', name: 'Impact' });
    expect(styles.font(parseXml('<f><color indexed="64"/><sz val="9999"/></f>'))).toEqual({
      color: '#000000',
    });
    expect(styles.font(parseXml('<f><color indexed="65"/></f>')).color).toBe('#ffffff');
    expect(styles.font(parseXml('<f><color indexed="1"/></f>')).color).toBe('#ABCDEF');
    expect(styles.font(parseXml('<f><color indexed="999"/></f>')).color).toBeUndefined();
    expect(styles.font(parseXml('<f><color auto="1"/><rFont val="Georgia"/></f>'))).toEqual({
      name: 'Georgia',
    });
    expect(styles.font(null)).toEqual({});
    expect(styles.color(null)).toBeUndefined();
    expect(styles.defaultFont.size).toBe(11);
  });

  it('reads fills, borders, alignment and number formats', () => {
    const [base, rich, dated, broken] = styles.cellXfs;
    expect(base).toMatchObject({
      numberFormat: 'General',
      fill: undefined,
      explicitFont: false,
      explicitFill: false,
      explicitBorder: false,
    });
    expect(rich!.numberFormat).toBe('0.000');
    expect(rich!.fill).toMatch(/^#[0-9a-f]{6}$/);
    expect(rich).toMatchObject({ explicitFont: true, explicitFill: true, explicitBorder: true });
    expect(rich!.align).toEqual({ horizontal: 'right', vertical: 'center', wrap: true, indent: 2 });
    expect(rich!.border.left).toEqual({ width: 1, style: 'dotted', color: '#111111' });
    expect(rich!.border.right).toEqual({ width: 2, style: 'dashed', color: '#000000' });
    expect(rich!.border.top).toMatchObject({ width: 3, style: 'solid' });
    expect(rich!.border.bottom).toEqual({ width: 3, style: 'double', color: '#00FF00' });
    expect(dated!.numberFormat).toBe('m/d/yyyy');
    expect(dated!.align.indent).toBe(50);
    expect(broken).toMatchObject({
      numberFormat: 'General',
      font: {},
      fill: undefined,
      border: {},
    });
  });

  it('approximates pattern and gradient fills', () => {
    const fill = (index: number): string | undefined =>
      parseStyles(
        parseXml(
          STYLES_XML.replace(
            '<xf numFmtId="999" fontId="99" fillId="99" borderId="99"/>',
            `<xf fillId="${index}"/>`,
          ),
        ),
        theme,
      ).cellXfs[3]!.fill;
    expect(fill(1)).toBe('#dfdfdf');
    expect(fill(3)).toBe('#808080');
    expect(fill(4)).toBe('linear-gradient(180deg, #FF0000, #0000FF)');
    expect(fill(5)).toBe('#00FF00');
    expect(fill(6)).toBeUndefined();
    expect(fill(7)).toBeUndefined();
    expect(fill(8)).toBe('#b3b3b3');
    expect(fill(9)).toBeUndefined();
  });

  it('reads differential formats and table styles', () => {
    expect(styles.dxfs[0]).toEqual({ font: { strike: true, color: '#808080' } });
    // Differential solid fills take the background colour.
    expect(styles.dxfs[1]).toEqual({ fill: '#F2F2F2' });
    expect(styles.dxfs[2]).toEqual({
      fill: null,
      border: {
        left: null,
        top: { width: 1, style: 'dotted', color: '#AA0000' },
        horizontal: { width: 1, style: 'solid', color: '#000000' },
      },
      numberFormat: '0%',
    });
    expect(styles.dxfs[3]).toEqual({ fill: '#FFC7CE' });
    expect(styles.defaultTableStyle).toBe('Mine');
    const mine = styles.tableStyles.get('Mine')!;
    expect([...mine.elements.keys()]).toEqual(['wholeTable', 'firstRowStripe']);
    expect(mine.elements.get('firstRowStripe')!.size).toBe(2);
  });

  it('copes with a missing style sheet', () => {
    const empty = parseStyles(null, theme);
    expect(empty.cellXfs).toEqual([]);
    expect(empty.defaultFont).toEqual({});
    expect(empty.tableStyles.size).toBe(0);
  });

  it('parses extension dxfs and merges fonts into grid styles', () => {
    const dxf = parseInlineDxf(
      parseXml(
        '<dxf><font><b/><i/><strike val="0"/><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>',
      ),
      styles,
    );
    expect(dxf).toEqual({
      font: { bold: true, italic: true, strike: false, color: '#9C0006' },
      fill: '#FFC7CE',
    });
    expect(parseInlineDxf(parseXml('<dxf><fill><patternFill/></fill></dxf>'), styles)).toEqual({});
    const style = { bold: true };
    applyFont(style, undefined);
    applyFont(style, { bold: false, italic: true, underline: true, strike: true, size: 12 });
    expect(style).toEqual({
      bold: false,
      italic: true,
      underline: true,
      strike: true,
      fontSize: 16,
    });
  });
});

describe('builtinTableStyle', () => {
  it('generates every family from the theme accents', () => {
    expect(builtinTableStyle('Custom', theme)).toBeUndefined();
    const medium2 = builtinTableStyle('TableStyleMedium2', theme)!;
    expect(medium2.elements.get('headerRow')!.dxf).toEqual({
      font: { bold: true, color: '#ffffff' },
      fill: '#4f81bd',
    });
    expect(medium2.elements.get('firstRowStripe')!.dxf.fill).toMatch(/^#[0-9a-f]{6}$/);
    expect(medium2.elements.get('wholeTable')!.dxf.border).toHaveProperty('horizontal');

    for (const name of [
      'TableStyleLight1',
      'TableStyleLight2',
      'TableStyleLight9',
      'TableStyleLight16',
      'TableStyleMedium1',
      'TableStyleMedium9',
      'TableStyleMedium15',
      'TableStyleMedium16',
      'TableStyleMedium23',
      'TableStyleDark1',
      'TableStyleDark3',
    ]) {
      const style = builtinTableStyle(name, theme)!;
      expect(style.elements.has('wholeTable'), name).toBe(true);
      expect(style.elements.has('headerRow'), name).toBe(true);
      expect(style.elements.get('firstColumn')!.dxf).toEqual({ font: { bold: true } });
    }
    // The first style of each family is neutral rather than accent coloured.
    expect(
      builtinTableStyle('TableStyleLight1', theme)!.elements.get('wholeTable')!.dxf.font,
    ).toEqual({ color: '#000000' });
    expect(builtinTableStyle('TableStyleDark1', theme)!.elements.get('headerRow')!.dxf.fill).toBe(
      '#000000',
    );
  });
});

describe('compileFormula', () => {
  const cells: Record<string, FormulaValue> = {
    '0:0': 5,
    '0:1': 'Apple',
    '1:0': 12,
    '1:1': '',
    '2:0': true,
    '3:1': ' padded  text ',
  };
  const read = (row: number, col: number): FormulaValue => cells[`${row}:${col}`] ?? null;
  const run = (formula: string, rowShift = 0, colShift = 0): FormulaValue =>
    compileFormula(formula)!(read, rowShift, colShift, rowShift, colShift);

  it('evaluates arithmetic, comparison and text operators', () => {
    expect(run('=1+2*3')).toBe(7);
    expect(run('(1+2)*3')).toBe(9);
    expect(run('2^3')).toBe(8);
    expect(run('10/4')).toBe(2.5);
    expect(run('-A1+10')).toBe(5);
    expect(run('+A1')).toBe(5);
    expect(run('50%')).toBe(0.5);
    expect(run('1.5e2')).toBe(150);
    expect(run('"a"&"b"&1&TRUE')).toBe('ab1TRUE');
    expect(run('"say ""hi"""')).toBe('say "hi"');
    expect(run('A1>3')).toBe(true);
    expect(run('A1>=5')).toBe(true);
    expect(run('A1<5')).toBe(false);
    expect(run('A1<=4')).toBe(false);
    expect(run('A1<>5')).toBe(false);
    expect(run('B1="apple"')).toBe(true);
    expect(run('"b">"a"')).toBe(true);
    expect(run('"a"<"b"')).toBe(true);
    expect(run('TRUE=TRUE')).toBe(true);
    expect(run('FALSE<TRUE')).toBe(true);
    // Empty cells compare as zero or as an empty string.
    expect(run('Z9=0')).toBe(true);
    expect(run('Z9=""')).toBe(true);
    // Numbers sort before text, text before booleans.
    expect(run('1<"a"')).toBe(true);
    expect(run('"a"<TRUE')).toBe(true);
    expect(run('A3+1')).toBe(2);
    expect(run('Z9+B2+1')).toBe(1);
    expect(run('"4"*2')).toBe(8);
  });

  it('moves relative references and keeps absolute ones', () => {
    expect(run('A1', 1, 0)).toBe(12);
    expect(run('$A$1', 1, 1)).toBe(5);
    expect(run('$A1', 1, 5)).toBe(12);
    expect(run('A$1', 5, 1)).toBe('Apple');
    expect(run('ROW()', 4, 2)).toBe(5);
    expect(run('COLUMN()', 4, 2)).toBe(3);
    expect(run('ROW(A1)', 2, 0)).toBe(3);
    expect(run('ROW($A$1)', 2, 0)).toBe(1);
    expect(run('COLUMN(B1)', 0, 2)).toBe(4);
    expect(run('COLUMN($B1)', 0, 2)).toBe(2);
    expect(run('MOD(ROW(),2)=0', 1, 0)).toBe(true);
  });

  it('supports the common functions', () => {
    expect(run('LEN(B1)')).toBe(5);
    expect(run('LEN(TRIM(B4))')).toBe(11);
    expect(run('UPPER(B1)&LOWER(B1)')).toBe('APPLEapple');
    expect(run('LEFT(B1,2)&RIGHT(B1,2)&LEFT(B1)&RIGHT(B1)')).toBe('ApleAe');
    expect(run('RIGHT(B1,0)')).toBe('');
    expect(run('AND(A1>1,A2>1)')).toBe(true);
    expect(run('OR(A1>100,A2>100)')).toBe(false);
    expect(run('NOT(A1>100)')).toBe(true);
    expect(run('NOT("true")')).toBe(false);
    expect(run('IF(A1>1,"big","small")')).toBe('big');
    expect(run('IF(A1>100,"big","small")')).toBe('small');
    expect(run('IF(A1>1)')).toBe(true);
    expect(run('IF(A1>100,1)')).toBe(false);
    expect(run('ISBLANK(Z9)')).toBe(true);
    expect(run('ISNUMBER(A1)')).toBe(true);
    expect(run('ISTEXT(B1)')).toBe(true);
    expect(run('ABS(0-3)')).toBe(3);
    expect(run('MOD(0-1,3)')).toBe(2);
    expect(run('TODAY()>40000')).toBe(true);
    expect(run('LEN(A3)')).toBe(4);
  });

  it('declines what it does not understand', () => {
    for (const formula of [
      'SUM(A1:A3)',
      'Sheet2!A1',
      'Table1[Column]',
      '"unterminated',
      'MyName',
      '(1+2',
      '1 2',
      'LEN(1;2)',
      '1+',
      '{1,2}',
      '.',
    ]) {
      expect(compileFormula(formula), formula).toBeNull();
    }
    // Problems that only show at evaluation time give null for that cell.
    expect(run('VLOOKUP(1,2)')).toBeNull();
    expect(run('1/0')).toBeNull();
    expect(run('MOD(1,0)')).toBeNull();
    expect(run('B1*2')).toBeNull();
    expect(run('LEN()')).toBeNull();
    expect(run('ROW(1)')).toBeNull();
    expect(run('COLUMN(1)')).toBeNull();
  });

  it('defines truthiness like conditional formatting', () => {
    expect(isTruthy(true)).toBe(true);
    expect(isTruthy(2)).toBe(true);
    expect(isTruthy(0)).toBe(false);
    expect(isTruthy('TRUE')).toBe(false);
    expect(isTruthy(null)).toBe(false);
  });
});

describe('parseRange', () => {
  it('parses cells, areas, whole columns and whole rows', () => {
    expect(parseRange('B2:D9', 100, 100)).toEqual({ top: 1, left: 1, bottom: 8, right: 3 });
    expect(parseRange('$C$4', 100, 100)).toEqual({ top: 3, left: 2, bottom: 3, right: 2 });
    expect(parseRange('D9:B2', 100, 100)).toEqual({ top: 1, left: 1, bottom: 8, right: 3 });
    expect(parseRange('A:C', 50, 100)).toEqual({ top: 0, left: 0, bottom: 49, right: 2 });
    expect(parseRange('3:8', 50, 10)).toEqual({ top: 2, left: 0, bottom: 7, right: 9 });
    expect(parseRange('B', 5, 5)).toEqual({ top: 0, left: 1, bottom: 4, right: 1 });
    expect(parseRange('7', 9, 5)).toEqual({ top: 6, left: 0, bottom: 6, right: 4 });
    // Clipped to the extent of the data.
    expect(parseRange('A1:Z99', 3, 2)).toEqual({ top: 0, left: 0, bottom: 2, right: 1 });
    expect(parseRange('K20', 3, 2)).toBeNull();
    expect(parseRange('', 3, 2)).toBeNull();
    expect(parseRange('not a range', 3, 2)).toBeNull();
    expect(parseRanges('A1 bogus B2:B3', 9, 9)).toHaveLength(2);
  });
});

describe('evaluateConditionalFormats', () => {
  const styles = parseStyles(parseXml(STYLES_XML), theme);
  /** Column A holds 1…10, column B holds words, C1 an error. */
  const data: (CfCell | undefined)[][] = Array.from({ length: 10 }, (_unused, row) => [
    { text: String(row + 1), value: row + 1 },
    {
      text: ['apple', 'Banana', 'apple', '', 'cherry pie', ' ', 'date', 'egg', 'fig', 'grape'][
        row
      ]!,
    },
  ]);
  data[0]![2] = { text: '#N/A', error: true };
  const sheet = {
    rowCount: 10,
    colCount: 3,
    cell: (row: number, col: number) => data[row]?.[col],
  };
  const evaluate = (rules: string, extension = ''): ReturnType<typeof evaluateConditionalFormats> =>
    evaluateConditionalFormats(
      parseXml(`<worksheet xmlns:x14="x" xmlns:xm="m">${rules}${extension}</worksheet>`),
      styles,
      sheet,
    );
  const rowsOf = (effects: Map<number, unknown>, col = 0): number[] =>
    [...effects.keys()]
      .filter((key) => key % CF_STRIDE === col)
      .map((key) => Math.floor(key / CF_STRIDE) + 1)
      .sort((a, b) => a - b);
  const rule = (sqref: string, attrs: string, body = ''): string =>
    `<conditionalFormatting sqref="${sqref}"><cfRule ${attrs} dxfId="0">${body}</cfRule></conditionalFormatting>`;

  it('returns nothing when there are no rules', () => {
    expect(evaluate('').size).toBe(0);
    expect(evaluate(rule('Z99', 'type="expression"', '<formula>TRUE</formula>')).size).toBe(0);
  });

  it('applies cellIs rules with every operator', () => {
    const cellIs = (operator: string, ...operands: string[]): number[] =>
      rowsOf(
        evaluate(
          rule(
            'A1:A10',
            `type="cellIs" operator="${operator}"`,
            operands.map((operand) => `<formula>${operand}</formula>`).join(''),
          ),
        ),
      );
    expect(cellIs('greaterThan', '8')).toEqual([9, 10]);
    expect(cellIs('greaterThanOrEqual', '9')).toEqual([9, 10]);
    expect(cellIs('lessThan', '3')).toEqual([1, 2]);
    expect(cellIs('lessThanOrEqual', '2')).toEqual([1, 2]);
    expect(cellIs('equal', '4')).toEqual([4]);
    expect(cellIs('notEqual', '4')).toHaveLength(9);
    expect(cellIs('between', '3', '5')).toEqual([3, 4, 5]);
    expect(cellIs('between', '5', '3')).toEqual([3, 4, 5]);
    expect(cellIs('notBetween', '2', '9')).toEqual([1, 10]);
    expect(cellIs('between', '3')).toEqual([]);
    expect(cellIs('nonsense', '3')).toEqual([]);
    // A text operand never equals a number, so only "not equal" matches.
    expect(cellIs('equal', '"x"')).toEqual([]);
    expect(cellIs('notEqual', '"x"')).toHaveLength(10);
    // Operands may reference other cells, relative to the first cell of the range.
    expect(cellIs('greaterThan', '$A$5')).toEqual([6, 7, 8, 9, 10]);
    expect(cellIs('equal')).toEqual([]);
    expect(cellIs('equal', 'SUM(A1:A2)')).toEqual([]);
    const words = rowsOf(
      evaluate(rule('B1:B10', 'type="cellIs" operator="equal"', '<formula>"APPLE"</formula>')),
      1,
    );
    expect(words).toEqual([1, 3]);
    const sorted = rowsOf(
      evaluate(rule('B1:B10', 'type="cellIs" operator="greaterThan"', '<formula>"f"</formula>')),
      1,
    );
    expect(sorted).toEqual([9, 10]);
  });

  it('applies text, blank, error and duplicate rules', () => {
    const on = (attrs: string, range = 'B1:B10', col = 1): number[] =>
      rowsOf(evaluate(rule(range, attrs)), col);
    expect(on('type="containsText" text="AP"')).toEqual([1, 3, 10]);
    expect(on('type="notContainsText" text="a"')).toEqual([4, 5, 6, 8, 9]);
    expect(on('type="beginsWith" text="b"')).toEqual([2]);
    expect(on('type="endsWith" text="e"')).toEqual([1, 3, 5, 7, 10]);
    expect(on('type="containsText" text=""')).toEqual([]);
    expect(on('type="containsBlanks"')).toEqual([4, 6]);
    expect(on('type="notContainsBlanks"')).toHaveLength(8);
    expect(on('type="containsErrors"', 'C1:C3', 2)).toEqual([1]);
    expect(on('type="notContainsErrors"', 'C1:C3', 2)).toEqual([2, 3]);
    expect(on('type="duplicateValues"')).toEqual([1, 3]);
    expect(on('type="uniqueValues"')).toEqual([2, 5, 6, 7, 8, 9, 10]);
    expect(on('type="timePeriod"')).toEqual([]);
  });

  it('applies ranking and average rules', () => {
    const on = (attrs: string): number[] => rowsOf(evaluate(rule('A1:A10', attrs)));
    expect(on('type="top10" rank="3"')).toEqual([8, 9, 10]);
    expect(on('type="top10" rank="2" bottom="1"')).toEqual([1, 2]);
    expect(on('type="top10" rank="20" percent="1"')).toEqual([9, 10]);
    expect(on('type="top10"')).toHaveLength(10);
    expect(on('type="aboveAverage"')).toEqual([6, 7, 8, 9, 10]);
    expect(on('type="aboveAverage" aboveAverage="0"')).toEqual([1, 2, 3, 4, 5]);
    expect(rowsOf(evaluate(rule('B1:B10', 'type="top10" rank="3"')), 1)).toEqual([]);
    expect(rowsOf(evaluate(rule('B1:B10', 'type="aboveAverage"')), 1)).toEqual([]);
    const equal = evaluate(
      `<conditionalFormatting sqref="A1:A3"><cfRule type="aboveAverage" aboveAverage="0" equalAverage="1" dxfId="0"/></conditionalFormatting>`,
    );
    expect(rowsOf(equal)).toEqual([1, 2]);
  });

  it('evaluates expressions per cell and orders rules by priority', () => {
    const effects = evaluate(
      rule('A1:B10', 'type="expression" priority="1"', '<formula>MOD($A1,2)=0</formula>') +
        `<conditionalFormatting sqref="A1:A10"><cfRule type="expression" priority="2" dxfId="1"><formula>A1&gt;0</formula></cfRule><cfRule type="expression" priority="3"><formula>TRUE</formula></cfRule></conditionalFormatting>`,
    );
    expect(rowsOf(effects, 1)).toEqual([2, 4, 6, 8, 10]);
    // Priority 1 is applied last so that it wins.
    expect(effects.get(1 * CF_STRIDE)!.dxfs).toEqual([styles.dxfs[1], styles.dxfs[0]]);
    expect(effects.get(0)!.dxfs).toEqual([styles.dxfs[1]]);
    expect(
      rowsOf(evaluate(rule('A1:A3', 'type="expression"', '<formula>SUM(A:A)</formula>'))),
    ).toEqual([]);
  });

  it('computes colour scales', () => {
    const effects = evaluate(
      `<conditionalFormatting sqref="A1:B10"><cfRule type="colorScale" priority="1"><colorScale><cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/><color rgb="FFFF0000"/><color rgb="FFFFFF00"/><color rgb="FF00FF00"/></colorScale></cfRule></conditionalFormatting>`,
    );
    expect(effects.get(0)!.background).toBe('#ff0000');
    expect(effects.get(9 * CF_STRIDE)!.background).toBe('#00ff00');
    expect(effects.get(4 * CF_STRIDE)!.background).toBe('#ffe300');
    // Text cells are left alone.
    expect(effects.has(1)).toBe(false);
    const broken = evaluate(
      `<conditionalFormatting sqref="A1:A10"><cfRule type="colorScale"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFFF0000"/></colorScale></cfRule><cfRule type="colorScale"/></conditionalFormatting>`,
    );
    expect(broken.size).toBe(0);
    const flat = evaluate(
      `<conditionalFormatting sqref="A1"><cfRule type="colorScale"><colorScale><cfvo type="num" val="5"/><cfvo type="num" val="5"/><color rgb="FF000000"/><color rgb="FFFFFFFF"/></colorScale></cfRule></conditionalFormatting>`,
    );
    expect(flat.get(0)!.background).toBe('#000000');
  });

  it('computes data bars', () => {
    const effects = evaluate(
      `<conditionalFormatting sqref="A1:A10"><cfRule type="dataBar"><dataBar showValue="0"><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar></cfRule></conditionalFormatting>`,
    );
    expect(effects.get(9 * CF_STRIDE)).toEqual({
      dxfs: [],
      bar: { color: '#638EC6', start: 0, end: 1 },
      hideText: true,
    });
    expect(effects.get(4 * CF_STRIDE)!.bar!.end).toBeCloseTo(0.5);
    const fixed = evaluate(
      `<conditionalFormatting sqref="A1:A10"><cfRule type="dataBar"><dataBar><cfvo type="num" val="5"/><cfvo type="percent" val="100"/><color rgb="FF00FF00"/></dataBar></cfRule></conditionalFormatting>`,
    );
    expect(fixed.get(0)!.bar!.end).toBe(0.02);
    expect(fixed.get(0)!.hideText).toBeUndefined();
    const none = evaluate(
      `<conditionalFormatting sqref="A1:A10"><cfRule type="dataBar"><dataBar><cfvo type="min"/><color rgb="FF00FF00"/></dataBar></cfRule><cfRule type="dataBar"><dataBar><cfvo type="min"/><cfvo type="max"/></dataBar></cfRule><cfRule type="dataBar"><dataBar><cfvo type="num" val="9"/><cfvo type="num" val="1"/><color rgb="FF00FF00"/></dataBar></cfRule></conditionalFormatting>`,
    );
    expect(none.size).toBe(0);
  });

  it('picks icons from classic and extended icon sets', () => {
    const classic = evaluate(
      `<conditionalFormatting sqref="A1:A10"><cfRule type="iconSet"><iconSet iconSet="3Arrows"><cfvo type="percent" val="0"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/></iconSet></cfRule></conditionalFormatting>`,
    );
    expect(classic.get(0)!.icon).toMatchObject({ shape: 'down' });
    expect(classic.get(4 * CF_STRIDE)!.icon).toMatchObject({ shape: 'right' });
    expect(classic.get(9 * CF_STRIDE)!.icon).toMatchObject({ shape: 'up' });
    expect(classic.get(0)!.hideText).toBeUndefined();

    const reversed = evaluate(
      `<conditionalFormatting sqref="A1:A10"><cfRule type="iconSet"><iconSet reverse="1" showValue="0"><cfvo type="num" val="0"/><cfvo type="num" val="4" gte="0"/><cfvo type="formula" val="8"/></iconSet></cfRule><cfRule type="iconSet"/><cfRule type="iconSet"><iconSet iconSet="Unknown"><cfvo type="num" val="0"/></iconSet></cfRule></conditionalFormatting>`,
    );
    // Reversed traffic lights: low values are green; 4 is not above the exclusive threshold.
    expect(reversed.get(3 * CF_STRIDE)!.icon).toMatchObject({ shape: 'circle', color: '#3c9d4e' });
    expect(reversed.get(4 * CF_STRIDE)!.icon).toMatchObject({ color: '#e9b21c' });
    expect(reversed.get(9 * CF_STRIDE)!.icon).toMatchObject({ color: '#d9453c' });
    expect(reversed.get(0)!.hideText).toBe(true);

    const extended = evaluate(
      '',
      `<extLst><ext><x14:conditionalFormattings>
        <x14:conditionalFormatting><x14:cfRule type="iconSet" priority="5"><x14:iconSet iconSet="3Symbols" showValue="0" custom="1"><x14:cfvo type="percent"><xm:f>0</xm:f></x14:cfvo><x14:cfvo type="num" gte="0"><xm:f>0</xm:f></x14:cfvo><x14:cfvo type="num"><xm:f>6</xm:f></x14:cfvo><x14:cfIcon iconSet="NoIcons" iconId="0"/><x14:cfIcon iconSet="3Flags" iconId="1"/><x14:cfIcon iconSet="3Symbols2" iconId="2"/></x14:iconSet></x14:cfRule><xm:sqref>A1:A10</xm:sqref></x14:conditionalFormatting>
        <x14:conditionalFormatting><x14:cfRule type="dataBar"/><xm:sqref>A1:A10</xm:sqref></x14:conditionalFormatting>
      </x14:conditionalFormattings></ext></extLst>`,
    );
    expect(extended.get(0)!.icon).toEqual({ shape: 'flag', color: '#e9b21c' });
    expect(extended.get(5 * CF_STRIDE)!.icon).toEqual({ shape: 'check', color: '#3c9d4e' });
    expect(extended.get(5 * CF_STRIDE)!.hideText).toBe(true);
  });

  it('uses dxfs declared inside extension rules and other threshold kinds', () => {
    const effects = evaluate(
      `<conditionalFormatting sqref="A1:A10"><cfRule type="cellIs" operator="equal"><formula>1</formula><dxf><font><b/></font></dxf></cfRule><cfRule type="iconSet"><iconSet iconSet="5Rating"><cfvo type="autoMin"/><cfvo type="percentile" val="25"/><cfvo type="num" val="5"/><cfvo type="autoMax"/><cfvo type="max"/></iconSet></cfRule><cfRule type="iconSet"><iconSet iconSet="4Arrows"><cfvo type="min"/><cfvo type="num"/><cfvo type="num" val="x"/><cfvo type="num" val="9"/></iconSet></cfRule></conditionalFormatting>`,
    );
    expect(effects.get(0)!.dxfs).toEqual([{ font: { bold: true } }]);
    expect(effects.get(9 * CF_STRIDE)!.icon).toMatchObject({ shape: 'up' });
  });
});

const CHART_NS =
  'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
const strRef = (values: string[]): string =>
  `<c:strRef><c:strCache><c:ptCount val="${values.length}"/>${values
    .map((value, index) => `<c:pt idx="${index}"><c:v>${value}</c:v></c:pt>`)
    .join('')}</c:strCache></c:strRef>`;
const numRef = (values: (number | null)[], format = ''): string =>
  `<c:numRef><c:numCache>${format ? `<c:formatCode>${format}</c:formatCode>` : ''}<c:ptCount val="${values.length}"/>${values
    .map((value, index) =>
      value === null ? '' : `<c:pt idx="${index}"><c:v>${value}</c:v></c:pt>`,
    )
    .join('')}</c:numCache></c:numRef>`;
const series = (name: string, values: (number | null)[], extra = ''): string =>
  `<c:ser><c:tx>${strRef([name])}</c:tx>${extra}<c:cat>${strRef(['Q1', 'Q2', 'Q3'])}</c:cat><c:val>${numRef(values, '0.0')}</c:val></c:ser>`;
const chartXml = (plot: string, extra = '', space = ''): string =>
  `<c:chartSpace ${CHART_NS}><c:chart>${extra}<c:plotArea>${plot}</c:plotArea></c:chart>${space}</c:chartSpace>`;
const AXES =
  '<c:catAx><c:delete val="0"/><c:axPos val="b"/></c:catAx><c:valAx><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/><c:majorTickMark val="out"/></c:valAx>';
const parse = (xml: string): ChartModel => parseChart(parseXml(xml), theme)!;

describe('drawingColor', () => {
  const color = (xml: string): string | undefined =>
    drawingColor(parseXml(`<a:solidFill xmlns:a="a">${xml}</a:solidFill>`), theme);

  it('resolves colour kinds and modifiers', () => {
    expect(color('<a:srgbClr val="FF0000"/>')).toBe('#ff0000');
    expect(color('<a:schemeClr val="accent1"/>')).toBe('#4f81bd');
    expect(color('<a:schemeClr val="tx1"/>')).toBe('#000000');
    expect(color('<a:schemeClr val="bg1"/>')).toBe('#ffffff');
    expect(color('<a:schemeClr val="tx2"/>')).toBe('#1f497d');
    expect(color('<a:schemeClr val="bg2"/>')).toBe('#eeece1');
    expect(color('<a:sysClr val="windowText" lastClr="112233"/>')).toBe('#112233');
    expect(color('<a:sysClr val="windowText"/>')).toBe('#000000');
    expect(color('<a:prstClr val="red"/>')).toBe('#ff0000');
    expect(color('<a:prstClr val="chartreuse"/>')).toBeUndefined();
    expect(color('<a:schemeClr val="nope"/>')).toBeUndefined();
    expect(color('<a:other/>')).toBeUndefined();
    expect(drawingColor(null, theme)).toBeUndefined();
    // Luminance changes keep grey colours grey.
    expect(color('<a:srgbClr val="808080"><a:lumMod val="50000"/></a:srgbClr>')).toBe('#404040');
    expect(
      color('<a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr>'),
    ).toBe('#595959');
    expect(color('<a:srgbClr val="000000"><a:tint val="50000"/></a:srgbClr>')).toBe('#808080');
    expect(color('<a:srgbClr val="FFFFFF"><a:shade val="50000"/></a:srgbClr>')).toBe('#808080');
    expect(color('<a:srgbClr val="FF0000"><a:satMod val="0"/><a:alpha val="x"/></a:srgbClr>')).toBe(
      '#808080',
    );
    // Darkening a colour keeps its hue: red stays the strongest channel.
    const dark = color('<a:srgbClr val="CC6633"><a:lumMod val="50000"/></a:srgbClr>')!;
    expect(parseInt(dark.slice(1, 3), 16)).toBeGreaterThan(parseInt(dark.slice(5, 7), 16));
    expect(color('<a:srgbClr val="33CC66"><a:lumMod val="110000"/></a:srgbClr>')).toMatch(/^#/);
    expect(color('<a:srgbClr val="3366CC"><a:lumMod val="90000"/></a:srgbClr>')).toMatch(/^#/);
  });
});

describe('parseChart', () => {
  it('returns null without a drawable chart', () => {
    expect(parseChart(null, theme)).toBeNull();
    expect(parseChart(parseXml(chartXml('<c:surfaceChart/>')), theme)).toBeNull();
    expect(parseChart(parseXml(chartXml('<c:barChart/>')), theme)).toBeNull();
  });

  it('reads bar charts with axes, legend, title and colours', () => {
    const model = parse(
      chartXml(
        `<c:barChart><c:barDir val="bar"/><c:grouping val="stacked"/><c:varyColors val="0"/>
          ${series('North', [1, 2, 3], '<c:spPr><a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:ln w="19050"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></c:spPr><c:dPt><c:idx val="1"/><c:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></c:spPr></c:dPt><c:dPt><c:idx val="2"/><c:spPr><a:noFill/></c:spPr></c:dPt>')}
          ${series('South', [4, null, 6], '<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>')}
          <c:ser><c:val>${numRef([1])}</c:val><c:spPr><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="00FF00"/></a:gs></a:gsLst></a:gradFill><a:ln w="1"/></c:spPr></c:ser>
          <c:dLbls><c:showVal val="1"/></c:dLbls><c:gapWidth val="50"/>
        </c:barChart>
        <c:catAx><c:delete val="1"/></c:catAx>
        <c:valAx><c:scaling><c:min val="0"/><c:max val="10"/></c:scaling><c:delete val="0"/><c:majorGridlines><c:spPr><a:ln><a:solidFill><a:srgbClr val="CCCCCC"/></a:solidFill></a:ln></c:spPr></c:majorGridlines><c:numFmt formatCode="0%" sourceLinked="0"/><c:majorTickMark val="out"/><c:majorUnit val="2.5"/><c:spPr><a:ln><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:ln></c:spPr><c:txPr><a:p><a:pPr><a:defRPr sz="1200" b="1" i="1"><a:solidFill><a:srgbClr val="333333"/></a:solidFill></a:defRPr></a:pPr></a:p></c:txPr></c:valAx>
        <c:spPr><a:solidFill><a:srgbClr val="FAFAFA"/></a:solidFill><a:ln w="25400"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:ln></c:spPr>`,
        '<c:title><c:tx><c:rich><a:p><a:r><a:rPr sz="1800" b="1"/><a:t>Sales </a:t></a:r><a:r><a:t>by region</a:t></a:r></a:p></c:rich></c:tx></c:title><c:autoTitleDeleted val="0"/>',
        '<c:spPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln></c:spPr><c:txPr><a:p><a:pPr><a:defRPr sz="900"/></a:pPr></a:p></c:txPr>',
      ).replace('</c:plotArea>', '</c:plotArea><c:legend><c:legendPos val="b"/></c:legend>'),
    );
    expect(model.title).toBe('Sales by region');
    expect(model.titleText).toMatchObject({ size: 24, bold: true });
    expect(model.legend).toBe('b');
    expect(model.background).toBe('#ffffff');
    expect(model.border).toEqual({ color: '#d9d9d9', width: 1 });
    expect(model.plotBackground).toBe('#fafafa');
    expect(model.plotBorder!.width).toBeCloseTo(2.667, 2);
    expect(model.fontFamily).toBe('Corbel');
    const [group] = model.groups;
    expect(group).toMatchObject({
      kind: 'bar',
      horizontal: true,
      grouping: 'stacked',
      gapWidth: 50,
      showValues: true,
      varyColors: false,
    });
    const [north, south, third] = group!.series;
    expect(north).toMatchObject({
      name: 'North',
      categories: ['Q1', 'Q2', 'Q3'],
      values: [1, 2, 3],
      color: '#112233',
      formatCode: '0.0',
    });
    expect(north!.pointColors[1]).toBe('#ff0000');
    expect(north!.pointColors[2]).toBeUndefined();
    expect(north!.stroke).toEqual({ color: '#ffffff', width: 2 });
    // No fill: the automatic accent for the second series is used.
    expect(south).toMatchObject({ values: [4, null, 6], color: '#c0504d' });
    expect(south!.stroke).toBeUndefined();
    expect(third).toMatchObject({ name: 'Series 3', color: '#00ff00', categories: [] });
    expect(model.categoryAxis.deleted).toBe(true);
    expect(model.valueAxis).toMatchObject({
      deleted: false,
      min: 0,
      max: 10,
      majorUnit: 2.5,
      formatCode: '0%',
      gridlines: '#cccccc',
      line: '#000000',
      ticks: true,
      text: { size: 16, bold: true, italic: true, color: '#333333' },
    });
  });

  it('reads line, area, pie, doughnut and scatter charts', () => {
    const line = parse(
      chartXml(
        `<c:lineChart><c:grouping val="standard"/>
          ${series('A', [1, 2, 3], '<c:spPr><a:ln w="28575"><a:solidFill><a:srgbClr val="AA0000"/></a:solidFill></a:ln></c:spPr><c:marker><c:symbol val="circle"/></c:marker>')}
          ${series('B', [3, 2, 1], '<c:marker><c:symbol val="none"/></c:marker>')}
          <c:marker val="1"/></c:lineChart>${AXES}`,
        '<c:title/><c:autoTitleDeleted val="1"/>',
      ).replace('</c:plotArea>', '</c:plotArea><c:legend><c:legendPos val="tr"/></c:legend>'),
    );
    expect(line.title).toBeUndefined();
    expect(line.legend).toBe('r');
    expect(line.groups[0]!.series[0]).toMatchObject({
      color: '#aa0000',
      lineWidth: 3,
      markers: true,
    });
    expect(line.groups[0]!.series[1]).toMatchObject({ color: '#c0504d', markers: false });
    expect(line.valueAxis.gridlines).toBe('#d9d9d9');
    expect(line.categoryAxis.line).toBe('#bfbfbf');

    const untitled = parse(
      chartXml(`<c:areaChart>${series('Only', [1, 2, 3])}</c:areaChart>`, '<c:title/>'),
    );
    // With a single series, an automatic title shows the series name.
    expect(untitled.title).toBe('Only');
    expect(untitled.groups[0]!.kind).toBe('area');
    expect(
      parse(
        chartXml(`<c:areaChart>${series('A', [1])}${series('B', [1])}</c:areaChart>`, '<c:title/>'),
      ).title,
    ).toBeUndefined();

    const pie = parse(
      chartXml(`<c:pieChart><c:varyColors val="1"/>${series('P', [1, 2, 3])}</c:pieChart>`),
    );
    expect(pie.groups[0]).toMatchObject({ kind: 'pie', varyColors: true });
    const doughnut = parse(
      chartXml(`<c:doughnutChart>${series('D', [1, 2])}<c:holeSize val="75"/></c:doughnutChart>`),
    );
    expect(doughnut.groups[0]).toMatchObject({ kind: 'doughnut', holeSize: 75 });

    const scatter = parse(
      chartXml(
        `<c:scatterChart><c:scatterStyle val="marker"/>
          <c:ser><c:xVal>${numRef([1, 2, 4])}</c:xVal><c:yVal>${numRef([10, 20, 15])}</c:yVal></c:ser>
          <c:ser><c:spPr><a:ln><a:noFill/></a:ln></c:spPr><c:marker><c:symbol val="none"/></c:marker><c:xVal>${numRef([1])}</c:xVal><c:yVal>${numRef([1])}</c:yVal></c:ser>
          <c:ser><c:spPr><a:ln w="12700"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:ln></c:spPr><c:xVal>${numRef([1])}</c:xVal><c:yVal>${numRef([1])}</c:yVal></c:ser>
        </c:scatterChart>
        <c:valAx><c:axPos val="b"/><c:numFmt formatCode="0.0" sourceLinked="0"/></c:valAx><c:valAx><c:axPos val="l"/><c:spPr><a:ln><a:noFill/></a:ln></c:spPr><c:majorGridlines><c:spPr><a:ln><a:noFill/></a:ln></c:spPr></c:majorGridlines></c:valAx>`,
      ),
    );
    const [points, hidden, joined] = scatter.groups[0]!.series;
    expect(points).toMatchObject({
      xValues: [1, 2, 4],
      values: [10, 20, 15],
      connected: false,
      markers: true,
    });
    expect(hidden).toMatchObject({ connected: false, markers: false });
    expect(joined).toMatchObject({ connected: true, color: '#0000ff' });
    expect(scatter.categoryAxis.formatCode).toBe('0.0');
    expect(scatter.valueAxis.line).toBeUndefined();
    expect(scatter.valueAxis.gridlines).toBeUndefined();
  });

  it('reads literal and multi-level caches', () => {
    const model = parse(
      chartXml(
        `<c:barChart><c:ser><c:tx><c:v>Typed</c:v></c:tx>
          <c:cat><c:multiLvlStrRef><c:multiLvlStrCache><c:ptCount val="2"/><c:lvl><c:pt idx="0"><c:v>a</c:v></c:pt><c:pt idx="1"><c:v>b</c:v></c:pt></c:lvl></c:multiLvlStrCache></c:multiLvlStrRef></c:cat>
          <c:val><c:numLit><c:ptCount val="3"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="2"><c:v>oops</c:v></c:pt><c:pt idx="-4"><c:v>9</c:v></c:pt></c:numLit></c:val></c:ser></c:barChart>`,
      ),
    );
    expect(model.groups[0]!.series[0]).toMatchObject({
      name: 'Typed',
      categories: ['a', 'b'],
      values: [1, null, null],
    });
    expect(model.legend).toBeUndefined();
  });
});

describe('chart scales', () => {
  it('picks round steps', () => {
    expect(niceStep(280, 6)).toBe(50);
    expect(niceStep(1, 4)).toBe(0.25);
    expect(niceStep(9, 10)).toBe(1);
    expect(niceStep(90, 5)).toBe(20);
    expect(niceStep(950, 2)).toBe(500);
    expect(niceStep(7000, 1)).toBe(10000);
    expect(niceStep(0, 5)).toBe(1);
  });

  it('builds axis ranges', () => {
    const axis = {
      deleted: false,
      ticks: false,
      text: { size: 10, color: '#000', bold: false, italic: false },
    };
    expect(axisScale(30, 280, axis, 6)).toEqual({ min: 0, max: 300, step: 50 });
    expect(axisScale(-12, 30, axis, 4)).toEqual({ min: -20, max: 40, step: 20 });
    expect(axisScale(0, 0, axis, 4)).toMatchObject({ min: 0 });
    expect(axisScale(0, 0.17, { ...axis, max: 1, majorUnit: 0.25 }, 4)).toEqual({
      min: 0,
      max: 1,
      step: 0.25,
    });
    expect(axisScale(5, 9, { ...axis, min: 20 }, 4).max).toBeGreaterThan(20);
  });
});

describe('renderChart', () => {
  const count = (root: Element, selector: string): number => root.querySelectorAll(selector).length;
  const texts = (root: Element): string[] =>
    Array.from(root.querySelectorAll('text'), (node) => node.textContent ?? '');

  it('draws clustered columns with axes, gridlines and a legend', () => {
    const model = parse(
      chartXml(
        `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${series('North', [280, 30, 55], '<c:spPr><a:ln w="19050"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></c:spPr>')}${series('South', [-20, 40, 60])}<c:dLbls><c:showVal val="1"/></c:dLbls></c:barChart>${AXES}<c:spPr><a:solidFill><a:srgbClr val="FAFAFA"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></c:spPr>`,
        '<c:title><c:tx><c:rich><a:p><a:r><a:t>Costs</a:t></a:r></a:p></c:rich></c:tx></c:title>',
        '<c:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln></c:spPr>',
      ).replace('</c:plotArea>', '</c:plotArea><c:legend><c:legendPos val="b"/></c:legend>'),
    );
    const chart = renderChart(model, 400, 260);
    expect(chart.getAttribute('viewBox')).toBe('0 0 400 260');
    expect(chart.getAttribute('aria-label')).toBe('Costs');
    expect(chart.getAttribute('font-family')).toContain('Corbel');
    expect(count(chart, '.av-chart-series rect')).toBe(6);
    expect(chart.querySelector('.av-chart-series rect')!.getAttribute('stroke')).toBe('#ffffff');
    expect(count(chart, '.av-chart-legend rect')).toBe(2);
    expect(count(chart, '.av-chart-labels text')).toBe(6);
    const labels = texts(chart);
    expect(labels).toEqual(
      expect.arrayContaining(['Costs', 'North', 'South', 'Q1', 'Q3', '0.0', '300.0']),
    );
    // Gridlines, tick marks and the category axis line.
    expect(count(chart, '.av-chart-axes line')).toBeGreaterThan(8);
    expect(count(chart, '.av-chart-lines line')).toBe(1);
    expect(count(chart, '.av-chart-lines rect')).toBe(1);
    // Negative values hang below the zero line.
    const rects = Array.from(chart.querySelectorAll('.av-chart-series rect'));
    const zeroLine = Number(chart.querySelector('.av-chart-lines line')!.getAttribute('y1'));
    expect(Number(rects[3]!.getAttribute('y'))).toBeCloseTo(zeroLine, 1);
  });

  it('draws horizontal, stacked and percent-stacked bars', () => {
    const bar = (grouping: string, legend = ''): SVGElement =>
      renderChart(
        parse(
          chartXml(
            `<c:barChart><c:barDir val="bar"/><c:grouping val="${grouping}"/>${series('A', [1, 2, 3])}${series('B', [3, null, -1])}<c:dLbls><c:showVal val="1"/></c:dLbls></c:barChart>${AXES}`,
          ).replace('</c:plotArea>', `</c:plotArea>${legend}`),
        ),
        300,
        200,
      );
    const clustered = bar('clustered', '<c:legend><c:legendPos val="t"/></c:legend>');
    expect(count(clustered, '.av-chart-series rect')).toBe(5);
    const stacked = bar('stacked', '<c:legend><c:legendPos val="l"/></c:legend>');
    const first = stacked.querySelectorAll('.av-chart-series rect');
    // Stacked bars of one category share a row.
    expect(first[0]!.getAttribute('y')).toBe(first[3]!.getAttribute('y'));
    const percent = bar('percentStacked', '<c:legend><c:legendPos val="r"/></c:legend>');
    expect(texts(percent)).toEqual(expect.arrayContaining(['0%', '100%']));
    expect(count(percent, '.av-chart-legend text')).toBe(2);
  });

  it('draws the progress-bar style chart without a category axis', () => {
    const model = parse(
      chartXml(
        `<c:barChart><c:barDir val="bar"/>${series('Done', [0.17])}<c:gapWidth val="7"/></c:barChart>
         <c:valAx><c:scaling><c:max val="1"/></c:scaling><c:delete val="0"/><c:numFmt formatCode="0%" sourceLinked="0"/><c:majorTickMark val="out"/><c:majorUnit val="0.25"/></c:valAx>
         <c:catAx><c:delete val="1"/></c:catAx>
         <c:spPr><a:noFill/><a:ln w="25400"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:ln></c:spPr>`,
      ),
    );
    const chart = renderChart(model, 340, 76);
    expect(texts(chart)).toEqual(['0%', '25%', '50%', '75%', '100%']);
    expect(count(chart, '.av-chart-series rect')).toBe(1);
    expect(chart.querySelector('.av-chart-lines rect')!.getAttribute('stroke')).toBe('#000000');
    expect(chart.getAttribute('aria-label')).toBe('Chart');
    expect(count(chart, 'svg > rect')).toBe(0);
  });

  it('draws lines, areas and combinations', () => {
    const chart = renderChart(
      parse(
        chartXml(
          `<c:barChart>${series('Bars', [1, 2, 3])}</c:barChart>
           <c:lineChart>${series('Line', [2, null, 4], '<c:marker><c:symbol val="circle"/></c:marker>')}${series('Dot', [null, 5, null])}${series('None', [null, null, null])}<c:marker val="1"/></c:lineChart>
           <c:areaChart><c:grouping val="stacked"/>${series('Area', [1, 1, 1])}${series('Area 2', [1, 2, 1])}</c:areaChart>${AXES}`,
        ),
      ),
      500,
      300,
    );
    expect(count(chart, 'polyline')).toBe(1);
    expect(count(chart, 'polygon')).toBe(2);
    // Two markers on the line, plus the single-point series.
    expect(count(chart, '.av-chart-series circle')).toBe(3);
    expect(count(chart, '.av-chart-series rect')).toBe(3);
  });

  it('draws pies and doughnuts', () => {
    const pie = renderChart(
      parse(
        chartXml(
          `<c:pieChart><c:varyColors val="1"/>${series('Share', [1, 2, 0], '<c:dPt><c:idx val="1"/><c:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></c:spPr></c:dPt>')}</c:pieChart>`,
        ).replace('</c:plotArea>', '</c:plotArea><c:legend><c:legendPos val="r"/></c:legend>'),
      ),
      300,
      200,
    );
    const slices = pie.querySelectorAll('svg > path');
    expect(slices).toHaveLength(2);
    expect(slices[1]!.getAttribute('fill')).toBe('#ff0000');
    expect(slices[0]!.getAttribute('fill')).not.toBe(slices[1]!.getAttribute('fill'));
    expect(count(pie, '.av-chart-legend text')).toBe(3);
    expect(pie.querySelectorAll('.av-chart-legend rect')[1]!.getAttribute('fill')).toBe('#ff0000');

    const ring = renderChart(
      parse(chartXml(`<c:doughnutChart>${series('Ring', [5, 3, 2])}</c:doughnutChart>`)),
      200,
      200,
    );
    expect(ring.querySelector('path')!.getAttribute('d')).toMatch(/A.*L.*A.*Z$/);
    const full = renderChart(
      parse(chartXml(`<c:doughnutChart>${series('Full', [5])}</c:doughnutChart>`)),
      200,
      200,
    );
    expect(count(full, 'circle')).toBe(2);
    const solid = renderChart(
      parse(chartXml(`<c:pieChart>${series('Full', [5])}</c:pieChart>`)),
      200,
      200,
    );
    expect(count(solid, 'circle')).toBe(1);
    const empty = renderChart(
      parse(chartXml(`<c:pieChart>${series('Zero', [0, -1])}</c:pieChart>`)),
      200,
      200,
    );
    expect(count(empty, 'path')).toBe(0);
  });

  it('draws scatter charts on two numeric axes', () => {
    const chart = renderChart(
      parse(
        chartXml(
          `<c:scatterChart><c:scatterStyle val="lineMarker"/>
            <c:ser><c:xVal>${numRef([1, 2, null, 4])}</c:xVal><c:yVal>${numRef([10, 20, 5, 15])}</c:yVal></c:ser>
            <c:ser><c:spPr><a:ln><a:noFill/></a:ln></c:spPr><c:xVal>${numRef([3, 5])}</c:xVal><c:yVal>${numRef([1, 2])}</c:yVal></c:ser>
           </c:scatterChart><c:valAx><c:delete val="0"/><c:axPos val="b"/></c:valAx><c:valAx><c:delete val="0"/><c:axPos val="l"/></c:valAx>`,
        ),
      ),
      400,
      300,
    );
    expect(count(chart, 'polyline')).toBe(1);
    expect(count(chart, '.av-chart-series circle')).toBe(5);
    expect(texts(chart)).toContain('0');
  });

  it('copes with charts that have no values', () => {
    const chart = renderChart(
      parse(chartXml(`<c:barChart>${series('Empty', [null, null])}</c:barChart>${AXES}`)),
      20,
      20,
    );
    expect(count(chart, '.av-chart-series rect')).toBe(0);
    const crowded = renderChart(
      parse(
        chartXml(
          `<c:barChart><c:ser><c:cat>${strRef(Array.from({ length: 40 }, (_unused, index) => `Category ${index}`))}</c:cat><c:val>${numRef(Array.from({ length: 40 }, (_unused, index) => index))}</c:val></c:ser></c:barChart>${AXES}`,
        ),
      ),
      300,
      200,
    );
    // Overlapping category labels are thinned out.
    expect(texts(crowded).filter((label) => label.startsWith('Category')).length).toBeLessThan(20);
    const tall = renderChart(
      parse(
        chartXml(
          `<c:barChart><c:barDir val="bar"/><c:ser><c:val>${numRef(Array.from({ length: 40 }, (_unused, index) => index))}</c:val></c:ser></c:barChart>${AXES}`,
        ),
      ),
      300,
      200,
    );
    expect(texts(tall).length).toBeLessThan(40);
  });
});

function gridSource(cells: Record<string, GridCell>, extra: Partial<GridSource> = {}): GridSource {
  return {
    rowCount: 6,
    colCount: 6,
    sheet: true,
    cell: (row, col) => cells[`${row}:${col}`] ?? null,
    ...extra,
  };
}
const cellAt = (container: Element, text: string): HTMLElement =>
  Array.from(container.querySelectorAll<HTMLElement>('.av-grid-c')).find(
    (cell) => cell.textContent === text,
  )!;

describe('createGrid in sheet mode', () => {
  it('uses row heights, default sizes and hides empty rows and columns', () => {
    const container = mount();
    const grid = createGrid(
      container,
      gridSource(
        {
          '0:0': { text: 'a' },
          '1:0': { text: 'hidden row' },
          '2:1': { text: 'hidden col' },
          '2:2': { text: 'c' },
        },
        {
          rowHeights: [40, 0, 30],
          colWidths: [50, 0],
          defaultRowHeight: 20,
          defaultColWidth: 64,
          font: { family: 'Corbel', size: 14 },
        },
      ),
    );
    expect(grid.element.classList.contains('av-grid-sheet')).toBe(true);
    expect(cellAt(container, 'a').style.height).toBe('40px');
    expect(cellAt(container, 'hidden row')).toBeUndefined();
    expect(cellAt(container, 'hidden col')).toBeUndefined();
    const c = cellAt(container, 'c');
    expect(c.style.top).toBe('40px');
    expect(c.style.left).toBe('50px');
    expect(c.style.width).toBe('64px');
    expect(c.style.height).toBe('30px');
    const rowLabels = Array.from(
      container.querySelectorAll('.av-grid-rowhead .av-grid-h'),
      (h) => h.textContent,
    );
    expect(rowLabels).toEqual(['1', '3', '4', '5', '6']);
    const colLabels = Array.from(
      container.querySelectorAll('.av-grid-colhead .av-grid-h'),
      (h) => h.textContent,
    );
    expect(colLabels).toEqual(['A', 'C', 'D', 'E', 'F']);
    const cells = container.querySelector<HTMLElement>('.av-grid-cells')!;
    expect(cells.style.fontFamily).toContain('Corbel');
    expect(cells.style.fontSize).toBe('14px');
    grid.setZoom(2);
    expect(cells.style.fontSize).toBe('28px');
    expect(cellAt(container, 'c').style.top).toBe('80px');
    // Printing skips hidden rows and columns and trailing empty ones.
    const table = grid.toTable();
    expect(
      Array.from(table.rows, (row) => Array.from(row.cells, (cell) => cell.textContent)),
    ).toEqual([
      ['a', ''],
      ['', 'c'],
    ]);
  });

  it('draws borders once and drops default lines where needed', () => {
    const container = mount();
    const red = { width: 2, style: 'solid', color: '#ff0000' } as const;
    createGrid(
      container,
      gridSource({
        '0:0': { text: 'top', style: { borderBottom: red, borderRight: red } },
        '1:0': { text: 'below', style: { borderTop: red, borderLeft: red, background: '#eeeeee' } },
        '0:1': { text: 'right', style: { borderLeft: red, borderTop: red } },
        '3:3': { text: 'plain' },
      }),
    );
    expect(cellAt(container, 'top').style.borderBottom).toBe('2px solid #ff0000');
    expect(cellAt(container, 'top').style.borderRight).toBe('2px solid #ff0000');
    // The shared line belongs to the cell above / to the left.
    expect(cellAt(container, 'below').style.borderTop).toBe('');
    expect(cellAt(container, 'below').style.borderLeft).toBe('2px solid #ff0000');
    expect(cellAt(container, 'right').style.borderLeft).toBe('');
    expect(cellAt(container, 'right').style.borderTop).toBe('2px solid #ff0000');
    // Filled cells cover the default grid lines; plain ones keep them.
    expect(cellAt(container, 'below').style.borderBottomStyle).toBe('none');
    expect(cellAt(container, 'plain').style.borderBottomStyle).toBe('');
  });

  it('can hide the default grid lines', () => {
    const container = mount();
    createGrid(container, gridSource({ '0:0': { text: 'x' } }, { showGridLines: false }));
    expect(cellAt(container, 'x').style.borderRightStyle).toBe('none');
  });

  it('renders rich text, fills, icons, bars and filter buttons', () => {
    const container = mount();
    createGrid(
      container,
      gridSource({
        '0:0': {
          text: 'Rich text',
          runs: [
            { text: 'Rich ' },
            {
              text: 'text',
              style: { bold: true, color: '#ff0000', fontSize: 20, fontFamily: 'Impact' },
            },
          ],
        },
        '0:1': { text: ' $ 5.00 ', fill: { left: ' $', right: ' 5.00 ' } },
        '0:2': { text: '1', icon: { shape: 'check', color: '#00aa00' }, hideText: true },
        '0:3': { text: '7', bar: { color: '#638ec6', start: -1, end: 0.5 } },
        '0:4': { text: 'Header', filter: true, style: { valign: 'bottom', wrap: true, indent: 9 } },
        '1:0': { text: 'top', style: { valign: 'top', align: 'right', indent: 5 } },
      }),
    );
    const rich = cellAt(container, 'Rich text');
    const runs = rich.querySelectorAll<HTMLElement>('.av-grid-t span');
    expect(runs).toHaveLength(2);
    expect(runs[1]!.style.fontWeight).toBe('600');
    expect(runs[1]!.style.fontSize).toBe('20px');
    expect(runs[1]!.style.fontFamily).toContain('Impact');

    const money = cellAt(container, ' $ 5.00 ');
    const parts = money.querySelectorAll('.av-grid-fill span');
    expect(Array.from(parts, (part) => part.textContent)).toEqual([' $', ' 5.00 ']);

    const iconCell = container.querySelector('.av-grid-icon')!.parentElement!;
    expect(iconCell.textContent).toBe('');
    expect(iconCell.hasAttribute('title')).toBe(false);

    const bar = container.querySelector<HTMLElement>('.av-grid-bar')!;
    expect(bar.style.left).toBe('0%');
    expect(bar.style.width).toBe('50%');
    expect(bar.parentElement!.textContent).toBe('7');

    const header = cellAt(container, 'Header');
    expect(header.querySelector('.av-grid-filter svg')).not.toBeNull();
    expect(header.style.alignItems).toBe('flex-end');
    expect(header.querySelector('.av-grid-t')!.classList.contains('av-grid-wrap')).toBe(true);
    expect(header.querySelector<HTMLElement>('.av-grid-t')!.style.paddingLeft).toBe('9px');
    const top = cellAt(container, 'top');
    expect(top.style.alignItems).toBe('flex-start');
    expect(top.querySelector<HTMLElement>('.av-grid-t')!.style.paddingRight).toBe('5px');
  });

  it('lets text run over empty neighbours and stops at content and merges', () => {
    const container = mount();
    createGrid(
      container,
      gridSource(
        {
          '0:0': { text: 'a long title', overflow: true },
          '0:3': { text: 'stop' },
          '1:0': { text: 'blocked', overflow: true },
          '1:1': { text: 'x' },
          '2:5': { text: 'leftwards', overflow: true, style: { align: 'right' } },
          '2:2': { text: 'wall' },
          '3:0': { text: 'centred', overflow: true, style: { align: 'center' } },
          '4:0': { text: 'at merge', overflow: true },
          '5:0': { text: 'styled gap', overflow: true },
          '5:1': { text: '', style: { background: '#eee' } },
        },
        {
          colWidths: [50, 60, 70, 80, 90, 100],
          merges: [{ row: 4, col: 1, rowSpan: 1, colSpan: 2 }],
        },
      ),
    );
    const title = cellAt(container, 'a long title');
    expect(title.classList.contains('av-grid-spill')).toBe(true);
    // Columns A–C (50 + 60 + 70) minus the cell's own padding.
    expect(title.querySelector<HTMLElement>('.av-grid-t')!.style.width).toBe('173px');
    expect(cellAt(container, 'blocked').classList.contains('av-grid-spill')).toBe(false);
    const left = cellAt(container, 'leftwards');
    expect(left.style.justifyContent).toBe('flex-end');
    expect(left.querySelector<HTMLElement>('.av-grid-t')!.style.width).toBe('263px');
    expect(cellAt(container, 'centred').classList.contains('av-grid-spill')).toBe(false);
    expect(cellAt(container, 'at merge').classList.contains('av-grid-spill')).toBe(false);
    // Empty cells that only carry formatting do not stop the text.
    expect(cellAt(container, 'styled gap').classList.contains('av-grid-spill')).toBe(true);
  });

  it('keeps overflowing text that starts left of the visible columns', async () => {
    const container = mount();
    const cells: Record<string, GridCell> = {
      '0:1': { text: 'far left text', overflow: true },
      '1:2': { text: 'no overflow' },
      '2:3': { text: '', style: { background: '#eee' } },
    };
    const grid = createGrid(
      container,
      gridSource(cells, {
        colCount: 200,
        merges: [{ row: 3, col: 4, rowSpan: 1, colSpan: 2 }],
      }),
    );
    grid.element.scrollLeft = 100 * 20;
    grid.element.dispatchEvent(new Event('scroll'));
    await waitFor(
      () =>
        !Array.from(container.querySelectorAll('.av-grid-h')).some((h) => h.textContent === 'A'),
    );
    expect(cellAt(container, 'far left text')).toBeDefined();
    expect(cellAt(container, 'no overflow')).toBeUndefined();
  });

  it('places overlays from anchors and rescales them with zoom', () => {
    const container = mount();
    const sizes: [number, number][] = [];
    const make = (width: number, height: number): Element => {
      sizes.push([width, height]);
      return document.createElement('img');
    };
    const grid = createGrid(
      container,
      gridSource(
        {},
        {
          colWidths: [100, 100, 100],
          rowHeights: [20, 20, 20],
          overlays: [
            {
              from: { row: 1, col: 1, x: 5, y: 2 },
              to: { row: 2, col: 3, x: 0, y: 10 },
              create: make,
            },
            { from: { row: 0, col: 0 }, width: 30, height: 40, create: make },
            { from: { row: 0, col: 0 }, create: make },
            { from: { row: 99, col: 99 }, to: { row: 0, col: 0 }, create: make },
          ],
        },
      ),
    );
    const overlays = container.querySelectorAll<HTMLElement>('.av-grid-overlay');
    expect(overlays).toHaveLength(2);
    expect(sizes).toEqual([
      [195, 28],
      [30, 40],
    ]);
    expect(overlays[0]!.style.left).toBe('105px');
    expect(overlays[0]!.style.top).toBe('22px');
    expect(overlays[0]!.style.width).toBe('195px');
    grid.setZoom(2);
    expect(overlays[0]!.style.left).toBe('210px');
    expect(overlays[1]!.style.height).toBe('80px');
  });

  it('draws every icon shape', () => {
    const shapes = [
      'check',
      'cross',
      'exclamation',
      'up',
      'upRight',
      'right',
      'downRight',
      'down',
      'circle',
      'triangleUp',
      'triangleDown',
      'dash',
      'flag',
      'star',
    ] as const;
    for (const shape of shapes) {
      const icon = createGridIcon({ shape, color: '#123456' }, 16);
      expect(icon.getAttribute('viewBox'), shape).toBe('0 0 16 16');
      expect(icon.childNodes.length, shape).toBe(1);
    }
    expect(
      createGridIcon({ shape: 'down', color: '#000' }, 16).firstElementChild!.getAttribute(
        'transform',
      ),
    ).toBe('rotate(180 8 8)');
    for (const shape of ['check', 'cross', 'exclamation'] as const) {
      const circled = createGridIcon({ shape, color: '#123456', circled: true }, 12);
      expect(circled.querySelector('circle')!.getAttribute('fill')).toBe('#123456');
      expect(circled.lastElementChild!.getAttribute('stroke')).toBe('#ffffff');
    }
  });
});

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const rels = (entries: [string, string, string, boolean?][]): string =>
  `<Relationships>${entries
    .map(
      ([id, type, target, external]) =>
        `<Relationship Id="${id}" Type="${REL}/${type}" Target="${target}"${external ? ' TargetMode="External"' : ''}/>`,
    )
    .join('')}</Relationships>`;

const BOOK_STYLES = `<styleSheet>
  <numFmts><numFmt numFmtId="44" formatCode="_(&quot;$&quot;* #,##0.00_);_(&quot;$&quot;* \\(#,##0.00\\);_(&quot;$&quot;* &quot;-&quot;??_);_(@_)"/><numFmt numFmtId="165" formatCode="0.0;[Red]-0.0"/></numFmts>
  <fonts>
    <font><sz val="11"/><color theme="1"/><name val="Corbel"/><scheme val="minor"/></font>
    <font><b/><sz val="20"/><color theme="4"/><name val="Impact"/><scheme val="major"/></font>
    <font><sz val="11"/><color rgb="FF0000FF"/><name val="Corbel"/><scheme val="minor"/></font>
  </fonts>
  <fills>
    <fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill>
  </fills>
  <borders>
    <border/>
    <border><left style="thin"><color rgb="FF111111"/></left><right style="medium"><color rgb="FF222222"/></right><top style="thin"><color rgb="FF333333"/></top><bottom style="thick"><color rgb="FF444444"/></bottom></border>
  </borders>
  <cellXfs>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" applyFont="1"><alignment vertical="center"/></xf>
    <xf numFmtId="44" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="0" fontId="2" fillId="2" borderId="1" applyFont="1" applyFill="1" applyBorder="1"><alignment horizontal="centerContinuous" vertical="top" wrapText="1" indent="1"/></xf>
    <xf numFmtId="165" fontId="0" fillId="0" borderId="0"><alignment horizontal="left" vertical="distributed"/></xf>
    <xf numFmtId="0" fontId="0" fillId="2" borderId="0"><alignment horizontal="fill"/></xf>
  </cellXfs>
  <dxfs>
    <dxf><font><b/><color rgb="FFAA0000"/></font><fill><patternFill patternType="solid"><bgColor rgb="FFF2F2F2"/></patternFill></fill></dxf>
    <dxf><border><left/><right/><top style="dotted"><color rgb="FF00AA00"/></top><bottom style="dotted"><color rgb="FF00AA00"/></bottom><horizontal style="dotted"><color rgb="FF00AA00"/></horizontal></border><font><color rgb="FF123456"/></font></dxf>
    <dxf><font><b/></font><border><top/><bottom style="thick"><color rgb="FF0000AA"/></bottom></border></dxf>
    <dxf><fill><patternFill patternType="solid"><bgColor rgb="FFDDEEFF"/></patternFill></fill></dxf>
    <dxf><font><i/></font></dxf>
  </dxfs>
  <tableStyles>
    <tableStyle name="Mine">
      <tableStyleElement type="wholeTable" dxfId="1"/><tableStyleElement type="headerRow" dxfId="2"/>
      <tableStyleElement type="firstRowStripe" dxfId="0"/><tableStyleElement type="firstColumnStripe" dxfId="3"/>
      <tableStyleElement type="lastColumn" dxfId="4"/><tableStyleElement type="firstColumn" dxfId="4"/><tableStyleElement type="totalRow" dxfId="2"/>
    </tableStyle>
  </tableStyles>
</styleSheet>`;

const BOOK_SHEET = `<worksheet xmlns:r="${REL}" xmlns:x14="x" xmlns:xm="m">
  <sheetViews><sheetView showGridLines="0" workbookViewId="0"/></sheetViews>
  <sheetFormatPr defaultRowHeight="30" defaultColWidth="10"/>
  <cols>
    <col min="1" max="1" width="2.5" customWidth="1"/><col min="2" max="2" width="20" style="5" customWidth="1"/>
    <col min="4" max="4" width="9" hidden="1"/><col min="6" max="16384" width="12"/><col min="0" max="x"/>
  </cols>
  <sheetData>
    <row r="1" ht="60" customHeight="1"><c r="B1" s="1" t="s"><v>0</v></c><c r="C1" t="s"><v>1</v></c></row>
    <row r="2" hidden="1"><c r="B2"><v>1</v></c></row>
    <row r="3" s="3" customFormat="1"><c r="B3" s="3" t="s"><v>2</v></c><c r="C3" s="2"><v>55</v></c><c r="E3" s="4"><v>-1.25</v></c></row>
    <row r="4"><c r="B4" t="s"><v>3</v></c><c r="C4" t="s"><v>4</v></c><c r="D4" t="s"><v>5</v></c><c r="E4" t="s"><v>6</v></c></row>
    <row r="5"><c r="B5" t="s"><v>7</v></c><c r="C5"><v>1</v></c><c r="D5"><v>10</v></c><c r="E5" s="3"><v>3</v></c></row>
    <row r="6"><c r="B6" t="s"><v>8</v></c><c r="C6"><v>0</v></c><c r="D6"><v>20</v></c><c r="E6"><v>4</v></c></row>
    <row r="7"><c r="B7" t="s"><v>9</v></c><c r="C7"><v>1</v></c><c r="D7"><v>30</v></c><c r="E7"><v>5</v></c></row>
    <row r="8"><c r="B8" t="s"><v>10</v></c><c r="C8"/><c r="E8"><v>12</v></c></row>
    <row r="10"><c r="B10" s="3"/><c r="C10" s="3"/><c r="D10" s="3"/></row>
    <row r="${100_001}"><c r="A${100_001}"><v>1</v></c></row>
  </sheetData>
  <conditionalFormatting sqref="D5:D7"><cfRule type="dataBar" priority="2"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar></cfRule></conditionalFormatting>
  <conditionalFormatting sqref="E5:E7"><cfRule type="cellIs" operator="greaterThan" dxfId="0" priority="1"><formula>4</formula></cfRule><cfRule type="colorScale" priority="3"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFFFFFFF"/><color rgb="FF00FF00"/></colorScale></cfRule></conditionalFormatting>
  <mergeCells><mergeCell ref="B10:D10"/><mergeCell ref="bogus"/><mergeCell ref="ZZZ1:ZZZ2"/></mergeCells>
  <drawing r:id="rId1"/>
  <tableParts><tablePart r:id="rId2"/></tableParts>
  <extLst><ext><x14:conditionalFormattings><x14:conditionalFormatting><x14:cfRule type="iconSet" priority="9"><x14:iconSet iconSet="3Symbols" showValue="0" custom="1"><x14:cfvo type="percent"><xm:f>0</xm:f></x14:cfvo><x14:cfvo type="num" gte="0"><xm:f>0</xm:f></x14:cfvo><x14:cfvo type="num"><xm:f>1</xm:f></x14:cfvo><x14:cfIcon iconSet="NoIcons" iconId="0"/><x14:cfIcon iconSet="NoIcons" iconId="0"/><x14:cfIcon iconSet="3Symbols2" iconId="2"/></x14:iconSet></x14:cfRule><xm:sqref>C5:C8</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>
</worksheet>`;

const BOOK_STRINGS = `<sst>
  <si><r><rPr><b/><sz val="20"/><color rgb="FF000000"/><rFont val="Impact"/></rPr><t>Back to </t></r><r><rPr><color rgb="FFC67432"/></rPr><t>school</t></r><r><t> plan</t></r></si>
  <si><t>next</t></si><si><t>Wrapped and styled</t></si>
  <si><t>Item</t></si><si><t>To buy</t></si><si><t>Cost</t></si><si><t>Qty</t></si>
  <si><t>Math</t></si><si><t>English</t></si><si><t>Backpack</t></si><si><t>Total</t></si>
</sst>`;

const BOOK_DRAWING = `<xdr:wsDr xmlns:xdr="x" xmlns:a="a" xmlns:r="${REL}" xmlns:c="c">
  <xdr:twoCellAnchor><xdr:from><xdr:col>6</xdr:col><xdr:colOff>95250</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>19050</xdr:rowOff></xdr:from><xdr:to><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:pic><xdr:nvPicPr><xdr:cNvPr id="1" name="Picture" descr="A room"/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rId1"/></xdr:blipFill></xdr:pic></xdr:twoCellAnchor>
  <xdr:oneCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>11</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="3810000" cy="1905000"/>
    <xdr:graphicFrame><a:graphic><a:graphicData><c:chart r:id="rId2"/></a:graphicData></a:graphic></xdr:graphicFrame></xdr:oneCellAnchor>
  <xdr:absoluteAnchor><xdr:pos x="9525" y="19050"/><xdr:ext cx="952500" cy="476250"/>
    <xdr:graphicFrame><a:graphic><a:graphicData><c:chart r:id="rId2"/></a:graphicData></a:graphic></xdr:graphicFrame></xdr:absoluteAnchor>
  <xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:row>0</xdr:row></xdr:from><xdr:pic><xdr:blipFill><a:blip r:embed="rId1"/></xdr:blipFill></xdr:pic></xdr:twoCellAnchor>
  <xdr:twoCellAnchor><xdr:from><xdr:col>-1</xdr:col><xdr:row>0</xdr:row></xdr:from><xdr:to><xdr:col>1</xdr:col><xdr:row>1</xdr:row></xdr:to><xdr:pic><xdr:blipFill><a:blip r:embed="rId1"/></xdr:blipFill></xdr:pic></xdr:twoCellAnchor>
  <xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:row>0</xdr:row></xdr:from><xdr:ext cx="1" cy="1"/><xdr:pic><xdr:blipFill><a:blip r:embed="rId3"/></xdr:blipFill></xdr:pic></xdr:oneCellAnchor>
  <xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:row>0</xdr:row></xdr:from><xdr:ext cx="1" cy="1"/><xdr:pic><xdr:blipFill><a:blip r:embed="rId4"/></xdr:blipFill></xdr:pic></xdr:oneCellAnchor>
  <xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:row>0</xdr:row></xdr:from><xdr:ext cx="1" cy="1"/><xdr:pic><xdr:blipFill><a:blip r:embed="rId5"/></xdr:blipFill></xdr:pic></xdr:oneCellAnchor>
  <xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:row>0</xdr:row></xdr:from><xdr:ext cx="1" cy="1"/><xdr:graphicFrame><a:graphic><a:graphicData><c:chart r:id="rId6"/></a:graphicData></a:graphic></xdr:graphicFrame></xdr:oneCellAnchor>
  <xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:row>0</xdr:row></xdr:from><xdr:ext cx="1" cy="1"/><xdr:graphicFrame><a:graphic><a:graphicData><c:chart r:id="rId7"/></a:graphicData></a:graphic></xdr:graphicFrame></xdr:oneCellAnchor>
  <xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:row>0</xdr:row></xdr:from><xdr:ext cx="1" cy="1"/><xdr:sp/></xdr:oneCellAnchor>
  <xdr:somethingElse/>
</xdr:wsDr>`;

function bookFixture(): Uint8Array {
  return makeZip({
    'xl/workbook.xml': `<workbook xmlns:r="${REL}"><sheets><sheet name="Plan" r:id="rId1"/><sheet name="Bare" r:id="rId5"/></sheets></workbook>`,
    'xl/_rels/workbook.xml.rels': rels([
      ['rId1', 'worksheet', 'worksheets/sheet1.xml'],
      ['rId2', 'styles', 'styles.xml'],
      ['rId3', 'sharedStrings', 'sharedStrings.xml'],
      ['rId4', 'theme', 'theme/theme1.xml'],
      ['rId5', 'worksheet', 'worksheets/sheet2.xml'],
    ]),
    'xl/theme/theme1.xml': THEME_XML,
    'xl/styles.xml': BOOK_STYLES,
    'xl/sharedStrings.xml': BOOK_STRINGS,
    'xl/worksheets/sheet1.xml': BOOK_SHEET,
    'xl/worksheets/sheet2.xml':
      '<worksheet><sheetFormatPr zeroHeight="1"/><sheetData><row r="1"><c r="A1"><v>5</v></c></row></sheetData></worksheet>',
    'xl/worksheets/_rels/sheet1.xml.rels': rels([
      ['rId1', 'drawing', '../drawings/drawing1.xml'],
      ['rId2', 'table', '../tables/table1.xml'],
      ['rId3', 'table', '../tables/missing.xml'],
      ['rId4', 'table', '../tables/broken.xml'],
      ['rId5', 'hyperlink', 'https://example.com', true],
      ['rId6', 'table', '../tables/table2.xml'],
      ['rId7', 'drawing', '../drawings/missing.xml'],
    ]),
    'xl/worksheets/_rels/sheet2.xml.rels': rels([['rId1', 'table', '../tables/table3.xml']]),
    'xl/tables/table1.xml':
      '<table ref="B4:E8" totalsRowCount="1"><autoFilter ref="B4:E7"/><tableStyleInfo name="Mine" showFirstColumn="1" showLastColumn="1" showRowStripes="1" showColumnStripes="1"/></table>',
    'xl/tables/table2.xml': '<table ref="nowhere"/>',
    'xl/tables/table3.xml':
      '<table ref="A1:B3" headerRowCount="0"><tableStyleInfo name="TableStyleMedium2" showRowStripes="1"/></table>',
    'xl/tables/broken.xml': '<table',
    'xl/drawings/drawing1.xml': BOOK_DRAWING,
    'xl/drawings/_rels/drawing1.xml.rels': rels([
      ['rId1', 'image', '../media/image1.png'],
      ['rId2', 'chart', '../charts/chart1.xml'],
      ['rId3', 'image', '../media/image2.emf'],
      ['rId4', 'image', 'https://example.com/a.png', true],
      ['rId5', 'image', '../media/missing.png'],
      ['rId6', 'chart', 'https://example.com/chart.xml', true],
      ['rId7', 'chart', '../charts/empty.xml'],
    ]),
    'xl/media/image1.png': PNG_1X1,
    'xl/media/image2.emf': PNG_1X1,
    'xl/charts/chart1.xml': chartXml(
      `<c:barChart>${series('Costs', [280, 30, 55])}</c:barChart>${AXES}`,
    ),
    'xl/charts/empty.xml': chartXml(''),
  });
}

describe('XLSX fidelity', () => {
  const load = async (host?: { objectUrl(blob: Blob): string }): Promise<GridSource> => {
    const [plan] = await openWorkbook(openZip(bookFixture()), host);
    return plan!.load();
  };

  it('lays out rows and columns like the workbook', async () => {
    const grid = await load();
    expect(grid.sheet).toBe(true);
    expect(grid.showGridLines).toBe(false);
    expect(grid.font).toEqual({ size: (11 * 96) / 72, family: 'Corbel' });
    // 60pt and 30pt rows; hidden rows collapse.
    expect(grid.rowHeights?.[0]).toBe(80);
    expect(grid.rowHeights?.[1]).toBe(0);
    expect(grid.defaultRowHeight).toBe(40);
    // Widths count digits of the default font (7px each without a canvas).
    expect(grid.colWidths?.[0]).toBe(17);
    expect(grid.colWidths?.[1]).toBe(140);
    expect(grid.colWidths?.[3]).toBe(0);
    expect(grid.colWidths?.[5]).toBe(84);
    expect(grid.defaultColWidth).toBe(70);
    expect(grid.merges).toEqual([{ row: 9, col: 1, rowSpan: 1, colSpan: 3 }]);
    // Rows beyond the limit are reported as truncated.
    expect(grid.totalRows).toBe(100_001);
  });

  it('resolves fonts, fills, borders and alignment', async () => {
    const grid = await load();
    const title = grid.cell(0, 1)!;
    expect(title.style).toMatchObject({
      bold: true,
      color: '#4f81bd',
      fontFamily: 'Impact',
      valign: 'middle',
    });
    expect(title.style!.fontSize).toBeCloseTo(26.67, 1);
    expect(title.runs).toHaveLength(3);
    expect(title.runs![0]!.style).toMatchObject({
      bold: true,
      fontFamily: 'Impact',
      color: '#000000',
    });
    expect(title.runs![1]!.style).toEqual({ color: '#C67432' });
    expect(title.runs![2]!.style).toBeUndefined();
    expect(title.overflow).toBe(true);
    // Plain cells use the default font, so it is not repeated.
    expect(grid.cell(0, 2)!.style).toEqual({ color: '#000000', valign: 'bottom' });

    const styled = grid.cell(2, 1)!;
    expect(styled.style).toMatchObject({
      color: '#0000FF',
      background: '#FFFF00',
      align: 'center',
      valign: 'top',
      wrap: true,
      indent: 9,
      borderLeft: { width: 1, style: 'solid', color: '#111111' },
      borderRight: { width: 2, style: 'solid', color: '#222222' },
      borderTop: { width: 1, style: 'solid', color: '#333333' },
      borderBottom: { width: 3, style: 'solid', color: '#444444' },
    });
    expect(styled.overflow).toBeUndefined();
    // Empty cells take the row format, then the column format.
    expect(grid.cell(2, 6)!.style!.background).toBe('#FFFF00');
    expect(grid.cell(11, 1)!.style!.background).toBe('#FFFF00');
    expect(grid.cell(11, 2)).toBeNull();
  });

  it('formats accounting values and coloured number formats', async () => {
    const grid = await load();
    expect(grid.cell(2, 2)).toMatchObject({
      text: ' $55.00 ',
      fill: { left: ' $', right: '55.00 ' },
    });
    const negative = grid.cell(2, 4)!;
    expect(negative.text).toBe('-1.3');
    expect(negative.style).toMatchObject({ color: '#ff0000', align: 'left', valign: 'middle' });
  });

  it('applies table styles with banding, headers, totals and filter buttons', async () => {
    const grid = await load();
    const dotted = { width: 1, style: 'dotted', color: '#00AA00' };
    const thick = { width: 3, style: 'solid', color: '#0000AA' };
    const header = grid.cell(3, 1)!;
    expect(header.filter).toBe(true);
    expect(header.style).toMatchObject({ bold: true, color: '#123456', borderBottom: thick });
    // The header clears the table's top line.
    expect(header.style!.borderTop).toBeUndefined();
    expect(grid.cell(4, 1)!.filter).toBeUndefined();

    // First stripe row: banded fill, bold red text, dotted lines.
    const math = grid.cell(4, 1)!;
    expect(math.style).toMatchObject({
      background: '#F2F2F2',
      bold: true,
      italic: true,
      color: '#AA0000',
      borderTop: dotted,
      borderBottom: dotted,
    });
    // Second stripe row keeps the column stripe of the first column.
    const english = grid.cell(5, 1)!;
    expect(english.style).toMatchObject({ background: '#DDEEFF', color: '#123456', italic: true });
    expect(english.style!.bold).toBeUndefined();
    expect(grid.cell(5, 2)!.style!.background).toBeUndefined();
    // Formatting set on the cell itself beats the table style.
    const own = grid.cell(4, 4)!;
    expect(own.style).toMatchObject({ color: '#0000FF' });
    expect(own.style!.bold).toBeUndefined();
    expect(own.style!.borderBottom).toEqual({ width: 3, style: 'solid', color: '#444444' });
    // Last column and total row.
    expect(grid.cell(5, 4)!.style!.italic).toBe(true);
    const total = grid.cell(7, 1)!;
    expect(total.style).toMatchObject({ bold: true, borderBottom: thick });
    expect(total.style!.background).toBeUndefined();
    // Empty cells inside a table are styled too.
    expect(grid.cell(7, 3)!.style).toMatchObject({ borderBottom: thick });
  });

  it('applies conditional formats, icons and bars', async () => {
    const grid = await load();
    expect(grid.cell(4, 2)).toMatchObject({ hideText: true, icon: { shape: 'check' } });
    const none = grid.cell(5, 2)!;
    expect(none.icon).toBeUndefined();
    expect(none.hideText).toBe(true);
    expect(grid.cell(7, 2)!.icon).toBeUndefined();
    expect(grid.cell(4, 3)!.bar).toEqual({ color: '#638EC6', start: 0, end: 10 / 30 });
    expect(grid.cell(6, 3)!.bar!.end).toBe(1);
    // Colour scale plus a higher-priority highlight rule.
    expect(grid.cell(4, 4)!.style!.background).toBe('#ffffff');
    const high = grid.cell(6, 4)!;
    expect(high.style).toMatchObject({ background: '#00ff00', bold: true, color: '#AA0000' });
  });

  it('takes the far borders of a merged area from its last cell', async () => {
    const grid = await load();
    const merged = grid.cell(9, 1)!;
    expect(merged.style!.borderRight).toEqual({ width: 2, style: 'solid', color: '#222222' });
    expect(merged.style!.borderBottom).toEqual({ width: 3, style: 'solid', color: '#444444' });
  });

  it('loads pictures and charts as overlays', async () => {
    const blobs: Blob[] = [];
    const grid = await load({
      objectUrl(blob) {
        blobs.push(blob);
        return `blob:test/${blobs.length}`;
      },
    });
    expect(grid.overlays).toHaveLength(3);
    expect(blobs).toHaveLength(1);
    expect(blobs[0]!.type).toBe('image/png');
    const [picture, chart, absolute] = grid.overlays!;
    expect(picture!.from).toEqual({ row: 0, col: 6, x: 10, y: 2 });
    expect(picture!.to).toEqual({ row: 4, col: 8, x: 0, y: 0 });
    const image = picture!.create(100, 50) as HTMLImageElement;
    expect(image.tagName).toBe('IMG');
    expect(image.getAttribute('src')).toBe('blob:test/1');
    expect(image.alt).toBe('A room');
    expect(chart).toMatchObject({ from: { row: 11, col: 1 }, width: 400, height: 200 });
    const drawn = chart!.create(400, 200);
    expect(drawn.tagName.toLowerCase()).toBe('svg');
    expect(drawn.querySelectorAll('.av-chart-series rect')).toHaveLength(3);
    expect(absolute).toMatchObject({
      from: { row: 0, col: 0, x: 1, y: 2 },
      width: 100,
      height: 50,
    });
    // The grid is large enough to hold the overlays.
    expect(grid.colCount).toBeGreaterThanOrEqual(9);
    expect(grid.rowCount).toBeGreaterThanOrEqual(12);
  });

  it('skips pictures when no host can provide URLs, but keeps charts', async () => {
    const grid = await load();
    expect(grid.overlays).toHaveLength(2);
  });

  it('applies built-in table styles and collapsed default rows', async () => {
    const sheets = await openWorkbook(openZip(bookFixture()));
    const bare = await sheets[1]!.load();
    expect(bare.defaultRowHeight).toBe(0);
    expect(bare.showGridLines).toBe(true);
    expect(bare.totalRows).toBeUndefined();
    // No header row: banding starts on the first row.
    expect(bare.cell(0, 0)!.style!.background).toMatch(/^#[0-9a-f]{6}$/);
    expect(bare.cell(1, 0)!.style!.background).toBeUndefined();
    expect(bare.cell(0, 0)!.filter).toBeUndefined();
    expect(bare.cell(0, 0)!.style!.align).toBe('right');
  });

  it('renders the whole workbook in the viewer', async () => {
    const { context, container, objectUrls } = renderContext(bookFixture(), 'plan.xlsx', 'xlsx');
    const instance = instanceOf(await xlsxRenderer.render(context));
    await waitFor(() => container.querySelector('.av-grid-overlay') !== null);
    expect(objectUrls).toHaveLength(1);
    expect(container.querySelector('.av-grid-sheet')).not.toBeNull();
    expect(container.querySelector('.av-grid-overlay img')).not.toBeNull();
    expect(container.querySelector('.av-grid-overlay svg.av-chart')).not.toBeNull();
    expect(container.querySelectorAll('.av-grid-icon').length).toBeGreaterThan(0);
    // Four header cells, one of them in a hidden column.
    expect(container.querySelectorAll('.av-grid-filter')).toHaveLength(3);
    expect(container.querySelector('.av-grid-fill')).not.toBeNull();
    expect(container.querySelector('.av-notice')!.textContent).toContain('100,001');
    const printable = instance.printable!() as HTMLTableElement;
    expect(printable.rows.length).toBeGreaterThan(5);
    // Values hidden behind icons are not printed.
    expect(printable.textContent).toContain('Math');
    instance.destroy!();
  });
});
