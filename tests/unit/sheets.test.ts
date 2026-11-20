import { afterEach, describe, expect, it } from 'vitest';
import { openOdsWorkbook } from '../../src/formats/ods';
import { openWorkbook, parseCellRef } from '../../src/formats/xlsx';
import { estimateColumnWidths, renderer as csvRenderer } from '../../src/renderers/csv';
import { renderer as odsRenderer } from '../../src/renderers/ods';
import { renderer as xlsxRenderer } from '../../src/renderers/xlsx';
import { columnName, createGrid, type GridSource } from '../../src/ui/grid';
import { openZip } from '../../src/util/zip';
import { instanceOf, mount, renderContext, waitFor } from '../helpers/dom';
import { odsFixture, xlsxFixture } from '../helpers/fixtures';
import { makeZip } from '../helpers/zip';

afterEach(() => {
  document.body.textContent = '';
});

const cellTexts = (root: Element): string[] =>
  Array.from(root.querySelectorAll('.av-grid-c'), (cell) => cell.textContent ?? '');

function simpleSource(rows: string[][], extra: Partial<GridSource> = {}): GridSource {
  return {
    rowCount: rows.length,
    colCount: Math.max(...rows.map((row) => row.length)),
    cell: (row, col) => (rows[row]?.[col] ? { text: rows[row][col] } : null),
    ...extra,
  };
}

describe('columnName', () => {
  it.each([
    [0, 'A'],
    [25, 'Z'],
    [26, 'AA'],
    [27, 'AB'],
    [701, 'ZZ'],
    [702, 'AAA'],
  ])('names column %d as %s', (index, expected) => {
    expect(columnName(index)).toBe(expected);
  });
});

describe('createGrid', () => {
  it('renders headers and cells', () => {
    const container = mount();
    const grid = createGrid(
      container,
      simpleSource([
        ['a', 'b'],
        ['c', 'd'],
      ]),
    );
    expect(cellTexts(container)).toEqual(['a', 'b', 'c', 'd']);
    expect(
      Array.from(container.querySelectorAll('.av-grid-colhead .av-grid-h'), (h) => h.textContent),
    ).toEqual(['A', 'B']);
    expect(
      Array.from(container.querySelectorAll('.av-grid-rowhead .av-grid-h'), (h) => h.textContent),
    ).toEqual(['1', '2']);
    grid.destroy();
    expect(container.querySelector('.av-grid')).toBeNull();
  });

  it('only keeps visible rows in the DOM and follows scrolling', async () => {
    const container = mount();
    const rows = Array.from({ length: 10_000 }, (_unused, index) => [`row ${index + 1}`]);
    const grid = createGrid(container, simpleSource(rows));
    const rendered = container.querySelectorAll('.av-grid-c').length;
    expect(rendered).toBeGreaterThan(10);
    expect(rendered).toBeLessThan(100);
    expect(cellTexts(container)[0]).toBe('row 1');

    grid.element.scrollTop = 24 * 5000;
    grid.element.dispatchEvent(new Event('scroll'));
    await waitFor(() => cellTexts(container).includes('row 5001'));
    expect(cellTexts(container)).not.toContain('row 1');
    grid.destroy();
  });

  it('virtualizes columns too', () => {
    const container = mount();
    const row = Array.from({ length: 500 }, (_unused, index) => `c${index}`);
    createGrid(container, simpleSource([row]));
    const texts = cellTexts(container);
    expect(texts).toContain('c0');
    expect(texts).not.toContain('c499');
  });

  it('draws merged areas once with their full size', () => {
    const container = mount();
    createGrid(
      container,
      simpleSource(
        [
          ['merged', '', 'x'],
          ['', '', 'y'],
        ],
        {
          colWidths: [50, 70, 100],
          merges: [
            { row: 0, col: 0, rowSpan: 2, colSpan: 2 },
            { row: 0, col: 2, rowSpan: 1, colSpan: 1 },
            { row: 9, col: 9, rowSpan: 2, colSpan: 2 },
          ],
        },
      ),
    );
    const cells = Array.from(container.querySelectorAll<HTMLElement>('.av-grid-c'));
    expect(cells.map((cell) => cell.textContent)).toEqual(['merged', 'x', 'y']);
    expect(cells[0]!.style.width).toBe('120px');
    expect(cells[0]!.style.height).toBe('48px');
  });

  it('applies cell styles and zoom', () => {
    const container = mount();
    const grid = createGrid(container, {
      rowCount: 1,
      colCount: 1,
      colWidths: [80],
      cell: () => ({
        text: 'a long styled cell value',
        style: {
          bold: true,
          italic: true,
          underline: true,
          strike: true,
          color: '#ff0000',
          background: '#ffff00',
          align: 'right',
        },
      }),
    });
    const cell = container.querySelector<HTMLElement>('.av-grid-c')!;
    expect(cell.style.fontWeight).toBe('600');
    expect(cell.style.fontStyle).toBe('italic');
    expect(cell.style.textDecoration).toBe('underline line-through');
    expect(cell.style.textAlign).toBe('right');
    expect(cell.title).toBe('a long styled cell value');
    expect(cell.style.width).toBe('80px');

    grid.setZoom(2);
    expect(grid.getZoom()).toBe(2);
    expect(container.querySelector<HTMLElement>('.av-grid-c')!.style.width).toBe('160px');
    grid.setZoom(2);
    grid.setZoom(0);
    expect(grid.getZoom()).toBe(2);
  });

  it('builds a printable table with spans', () => {
    const container = mount();
    const grid = createGrid(
      container,
      simpleSource(
        [
          ['a', '', 'c'],
          ['d', 'e', 'f'],
          ['g', 'h', 'i'],
        ],
        { merges: [{ row: 0, col: 0, rowSpan: 1, colSpan: 2 }] },
      ),
    );
    const table = grid.toTable(2);
    expect(table.rows).toHaveLength(2);
    expect(table.rows[0]!.cells).toHaveLength(2);
    expect(table.rows[0]!.cells[0]!.getAttribute('colspan')).toBe('2');
  });

  it('handles empty data', () => {
    const container = mount();
    createGrid(container, { rowCount: 0, colCount: 0, cell: () => null });
    expect(container.querySelectorAll('.av-grid-c')).toHaveLength(0);
  });
});

describe('XLSX', () => {
  it('parses cell references', () => {
    expect(parseCellRef('A1')).toEqual({ row: 0, col: 0 });
    expect(parseCellRef('bc12')).toEqual({ row: 11, col: 54 });
    expect(parseCellRef('XFD1048576')).toEqual({ row: 1048575, col: 16383 });
    expect(parseCellRef('12')).toBeNull();
    expect(parseCellRef('AB')).toBeNull();
    expect(parseCellRef('A0')).toBeNull();
    expect(parseCellRef('A1.5')).toBeNull();
  });

  it('lists visible sheets and skips hidden ones', async () => {
    const sheets = await openWorkbook(openZip(xlsxFixture()));
    expect(sheets.map((sheet) => sheet.name)).toEqual(['Sales', 'Notes', 'Broken']);
  });

  it('reads strings, numbers, dates, booleans, errors and formulas', async () => {
    const [sales] = await openWorkbook(openZip(xlsxFixture()));
    const grid = await sales!.load();
    const text = (row: number, col: number): string | undefined => grid.cell(row, col)?.text;
    expect([text(0, 0), text(0, 1), text(0, 2), text(0, 3)]).toEqual([
      'Product',
      'Price',
      'Rich text',
      ' padded ',
    ]);
    expect([text(1, 0), text(1, 1), text(1, 2), text(1, 3)]).toEqual([
      'Inline',
      '1,234.50',
      '2023-03-15',
      '26%',
    ]);
    expect([text(2, 0), text(2, 1), text(2, 2), text(2, 3)]).toEqual([
      'TRUE',
      'FALSE',
      '#DIV/0!',
      'formula text',
    ]);
    expect(text(3, 0)).toBe('2024-01-31 10:00:00');
    expect(text(3, 1)).toBe('42');
    expect(grid.cell(3, 2)).toBeNull();
    expect(text(3, 4)).toBe('not-a-number');
    // Rows and cells without references continue after the previous ones.
    expect([text(4, 0), text(4, 1)]).toEqual(['7', '8']);
  });

  it('applies fonts, fills, alignment, widths and merges', async () => {
    const [sales] = await openWorkbook(openZip(xlsxFixture()));
    const grid = await sales!.load();
    expect(grid.cell(0, 0)?.style).toEqual({
      bold: true,
      italic: true,
      underline: true,
      strike: true,
      color: '#FF0000',
      background: '#FFFF00',
      align: 'center',
      valign: 'bottom',
    });
    expect(grid.cell(1, 1)?.style).toEqual({
      color: '#000000',
      align: 'right',
      valign: 'bottom',
    });
    expect(grid.cell(2, 0)?.style).toEqual({
      color: '#000000',
      align: 'center',
      valign: 'bottom',
    });
    // An empty cell with a fill is kept so its background shows.
    expect(grid.cell(3, 3)).toEqual({
      text: '',
      style: { color: '#000000', background: '#FFFF00', valign: 'bottom' },
    });
    expect(grid.colWidths?.[0]).toBe(140);
    // Hidden columns take no space.
    expect(grid.colWidths?.[1]).toBe(0);
    expect(grid.merges).toEqual([{ row: 5, col: 0, rowSpan: 2, colSpan: 3 }]);
    // Two spare columns follow the content, with a minimum of eight.
    expect(grid.colCount).toBe(8);
  });

  it('handles empty sheets and reports missing ones', async () => {
    const sheets = await openWorkbook(openZip(xlsxFixture()));
    const empty = await sheets[1]!.load();
    expect(empty).toMatchObject({ rowCount: 12, colCount: 8 });
    expect(empty.cell(0, 0)).toBeNull();
    await expect(sheets[2]!.load()).rejects.toMatchObject({ code: 'PARSE_FAILED' });
  });

  it('rejects packages that are not workbooks', async () => {
    await expect(openWorkbook(openZip(makeZip({ 'a.txt': 'x' })))).rejects.toMatchObject({
      code: 'PARSE_FAILED',
    });
    const noSheets = makeZip({
      'xl/workbook.xml': '<workbook><sheets/></workbook>',
    });
    await expect(openWorkbook(openZip(noSheets))).rejects.toMatchObject({ code: 'PARSE_FAILED' });
  });

  it('works without styles or shared strings and with the 1904 date system', async () => {
    const zip = makeZip({
      'xl/workbook.xml':
        '<workbook xmlns:r="r"><workbookPr date1904="1"/><sheets><sheet name="S" r:id="rId1"/></sheets></workbook>',
      'xl/_rels/workbook.xml.rels':
        '<Relationships><Relationship Id="rId1" Type="x/worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
      'xl/worksheets/sheet1.xml':
        '<worksheet><sheetData><row r="1"><c r="A1" t="s"><v>3</v></c><c r="B1"><v>1.5</v></c></row></sheetData></worksheet>',
    });
    const [sheet] = await openWorkbook(openZip(zip));
    const grid = await sheet!.load();
    expect(grid.cell(0, 0)).toBeNull();
    expect(grid.cell(0, 1)?.text).toBe('1.5');
  });

  it('renders with sheet tabs, switches sheets and zooms', async () => {
    const { context, container } = renderContext(xlsxFixture(), 'book.xlsx', 'xlsx');
    const instance = instanceOf(await xlsxRenderer.render(context));
    await waitFor(() => cellTexts(container).includes('Product'));
    const tabs = container.querySelectorAll<HTMLButtonElement>('.av-tab');
    expect(Array.from(tabs, (tab) => tab.textContent)).toEqual(['Sales', 'Notes', 'Broken']);
    expect(tabs[0]!.getAttribute('aria-selected')).toBe('true');

    instance.zoom!.set(1.5);
    expect(instance.zoom!.get()).toBe(1.5);
    expect(instance.printable!()).toBeInstanceOf(HTMLTableElement);

    tabs[1]!.click();
    await waitFor(() => !cellTexts(container).includes('Product'));
    expect(tabs[1]!.getAttribute('aria-selected')).toBe('true');
    tabs[1]!.click();

    tabs[2]!.click();
    await waitFor(() => container.textContent?.includes('missing from the workbook'));

    tabs[0]!.click();
    await waitFor(() => cellTexts(container).includes('Product'));
    instance.destroy!();
    expect(instance.printable!()).toBeNull();
  });

  it('ignores a slow sheet that finishes after the view is destroyed', async () => {
    const { context, container } = renderContext(xlsxFixture(), 'book.xlsx', 'xlsx');
    const instance = instanceOf(await xlsxRenderer.render(context));
    instance.destroy!();
    await new Promise((resolve) => setTimeout(resolve, 30));
    expect(container.querySelector('.av-grid')).toBeNull();
  });
});

describe('ODS', () => {
  it('reads tables, repeats, merges and trailing padding', async () => {
    const sheets = await openOdsWorkbook(openZip(odsFixture()));
    expect(sheets.map((sheet) => sheet.name)).toEqual(['Data', 'Sheet2']);
    const grid = await sheets[0]!.load();
    const text = (row: number, col: number): string | undefined => grid.cell(row, col)?.text;
    expect([text(0, 0), text(0, 1)]).toEqual(['Name', 'Amount']);
    expect(text(1, 1)).toBe('12.50');
    expect(grid.cell(1, 1)?.style).toEqual({ align: 'right' });
    expect([text(2, 0), text(2, 1), text(3, 0), text(3, 1)]).toEqual([
      'same',
      'same',
      'same',
      'same',
    ]);
    expect(text(4, 0)).toBe('merged cell\nline two');
    expect(grid.merges).toEqual([{ row: 4, col: 0, rowSpan: 2, colSpan: 2 }]);
    // A million padding rows and a thousand padding columns are not materialised.
    expect(grid.rowCount).toBe(5);
    expect(grid.colCount).toBe(2);
    expect(await sheets[0]!.load()).toBe(grid);

    const empty = await sheets[1]!.load();
    expect(empty).toMatchObject({ rowCount: 1, colCount: 1 });
  });

  it('rejects packages without content or sheets', async () => {
    await expect(openOdsWorkbook(openZip(makeZip({ 'a.txt': 'x' })))).rejects.toMatchObject({
      code: 'PARSE_FAILED',
    });
    const noTables = makeZip({
      'content.xml':
        '<office:document-content><office:body><office:spreadsheet/></office:body></office:document-content>',
    });
    await expect(openOdsWorkbook(openZip(noTables))).rejects.toMatchObject({
      code: 'PARSE_FAILED',
    });
  });

  it('renders through the sheet view', async () => {
    const { context, container } = renderContext(odsFixture(), 'book.ods', 'ods');
    await odsRenderer.render(context);
    await waitFor(() => cellTexts(container).includes('Widget'));
    expect(container.querySelectorAll('.av-tab')).toHaveLength(2);
  });
});

describe('CSV', () => {
  it('renders rows in a grid without tabs', async () => {
    const { context, container } = renderContext('name;qty\nWidget;3\n', 'stock.csv', 'csv');
    await csvRenderer.render(context);
    await waitFor(() => cellTexts(container).length > 0);
    expect(cellTexts(container)).toEqual(['name', 'qty', 'Widget', '3']);
    expect(container.querySelector('.av-tabs')).toBeNull();
  });

  it('forces tab separation for TSV', async () => {
    const { context, container } = renderContext('a,b\tc\n', 'data.tsv', 'tsv');
    await csvRenderer.render(context);
    await waitFor(() => cellTexts(container).length > 0);
    expect(cellTexts(container)).toEqual(['a,b', 'c']);
  });

  it('shows a notice when rows are truncated', async () => {
    const big = 'x\n'.repeat(100_050);
    const { context, container } = renderContext(big, 'big.csv', 'csv');
    await csvRenderer.render(context);
    const notice = await waitFor(() => {
      const element = container.querySelector<HTMLElement>('.av-notice');
      return element && !element.hidden ? element : null;
    });
    expect(notice.textContent).toContain('100,000');
    expect(notice.textContent).toContain('100,050');
  });

  it('estimates column widths within bounds', () => {
    expect(estimateColumnWidths([['a', 'x'.repeat(500)], ['bbbbbbbbbbbbbbbb']], 2)).toEqual([
      142, 360,
    ]);
    expect(estimateColumnWidths([], 1)).toEqual([64]);
  });
});
