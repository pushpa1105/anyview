import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { afterEach, describe, expect, it } from 'vitest';
import { sniffFormat } from '../../src';
import { renderDocx } from '../../src/formats/docx';
import { openPresentation } from '../../src/formats/pptx';
import { openWorkbook } from '../../src/formats/xlsx';
import { openZip } from '../../src/util/zip';

/**
 * These files are written by ExcelJS, docx and PptxGenJS in the global setup
 * (see generate-fixtures.ts), so they exercise the parsers on output from
 * independent implementations rather than on hand-written XML.
 */
const read = (name: string): Uint8Array =>
  new Uint8Array(readFileSync(resolve(__dirname, '../.generated', name)));

const host = { objectUrl: (): string => 'blob:test/image' };

afterEach(() => {
  document.body.textContent = '';
});

describe('real XLSX (written by ExcelJS)', () => {
  it('is detected from its content', () => {
    expect(sniffFormat(read('real.xlsx'))?.id).toBe('xlsx');
  });

  it('reads values, number formats, dates, formulas and styles', async () => {
    const sheets = await openWorkbook(openZip(read('real.xlsx')));
    expect(sheets.map((sheet) => sheet.name)).toEqual(['Sales', 'Notes']);
    const grid = await sheets[0]!.load();
    const row = (index: number): (string | undefined)[] =>
      [0, 1, 2, 3].map((col) => grid.cell(index, col)?.text);

    expect(row(0)).toEqual(['Product', 'Units', 'Price', 'Shipped']);
    expect(row(1)).toEqual(['Widget', '1,200', '19.50', '2024-01-31']);
    expect(row(2)).toEqual(['Gadget', '85', '1,234.50', '2023-12-25']);
    expect(grid.cell(3, 0)?.text).toBe('Total');
    expect(grid.cell(3, 1)?.text).toBe('1285');

    expect(grid.cell(0, 0)?.style).toMatchObject({ bold: true, background: '#1F4E78' });
    expect(grid.cell(1, 1)?.style?.align).toBe('right');
    expect(grid.merges).toEqual([{ row: 5, col: 0, rowSpan: 1, colSpan: 3 }]);
    expect(grid.cell(5, 0)).toMatchObject({ text: 'Merged note', style: { align: 'center' } });
    expect(grid.colWidths?.[0]).toBeGreaterThan(150);
  });

  it('reads Unicode text, booleans and percentages on other sheets', async () => {
    const sheets = await openWorkbook(openZip(read('real.xlsx')));
    const grid = await sheets[1]!.load();
    expect(grid.cell(0, 0)?.text).toBe('Unicode: café ✓ 日本語');
    expect(grid.cell(1, 0)?.text).toBe('TRUE');
    expect(grid.cell(1, 1)?.text).toBe('25.6%');
  });
});

describe('real DOCX (written by docx)', () => {
  it('is detected from its content', () => {
    expect(sniffFormat(read('real.docx'))?.id).toBe('docx');
  });

  it('renders headings, formatting, lists, links, tables and images', async () => {
    const paper = (await renderDocx(openZip(read('real.docx')), host)).element;
    expect(paper.querySelector('h1')?.textContent).toBe('Project Report');

    const paragraph = Array.from(paper.querySelectorAll<HTMLElement>('p')).find((element) =>
      element.textContent?.startsWith('This is'),
    )!;
    expect(paragraph.textContent).toBe('This is bold and italic red text.');
    expect(paragraph.style.textAlign).toBe('center');
    const spans = Array.from(paragraph.querySelectorAll<HTMLElement>('span'));
    expect(spans.find((span) => span.textContent === 'bold')!.style.fontWeight).toBe('bold');
    const red = spans.find((span) => span.textContent === 'italic red')!;
    expect(red.style.fontStyle).toBe('italic');
    expect(red.style.color.toLowerCase()).toBe('#ff0000');

    const text = paper.textContent ?? '';
    expect(text).toContain('1.First step');
    expect(text).toContain('2.Second step');
    expect(text).toMatch(/[•●▪]A bullet/);

    const link = paper.querySelector('a')!;
    expect(link.getAttribute('href')).toBe('https://example.com/docs');
    expect(link.textContent).toBe('the documentation');

    const rows = paper.querySelectorAll('table tr');
    expect(rows).toHaveLength(2);
    expect(rows[0]!.textContent).toBe('NameValue');
    expect((rows[1]!.children[0] as HTMLTableCellElement).colSpan).toBe(2);

    const image = paper.querySelector('img')!;
    expect(image.getAttribute('src')).toBe('blob:test/image');
    expect(parseFloat(image.style.width)).toBeCloseTo(120, 0);
    expect(parseFloat(image.style.height)).toBeCloseTo(60, 0);
  });
});

describe('real PPTX (written by PptxGenJS)', () => {
  it('is detected from its content', () => {
    expect(sniffFormat(read('real.pptx'))?.id).toBe('pptx');
  });

  it('reads slide size and renders text, shapes, tables and pictures', async () => {
    const deck = await openPresentation(openZip(read('real.pptx')), host);
    expect(deck.count).toBe(2);
    expect(deck.width / deck.height).toBeCloseTo(16 / 9, 2);

    const slide = await deck.renderSlide(0);
    expect(slide.style.background.toLowerCase()).toContain('#f1f1f1');
    const shapes = Array.from(slide.querySelectorAll<HTMLElement>('.av-slide-shape'));

    const title = shapes.find((shape) => shape.textContent === 'Quarterly Results')!;
    expect(parseFloat(title.style.left)).toBeCloseTo(48, 0);
    expect(parseFloat(title.style.width)).toBeCloseTo(864, 0);
    const run = title.querySelector<HTMLElement>('span')!;
    expect(run.style.fontSize).toBe('36pt');
    expect(run.style.fontWeight).toBe('bold');
    expect(run.style.color.toLowerCase()).toBe('#203864');

    const bullets = shapes.find((shape) => shape.textContent?.includes('Revenue up'))!;
    expect(bullets.querySelectorAll('p')).toHaveLength(2);
    expect(bullets.textContent).toMatch(/.Revenue up.Costs down/);

    const ellipse = shapes.find((shape) => shape.style.borderRadius === '50%')!;
    expect(ellipse.style.background.toLowerCase()).toBe('#c00000');
    expect(ellipse.style.border).toContain('solid');

    expect(slide.querySelector('table')?.textContent).toBe('RegionSalesNorth120');
    expect(slide.querySelector('img')?.getAttribute('src')).toBe('blob:test/image');

    const second = await deck.renderSlide(1);
    expect(second.textContent).toContain('Thank you');
    expect(second.querySelector<HTMLElement>('p')!.style.textAlign).toBe('center');
  });
});
