/**
 * Vitest global setup: writes real Office files with independent third-party
 * libraries (dev dependencies only). The integration tests then open them with
 * AnyView, which checks the parsers against files they did not produce themselves.
 */
import { mkdirSync, writeFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import {
  AlignmentType,
  Document,
  ExternalHyperlink,
  HeadingLevel,
  ImageRun,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
} from 'docx';
import ExcelJS from 'exceljs';
import PptxGenJSModule from 'pptxgenjs';
import { sampleFiles } from './samples';

export const GENERATED_DIR = resolve(dirname(fileURLToPath(import.meta.url)), '../.generated');

const PNG_BASE64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==';

async function writeXlsx(): Promise<void> {
  const workbook = new ExcelJS.Workbook();
  const sales = workbook.addWorksheet('Sales');
  sales.columns = [{ width: 24 }, { width: 14 }, { width: 14 }, { width: 14 }];
  sales.addRow(['Product', 'Units', 'Price', 'Shipped']);
  sales.addRow(['Widget', 1200, 19.5, new Date(Date.UTC(2024, 0, 31))]);
  sales.addRow(['Gadget', 85, 1234.5, new Date(Date.UTC(2023, 11, 25))]);
  sales.addRow(['Total', { formula: 'SUM(B2:B3)', result: 1285 }, null, null]);
  sales.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
  sales.getCell('A1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E78' } };
  sales.getColumn(3).numFmt = '#,##0.00';
  sales.getColumn(4).numFmt = 'yyyy-mm-dd';
  sales.getCell('B2').numFmt = '#,##0';
  sales.getCell('A6').value = 'Merged note';
  sales.mergeCells('A6:C6');
  sales.getCell('A6').alignment = { horizontal: 'center' };

  const notes = workbook.addWorksheet('Notes');
  notes.addRow(['Unicode: café ✓ 日本語']);
  notes.addRow([true, 0.256]);
  notes.getCell('B2').numFmt = '0.0%';

  const buffer = await workbook.xlsx.writeBuffer();
  writeFileSync(resolve(GENERATED_DIR, 'real.xlsx'), new Uint8Array(buffer));
}

async function writeDocx(): Promise<void> {
  const document = new Document({
    numbering: {
      config: [
        {
          reference: 'steps',
          levels: [{ level: 0, format: 'decimal', text: '%1.', alignment: AlignmentType.START }],
        },
      ],
    },
    sections: [
      {
        children: [
          new Paragraph({ text: 'Project Report', heading: HeadingLevel.HEADING_1 }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun('This is '),
              new TextRun({ text: 'bold', bold: true }),
              new TextRun(' and '),
              new TextRun({ text: 'italic red', italics: true, color: 'FF0000' }),
              new TextRun(' text.'),
            ],
          }),
          new Paragraph({ text: 'First step', numbering: { reference: 'steps', level: 0 } }),
          new Paragraph({ text: 'Second step', numbering: { reference: 'steps', level: 0 } }),
          new Paragraph({ text: 'A bullet', bullet: { level: 0 } }),
          new Paragraph({
            children: [
              new ExternalHyperlink({
                link: 'https://example.com/docs',
                children: [new TextRun({ text: 'the documentation', style: 'Hyperlink' })],
              }),
            ],
          }),
          new Table({
            rows: [
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph('Name')] }),
                  new TableCell({ children: [new Paragraph('Value')] }),
                ],
              }),
              new TableRow({
                children: [new TableCell({ children: [new Paragraph('Alpha')], columnSpan: 2 })],
              }),
            ],
          }),
          new Paragraph({
            children: [
              new ImageRun({
                type: 'png',
                data: Buffer.from(PNG_BASE64, 'base64'),
                transformation: { width: 120, height: 60 },
              }),
            ],
          }),
        ],
      },
    ],
  });
  writeFileSync(resolve(GENERATED_DIR, 'real.docx'), await Packer.toBuffer(document));
}

async function writePptx(): Promise<void> {
  // The package's default export differs between its CJS and ESM builds.
  const PptxGenJS =
    (PptxGenJSModule as unknown as { default?: typeof PptxGenJSModule }).default ?? PptxGenJSModule;
  const presentation = new PptxGenJS();
  presentation.layout = 'LAYOUT_16x9';

  const first = presentation.addSlide();
  first.background = { color: 'F1F1F1' };
  first.addText('Quarterly Results', {
    x: 0.5,
    y: 0.4,
    w: 9,
    h: 1,
    fontSize: 36,
    bold: true,
    color: '203864',
  });
  first.addText(
    [
      { text: 'Revenue up', options: { bullet: true } },
      { text: 'Costs down', options: { bullet: true, italic: true } },
    ],
    { x: 0.5, y: 1.6, w: 5, h: 1.5, fontSize: 18 },
  );
  first.addShape(presentation.ShapeType.ellipse, {
    x: 6.5,
    y: 1.6,
    w: 2,
    h: 2,
    fill: { color: 'C00000' },
    line: { color: '000000', width: 2 },
  });
  first.addImage({ data: `image/png;base64,${PNG_BASE64}`, x: 6.5, y: 4, w: 1, h: 1 });
  first.addTable(
    [
      [{ text: 'Region' }, { text: 'Sales' }],
      [{ text: 'North' }, { text: '120' }],
    ],
    { x: 0.5, y: 3.5, w: 5, colW: [2.5, 2.5] },
  );

  const second = presentation.addSlide();
  second.addText('Thank you', { x: 1, y: 2, w: 8, h: 1.5, fontSize: 44, align: 'center' });

  const buffer = (await presentation.write({ outputType: 'nodebuffer' })) as Buffer;
  writeFileSync(resolve(GENERATED_DIR, 'real.pptx'), buffer);
}

export default async function setup(): Promise<void> {
  mkdirSync(GENERATED_DIR, { recursive: true });
  for (const [name, content] of Object.entries(sampleFiles())) {
    writeFileSync(resolve(GENERATED_DIR, name), content);
  }
  await Promise.all([writeXlsx(), writeDocx(), writePptx()]);
}
