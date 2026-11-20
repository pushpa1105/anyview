import { AnyViewError } from '../core/errors';
import type { GridCell, GridMerge, GridSource } from '../ui/grid';
import { attr, childElements, find, type XmlElement } from '../util/xml';
import type { ZipArchive } from '../util/zip';
import { odfText, requireOdfContent } from './odf';
import { MAX_SHEET_COLS, MAX_SHEET_ROWS, type WorkbookSheet } from './xlsx';

function cellText(cell: XmlElement): string {
  return childElements(cell, 'p').map(odfText).join('\n');
}

function repeatCount(element: XmlElement, name: string): number {
  const value = Number(attr(element, name) ?? 1);
  return Number.isInteger(value) && value > 0 ? value : 1;
}

function parseTable(table: XmlElement): GridSource {
  const rows: (GridCell | undefined)[][] = [];
  const merges: GridMerge[] = [];
  let colCount = 0;
  let rowIndex = 0;
  let lastUsedRow = -1;

  const visit = (parent: XmlElement): void => {
    for (const element of childElements(parent)) {
      // Row groups and header rows wrap ordinary rows.
      if (
        element.local === 'table-row-group' ||
        element.local === 'table-header-rows' ||
        element.local === 'table-rows'
      ) {
        visit(element);
        continue;
      }
      if (element.local !== 'table-row') continue;
      const repeat = repeatCount(element, 'table:number-rows-repeated');
      const cells: (GridCell | undefined)[] = [];
      let col = 0;
      for (const cell of childElements(element)) {
        if (cell.local !== 'table-cell' && cell.local !== 'covered-table-cell') continue;
        const columns = repeatCount(cell, 'table:number-columns-repeated');
        const text = cell.local === 'table-cell' ? cellText(cell) : '';
        if (text !== '' && col < MAX_SHEET_COLS) {
          const numeric = ['float', 'currency', 'percentage', 'date', 'time'].includes(
            attr(cell, 'office:value-type') ?? '',
          );
          const value: GridCell = numeric ? { text, style: { align: 'right' } } : { text };
          // Spreadsheets pad to the sheet edge with huge repeat counts; only real content is copied.
          const copies = Math.min(columns, MAX_SHEET_COLS - col);
          for (let i = 0; i < copies; i++) cells[col + i] = value;
          colCount = Math.max(colCount, col + copies);
        }
        const rowSpan = repeatCount(cell, 'table:number-rows-spanned');
        const colSpan = repeatCount(cell, 'table:number-columns-spanned');
        if ((rowSpan > 1 || colSpan > 1) && col < MAX_SHEET_COLS) {
          merges.push({ row: rowIndex, col, rowSpan, colSpan });
          colCount = Math.max(colCount, Math.min(col + colSpan, MAX_SHEET_COLS));
        }
        col += columns;
      }
      if (cells.length > 0) {
        const copies = Math.min(repeat, Math.max(0, MAX_SHEET_ROWS - rowIndex));
        for (let i = 0; i < copies; i++) rows[rowIndex + i] = cells;
        lastUsedRow = rowIndex + repeat - 1;
      }
      rowIndex += repeat;
    }
  };
  visit(table);

  const totalRows = lastUsedRow + 1;
  return {
    rowCount: Math.max(1, Math.min(totalRows, MAX_SHEET_ROWS)),
    colCount: Math.max(colCount, 1),
    totalRows: Math.max(totalRows, 1),
    merges: merges.filter((merge) => merge.row <= lastUsedRow),
    cell: (row, col) => rows[row]?.[col] ?? null,
  };
}

/** Opens an OpenDocument spreadsheet. Each table becomes one sheet. */
export async function openOdsWorkbook(zip: ZipArchive): Promise<WorkbookSheet[]> {
  const content = await requireOdfContent(zip, 'OpenDocument spreadsheet');
  const spreadsheet = find(content, 'body', 'spreadsheet');
  const tables = childElements(spreadsheet, 'table');
  if (tables.length === 0) {
    throw new AnyViewError('PARSE_FAILED', 'This spreadsheet does not contain any sheets.');
  }
  return tables.map((table, index) => {
    let parsed: GridSource | null = null;
    return {
      name: attr(table, 'table:name') ?? `Sheet${index + 1}`,
      load: () => Promise.resolve((parsed ??= parseTable(table))),
    };
  });
}
