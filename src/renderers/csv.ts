import type { Renderer } from '../core/types';
import type { GridSource } from '../ui/grid';
import { createSheetView } from '../ui/sheets';
import { parseCsv } from '../util/csv';

/** Rows kept in memory and shown; larger files are truncated with a notice. */
export const MAX_ROWS = 100_000;

/** Estimates column widths from the first rows so short columns stay narrow. */
export function estimateColumnWidths(
  rows: readonly (readonly string[])[],
  columnCount: number,
): number[] {
  const widths = new Array<number>(columnCount).fill(0);
  const sample = Math.min(rows.length, 200);
  for (let row = 0; row < sample; row++) {
    const cells = rows[row]!;
    for (let col = 0; col < cells.length; col++) {
      const length = cells[col]!.length;
      if (length > widths[col]!) widths[col] = length;
    }
  }
  return widths.map((length) => Math.min(360, Math.max(64, Math.round(length * 7.6 + 20))));
}

export function rowsToGrid(rows: string[][], columnCount: number, totalRows: number): GridSource {
  return {
    rowCount: rows.length,
    colCount: columnCount,
    totalRows,
    colWidths: estimateColumnWidths(rows, columnCount),
    cell(row, col) {
      const text = rows[row]?.[col];
      return text ? { text } : null;
    },
  };
}

export const renderer: Renderer = {
  async render(context) {
    const text = await context.file.text();
    const parsed = parseCsv(text, {
      maxRows: MAX_ROWS,
      ...(context.format.id === 'tsv' ? { delimiter: '\t' } : {}),
    });
    return createSheetView(context, [
      {
        name: context.file.name || context.format.label,
        load: () => rowsToGrid(parsed.rows, parsed.columnCount, parsed.totalRows),
      },
    ]);
  },
};
