import { AnyViewError } from '../core/errors';
import { renderChart } from '../ui/chart';
import { h } from '../ui/dom';
import type {
  GridBorder,
  GridCell,
  GridCellStyle,
  GridMerge,
  GridOverlay,
  GridRun,
  GridSource,
} from '../ui/grid';
import { formatText, formatValue } from '../util/numfmt';
import { toBlobPart } from '../util/text';
import {
  attr,
  child,
  childElements,
  descendants,
  find,
  parseXml,
  textOf,
  type XmlElement,
} from '../util/xml';
import type { ZipArchive } from '../util/zip';
import { parseChart } from './chart';
import {
  emuToPx,
  imageMimeType,
  readRelationships,
  readXmlPart,
  requireXmlPart,
  type Relationship,
} from './ooxml';
import { loadTheme } from './theme';
import {
  CF_STRIDE,
  evaluateConditionalFormats,
  parseRange,
  type CfEffect,
  type Rect,
} from './xlsx-cf';
import {
  applyFont,
  builtinTableStyle,
  parseStyles,
  pointsToPx,
  type BorderDef,
  type CellXf,
  type Dxf,
  type StyleSheet,
  type TableStyleDef,
} from './xlsx-styles';

export interface WorkbookSheet {
  name: string;
  /** Parses the sheet on demand. */
  load(): Promise<GridSource>;
}

/** What the workbook needs from the viewer to show embedded pictures. */
export interface WorkbookHost {
  /** Registers a blob and returns a URL that the viewer revokes later. */
  objectUrl(blob: Blob): string;
}

/** Rows and columns kept per sheet; anything beyond is reported as truncated. */
export const MAX_SHEET_ROWS = 100_000;
export const MAX_SHEET_COLS = 2_000;
/** Pictures and charts drawn per sheet. */
const MAX_OVERLAYS = 200;

/** `"BC12"` → `{ row: 11, col: 54 }` (zero-based). Returns `null` for invalid references. */
export function parseCellRef(reference: string): { row: number; col: number } | null {
  let col = 0;
  let i = 0;
  for (; i < reference.length; i++) {
    const code = reference.charCodeAt(i) & ~0x20;
    if (code < 65 || code > 90) break;
    col = col * 26 + (code - 64);
  }
  if (i === 0 || i === reference.length) return null;
  const row = Number(reference.slice(i).replace(/\$/g, ''));
  if (!Number.isInteger(row) || row < 1) return null;
  return { row: row - 1, col: col - 1 };
}

interface SharedString {
  text: string;
  runs?: GridRun[];
}

/** A cell as stored in the sheet, before styles are resolved. */
interface RawCell {
  text: string;
  xf: number;
  /** Numeric value of number cells. */
  value?: number;
  kind: 'text' | 'number' | 'blank' | 'other';
  error?: boolean;
  runs?: GridRun[];
  fill?: { left: string; right: string };
  /** Colour requested by the number format. */
  color?: string;
}

interface TableInfo extends Rect {
  headerRows: number;
  totalRows: number;
  style: TableStyleDef | undefined;
  rowStripes: boolean;
  columnStripes: boolean;
  firstColumn: boolean;
  lastColumn: boolean;
  filter: boolean;
}

const isTrue = (value: string | undefined): boolean => value === '1' || value === 'true';

/** Text of a string item: plain `<t>` or rich-text runs; phonetic hints are skipped. */
function parseStringItem(item: XmlElement, styles: StyleSheet): SharedString {
  let text = '';
  const runs: GridRun[] = [];
  let rich = false;
  for (const node of item.children) {
    if (typeof node === 'string') continue;
    if (node.local === 't') {
      const part = textOf(node);
      text += part;
      runs.push({ text: part });
    } else if (node.local === 'r') {
      const part = textOf(child(node, 't'));
      text += part;
      const properties = child(node, 'rPr');
      if (properties) {
        const style: GridCellStyle = {};
        applyFont(style, styles.font(properties));
        runs.push({ text: part, style });
        rich = true;
      } else {
        runs.push({ text: part });
      }
    }
  }
  return rich ? { text, runs } : { text };
}

/** Width of the widest digit of the default font, the unit of column widths. */
function digitWidth(family: string | undefined, sizePx: number): number {
  const fallback = Math.max(1, Math.round(sizePx * 0.48));
  try {
    const context = document.createElement('canvas').getContext('2d');
    if (!context) return fallback;
    context.font = `${sizePx}px ${family ? `"${family}", ` : ''}Calibri, Carlito, Arial, sans-serif`;
    let widest = 0;
    for (const digit of '0123456789') {
      widest = Math.max(widest, context.measureText(digit).width);
    }
    return widest > 0 ? Math.round(widest) : fallback;
  } catch {
    return fallback;
  }
}

const border = (side: GridBorder | null | undefined): GridBorder | undefined => side ?? undefined;

/** Applies a differential border to a cell at the given position inside its region. */
function applyRegionBorder(
  style: GridCellStyle,
  definition: BorderDef | undefined,
  edges: { top: boolean; bottom: boolean; left: boolean; right: boolean },
): void {
  if (!definition) return;
  const sides = [
    ['borderTop', edges.top ? definition.top : definition.horizontal],
    ['borderBottom', edges.bottom ? definition.bottom : definition.horizontal],
    ['borderLeft', edges.left ? definition.left : definition.vertical],
    ['borderRight', edges.right ? definition.right : definition.vertical],
  ] as const;
  for (const [name, side] of sides) {
    if (side === undefined) continue;
    if (side === null) delete style[name];
    else style[name] = side;
  }
}

function applyDxf(
  style: GridCellStyle,
  dxf: Dxf,
  allow: { font: boolean; fill: boolean; border: boolean },
  edges = { top: true, bottom: true, left: true, right: true },
): void {
  if (allow.font) applyFont(style, dxf.font);
  if (allow.fill && typeof dxf.fill === 'string') style.background = dxf.fill;
  if (allow.border) applyRegionBorder(style, dxf.border, edges);
}

class SheetParser {
  private readonly rows: (RawCell | undefined)[][] = [];
  private readonly rowXf: (number | undefined)[] = [];
  private readonly colXf: (number | undefined)[] = [];
  private readonly rowHeights: (number | undefined)[] = [];
  private readonly colWidths: (number | undefined)[] = [];
  private readonly merges: GridMerge[] = [];
  private readonly mergeEnds = new Map<number, { row: number; col: number }>();
  private readonly tables: TableInfo[] = [];
  private readonly overlays: GridOverlay[] = [];
  private effects = new Map<number, CfEffect>();
  private readonly baseStyles = new Map<string, GridCellStyle>();
  private usedRows = 0;
  private usedCols = 0;
  private totalRows = 0;

  constructor(
    private readonly root: XmlElement,
    private readonly sharedStrings: readonly SharedString[],
    private readonly styles: StyleSheet,
    private readonly date1904: boolean,
  ) {}

  private touch(row: number, col: number): void {
    if (row + 1 > this.usedRows) this.usedRows = Math.min(row + 1, MAX_SHEET_ROWS);
    if (col + 1 > this.usedCols) this.usedCols = Math.min(col + 1, MAX_SHEET_COLS);
  }

  parseCells(): void {
    const { styles, sharedStrings, date1904 } = this;
    let nextRow = 0;
    for (const rowElement of childElements(child(this.root, 'sheetData'), 'row')) {
      const declared = Number(attr(rowElement, 'r'));
      const rowIndex = Number.isInteger(declared) && declared >= 1 ? declared - 1 : nextRow;
      nextRow = rowIndex + 1;
      if (rowIndex + 1 > this.totalRows) this.totalRows = rowIndex + 1;
      if (rowIndex >= MAX_SHEET_ROWS) continue;

      const height = Number(attr(rowElement, 'ht'));
      if (isTrue(attr(rowElement, 'hidden'))) this.rowHeights[rowIndex] = 0;
      else if (height > 0) this.rowHeights[rowIndex] = pointsToPx(height);
      if (isTrue(attr(rowElement, 'customFormat'))) {
        const rowStyle = Number(attr(rowElement, 's'));
        if (rowStyle > 0) this.rowXf[rowIndex] = rowStyle;
      }

      const cells: (RawCell | undefined)[] = [];
      let nextCol = 0;
      for (const cellElement of childElements(rowElement, 'c')) {
        const reference = attr(cellElement, 'r');
        const position = reference ? parseCellRef(reference) : null;
        const col = position ? position.col : nextCol;
        nextCol = col + 1;
        if (col >= MAX_SHEET_COLS) continue;

        const type = attr(cellElement, 't') ?? 'n';
        const xf = Number(attr(cellElement, 's') ?? 0) || 0;
        const format = styles.cellXfs[xf]?.numberFormat ?? 'General';
        const raw = textOf(child(cellElement, 'v'));
        const cell: RawCell = { text: '', xf, kind: 'blank' };
        if (type === 's') {
          const item = sharedStrings[Number(raw)];
          cell.text = formatText(item?.text ?? '', format);
          if (item?.runs && cell.text === item.text) cell.runs = item.runs;
          cell.kind = 'text';
        } else if (type === 'inlineStr') {
          const inline = child(cellElement, 'is');
          const item = inline ? parseStringItem(inline, styles) : { text: '' };
          cell.text = item.text;
          if (item.runs) cell.runs = item.runs;
          cell.kind = 'text';
        } else if (type === 'str') {
          cell.text = raw;
          cell.kind = 'text';
        } else if (type === 'b') {
          cell.text = raw === '1' ? 'TRUE' : 'FALSE';
          cell.kind = 'other';
        } else if (type === 'e') {
          cell.text = raw;
          cell.error = true;
          cell.kind = 'other';
        } else if (type === 'd') {
          cell.text = raw.replace('T', ' ').replace(/Z$/, '');
          cell.kind = 'number';
        } else if (raw !== '') {
          const value = Number(raw);
          cell.kind = 'number';
          if (Number.isFinite(value)) {
            const formatted = formatValue(value, format, date1904);
            cell.text = formatted.text;
            cell.value = value;
            if (formatted.fill) cell.fill = formatted.fill;
            if (formatted.color) cell.color = formatted.color;
          } else {
            cell.text = raw;
          }
        }
        // Cells without a value still matter when they carry formatting.
        if (cell.text === '' && xf === 0) continue;
        cells[col] = cell;
        if (cell.text !== '' || styles.cellXfs[xf]?.fill || this.hasBorder(xf)) {
          this.touch(rowIndex, col);
        }
      }
      if (cells.length > 0) this.rows[rowIndex] = cells;
    }
  }

  private hasBorder(xf: number): boolean {
    const definition = this.styles.cellXfs[xf]?.border;
    return Boolean(
      definition && (definition.top || definition.bottom || definition.left || definition.right),
    );
  }

  parseLayout(): { defaultRowHeight: number; defaultColWidth: number; showGridLines: boolean } {
    const { root, styles } = this;
    const fontSize = pointsToPx(styles.defaultFont.size ?? 11);
    const digit = digitWidth(styles.defaultFont.name, fontSize);
    // Column widths count characters of the default font, including cell padding.
    const charsToPx = (chars: number): number =>
      Math.floor(((256 * chars + Math.floor(128 / digit)) / 256) * digit);

    const format = child(root, 'sheetFormatPr');
    const defaultHeight = Number(attr(format, 'defaultRowHeight'));
    const defaultChars = Number(attr(format, 'defaultColWidth'));
    const baseChars = Number(attr(format, 'baseColWidth') ?? 8) || 8;
    const defaultColWidth =
      defaultChars > 0 ? charsToPx(defaultChars) : Math.ceil((baseChars * digit + 5) / 8) * 8;
    const defaultRowHeight = isTrue(attr(format, 'zeroHeight'))
      ? 0
      : defaultHeight > 0
        ? pointsToPx(defaultHeight)
        : Math.round(fontSize * 1.36);

    for (const column of childElements(child(root, 'cols'), 'col')) {
      const min = Number(attr(column, 'min'));
      const max = Math.min(Number(attr(column, 'max')), MAX_SHEET_COLS);
      if (!Number.isFinite(min) || !Number.isFinite(max) || min < 1) continue;
      const width = Number(attr(column, 'width'));
      const hidden = isTrue(attr(column, 'hidden'));
      const columnStyle = Number(attr(column, 'style'));
      // A definition that runs to the last column sets the default for the rest of the sheet.
      const last = Math.min(max, Math.max(this.usedCols + 3, min));
      for (let col = min; col <= last; col++) {
        if (hidden) this.colWidths[col - 1] = 0;
        else if (width > 0) this.colWidths[col - 1] = charsToPx(width);
        if (columnStyle > 0) this.colXf[col - 1] = columnStyle;
      }
      if ((hidden || width > 0) && max - min < 64) this.touch(0, last - 1);
    }

    for (const merge of childElements(child(root, 'mergeCells'), 'mergeCell')) {
      const [from, to] = (attr(merge, 'ref') ?? '').split(':');
      const start = from ? parseCellRef(from) : null;
      const end = to ? parseCellRef(to) : null;
      if (!start || !end || start.row >= MAX_SHEET_ROWS || start.col >= MAX_SHEET_COLS) continue;
      const last = {
        row: Math.min(end.row, MAX_SHEET_ROWS - 1),
        col: Math.min(end.col, MAX_SHEET_COLS - 1),
      };
      this.merges.push({
        row: start.row,
        col: start.col,
        rowSpan: last.row - start.row + 1,
        colSpan: last.col - start.col + 1,
      });
      this.mergeEnds.set(start.row * CF_STRIDE + start.col, last);
      this.touch(last.row, last.col);
    }

    const view = find(root, 'sheetViews', 'sheetView');
    const showGridLines = !['0', 'false'].includes(attr(view, 'showGridLines') ?? '1');
    return { defaultRowHeight, defaultColWidth, showGridLines };
  }

  addTable(table: XmlElement | null): void {
    const { styles } = this;
    const range = parseRange(attr(table, 'ref') ?? '', MAX_SHEET_ROWS, MAX_SHEET_COLS);
    if (!table || !range) return;
    const info = child(table, 'tableStyleInfo');
    const name = attr(info, 'name');
    const headerRows = Number(attr(table, 'headerRowCount') ?? 1) || 0;
    this.tables.push({
      ...range,
      headerRows: Math.max(0, Math.min(headerRows, 1)),
      totalRows:
        isTrue(attr(table, 'totalsRowShown')) || attr(table, 'totalsRowCount') === '1' ? 1 : 0,
      style: name
        ? (styles.tableStyles.get(name) ?? builtinTableStyle(name, styles.theme))
        : undefined,
      rowStripes: isTrue(attr(info, 'showRowStripes')),
      columnStripes: isTrue(attr(info, 'showColumnStripes')),
      firstColumn: isTrue(attr(info, 'showFirstColumn')),
      lastColumn: isTrue(attr(info, 'showLastColumn')),
      filter: child(table, 'autoFilter') !== null && headerRows > 0,
    });
    this.touch(range.bottom, range.right);
  }

  addOverlay(overlay: GridOverlay, corner: { row: number; col: number }): void {
    if (this.overlays.length >= MAX_OVERLAYS) return;
    this.overlays.push(overlay);
    this.touch(corner.row, corner.col);
  }

  applyConditionalFormats(): void {
    this.effects = evaluateConditionalFormats(this.root, this.styles, {
      rowCount: this.usedRows,
      colCount: this.usedCols,
      cell: (row, col) => this.rows[row]?.[col],
    });
  }

  /** Style of a cell from its cell format alone; shared between cells with the same format. */
  private baseStyle(xfIndex: number, kind: RawCell['kind']): GridCellStyle {
    const key = `${xfIndex}:${kind}`;
    let style = this.baseStyles.get(key);
    if (style) return style;
    style = {};
    const xf: CellXf | undefined = this.styles.cellXfs[xfIndex];
    if (xf) {
      applyFont(style, xf.font);
      // The grid already uses the default font; repeating it on every cell is wasted work.
      const base = this.styles.defaultFont;
      if (xf.font.size === base.size) delete style.fontSize;
      if (xf.font.name === base.name) delete style.fontFamily;
      if (xf.fill) style.background = xf.fill;
      const top = border(xf.border.top);
      const right = border(xf.border.right);
      const bottom = border(xf.border.bottom);
      const left = border(xf.border.left);
      if (top) style.borderTop = top;
      if (right) style.borderRight = right;
      if (bottom) style.borderBottom = bottom;
      if (left) style.borderLeft = left;
    }
    const horizontal = xf?.align.horizontal;
    if (horizontal === 'left' || horizontal === 'right' || horizontal === 'center') {
      style.align = horizontal;
    } else if (horizontal === 'centerContinuous') {
      style.align = 'center';
    } else if (horizontal === undefined || horizontal === 'general') {
      // General alignment: numbers right, text left, booleans and errors centred.
      if (kind === 'number') style.align = 'right';
      else if (kind === 'other') style.align = 'center';
    }
    const vertical = xf?.align.vertical;
    style.valign =
      vertical === 'top'
        ? 'top'
        : vertical === 'center' || vertical === 'justify' || vertical === 'distributed'
          ? 'middle'
          : 'bottom';
    if (xf?.align.wrap) style.wrap = true;
    if (xf?.align.indent) style.indent = xf.align.indent * 9;
    this.baseStyles.set(key, style);
    return style;
  }

  private tableAt(row: number, col: number): TableInfo | undefined {
    for (const table of this.tables) {
      if (row >= table.top && row <= table.bottom && col >= table.left && col <= table.right) {
        return table;
      }
    }
    return undefined;
  }

  /** Adds the formatting a table style gives to the cell at (row, col). */
  private applyTable(
    style: GridCellStyle,
    table: TableInfo,
    row: number,
    col: number,
    xf: CellXf | undefined,
  ): void {
    const elements = table.style?.elements;
    if (!elements) return;
    const allow = {
      font: !xf?.explicitFont,
      fill: !xf?.explicitFill,
      border: !xf?.explicitBorder,
    };
    const dataTop = table.top + table.headerRows;
    const dataBottom = table.bottom - table.totalRows;
    const isHeader = row < dataTop;
    const isTotal = row > dataBottom;
    const atLeft = col === table.left;
    const atRight = col === table.right;
    const use = (
      type: string,
      edges: { top: boolean; bottom: boolean; left: boolean; right: boolean },
    ): void => {
      const element = elements.get(type);
      if (element) applyDxf(style, element.dxf, allow, edges);
    };
    const stripe = (index: number, first: string, second: string): string => {
      const firstSize = elements.get(first)?.size ?? 1;
      const secondSize = elements.get(second)?.size ?? 1;
      return index % (firstSize + secondSize) < firstSize ? first : second;
    };

    use('wholeTable', {
      top: row === table.top,
      bottom: row === table.bottom,
      left: atLeft,
      right: atRight,
    });
    if (!isHeader && !isTotal) {
      const rowEdges = { top: true, bottom: true, left: atLeft, right: atRight };
      const colEdges = {
        top: row === dataTop,
        bottom: row === dataBottom,
        left: true,
        right: true,
      };
      if (table.columnStripes) {
        use(stripe(col - table.left, 'firstColumnStripe', 'secondColumnStripe'), colEdges);
      }
      if (table.rowStripes) {
        use(stripe(row - dataTop, 'firstRowStripe', 'secondRowStripe'), rowEdges);
      }
      if (table.lastColumn && atRight) use('lastColumn', colEdges);
      if (table.firstColumn && atLeft) use('firstColumn', colEdges);
    }
    const lineEdges = { top: true, bottom: true, left: atLeft, right: atRight };
    if (isHeader) use('headerRow', lineEdges);
    if (isTotal) use('totalRow', lineEdges);
  }

  private cellAt = (row: number, col: number): GridCell | null => {
    const raw = this.rows[row]?.[col];
    const table = this.tables.length > 0 ? this.tableAt(row, col) : undefined;
    const effect = this.effects.size > 0 ? this.effects.get(row * CF_STRIDE + col) : undefined;
    const xfIndex = raw?.xf ?? this.rowXf[row] ?? this.colXf[col] ?? 0;
    const mergeEnd =
      this.mergeEnds.size > 0 ? this.mergeEnds.get(row * CF_STRIDE + col) : undefined;
    if (!raw && !table && !effect && xfIndex === 0 && !mergeEnd) return null;

    const xf = this.styles.cellXfs[xfIndex];
    let style = this.baseStyle(xfIndex, raw?.kind ?? 'blank');
    if (table || effect || raw?.color || mergeEnd) {
      style = { ...style };
      if (table) this.applyTable(style, table, row, col, xf);
      if (raw?.color) style.color = raw.color;
      if (effect) {
        const all = { font: true, fill: true, border: true };
        for (const dxf of effect.dxfs) applyDxf(style, dxf, all);
        if (effect.background) style.background = effect.background;
      }
      if (mergeEnd) {
        // The far edges of a merged area come from its last cell.
        const far = this.cellAt(mergeEnd.row, mergeEnd.col)?.style;
        if (far?.borderRight) style.borderRight = far.borderRight;
        if (far?.borderBottom) style.borderBottom = far.borderBottom;
      }
    }

    const cell: GridCell = { text: raw?.text ?? '', style };
    if (raw?.runs) cell.runs = raw.runs;
    if (raw?.fill) cell.fill = raw.fill;
    if (raw?.kind === 'text' && !style.wrap) cell.overflow = true;
    if (effect?.icon) cell.icon = effect.icon;
    if (effect?.bar) cell.bar = effect.bar;
    if (effect?.hideText) cell.hideText = true;
    if (table?.filter && row === table.top) cell.filter = true;
    return cell;
  };

  build(layout: {
    defaultRowHeight: number;
    defaultColWidth: number;
    showGridLines: boolean;
  }): GridSource {
    // A little room past the content, so overflowing text and the last borders stay visible.
    const colCount = Math.min(Math.max(this.usedCols + 2, 8), MAX_SHEET_COLS);
    const rowCount = Math.min(Math.max(this.usedRows + 3, 12), MAX_SHEET_ROWS);
    const font = this.styles.defaultFont;
    const source: GridSource = {
      rowCount,
      colCount,
      colWidths: this.colWidths,
      rowHeights: this.rowHeights,
      defaultColWidth: layout.defaultColWidth,
      defaultRowHeight: layout.defaultRowHeight,
      merges: this.merges,
      sheet: true,
      showGridLines: layout.showGridLines,
      font: { size: pointsToPx(font.size ?? 11), ...(font.name ? { family: font.name } : {}) },
      cell: this.cellAt,
    };
    // Only rows cut off by the row limit count as truncated; trailing empty rows do not.
    if (this.totalRows > MAX_SHEET_ROWS) source.totalRows = this.totalRows;
    if (this.overlays.length > 0) source.overlays = this.overlays;
    return source;
  }
}

interface Anchor {
  row: number;
  col: number;
  x: number;
  y: number;
}

function parseAnchorPoint(element: XmlElement | null): Anchor | null {
  if (!element) return null;
  const value = (name: string): number => Number(textOf(child(element, name))) || 0;
  const row = value('row');
  const col = value('col');
  if (row < 0 || col < 0 || row >= MAX_SHEET_ROWS || col >= MAX_SHEET_COLS) return null;
  return { row, col, x: emuToPx(value('colOff')), y: emuToPx(value('rowOff')) };
}

/** Reads the pictures and charts of a drawing part and adds them to the sheet. */
async function loadDrawing(
  zip: ZipArchive,
  path: string,
  parser: SheetParser,
  styles: StyleSheet,
  host: WorkbookHost | undefined,
): Promise<void> {
  const drawing = await readXmlPart(zip, path);
  if (!drawing) return;
  const relationships = await readRelationships(zip, path);
  for (const anchor of childElements(drawing)) {
    let from: Anchor | null;
    let to: Anchor | null = null;
    let size: { width: number; height: number } | null = null;
    const extent = child(anchor, 'ext');
    if (anchor.local === 'twoCellAnchor') {
      from = parseAnchorPoint(child(anchor, 'from'));
      to = parseAnchorPoint(child(anchor, 'to'));
      if (!to) continue;
    } else if (anchor.local === 'oneCellAnchor') {
      from = parseAnchorPoint(child(anchor, 'from'));
      size = {
        width: emuToPx(Number(attr(extent, 'cx')) || 0),
        height: emuToPx(Number(attr(extent, 'cy')) || 0),
      };
    } else if (anchor.local === 'absoluteAnchor') {
      const position = child(anchor, 'pos');
      from = {
        row: 0,
        col: 0,
        x: emuToPx(Number(attr(position, 'x')) || 0),
        y: emuToPx(Number(attr(position, 'y')) || 0),
      };
      size = {
        width: emuToPx(Number(attr(extent, 'cx')) || 0),
        height: emuToPx(Number(attr(extent, 'cy')) || 0),
      };
    } else continue;
    if (!from) continue;

    let create: GridOverlay['create'] | null = null;
    const picture = child(anchor, 'pic');
    const chartReference = descendants(child(anchor, 'graphicFrame'), 'chart')[0];
    if (picture && host) {
      const relationship = relationships.get(
        attr(descendants(picture, 'blip')[0], 'r:embed') ?? '',
      );
      const mimeType =
        relationship && !relationship.external ? imageMimeType(relationship.target) : null;
      if (!relationship || !mimeType) continue;
      const data = await zip.read(relationship.target);
      if (!data) continue;
      const url = host.objectUrl(new Blob([toBlobPart(data)], { type: mimeType }));
      const description = attr(find(picture, 'nvPicPr', 'cNvPr'), 'descr') ?? '';
      create = () => h('img', { attrs: { src: url, alt: description, draggable: 'false' } });
    } else if (chartReference) {
      const relationship: Relationship | undefined = relationships.get(
        attr(chartReference, 'r:id') ?? '',
      );
      if (!relationship || relationship.external) continue;
      const model = parseChart(await readXmlPart(zip, relationship.target), styles.theme);
      if (!model) continue;
      create = (width, height) => renderChart(model, width, height);
    }
    if (!create) continue;
    const overlay: GridOverlay = { from, create };
    if (to) overlay.to = to;
    else if (size) {
      overlay.width = size.width;
      overlay.height = size.height;
    }
    parser.addOverlay(overlay, to ?? from);
  }
}

/** Opens an XLSX workbook. Sheets are parsed lazily, one at a time. */
export async function openWorkbook(zip: ZipArchive, host?: WorkbookHost): Promise<WorkbookSheet[]> {
  const workbookPath = 'xl/workbook.xml';
  const workbook = await requireXmlPart(zip, workbookPath, 'Excel workbook');
  const relationships = await readRelationships(zip, workbookPath);
  const date1904 = isTrue(attr(child(workbook, 'workbookPr'), 'date1904'));
  const partOfType = (type: string, fallback: string): string =>
    [...relationships.values()].find((rel) => rel.type === type && !rel.external)?.target ??
    fallback;

  // Styles and shared strings are needed by every sheet, so they are parsed once and shared.
  let shared: Promise<{ strings: SharedString[]; styles: StyleSheet }> | null = null;
  const loadShared = (): Promise<{ strings: SharedString[]; styles: StyleSheet }> =>
    (shared ??= (async () => {
      const [stringsXml, stylesXml, theme] = await Promise.all([
        zip.readText(partOfType('sharedStrings', 'xl/sharedStrings.xml')),
        zip.readText(partOfType('styles', 'xl/styles.xml')),
        loadTheme(zip, partOfType('theme', 'xl/theme/theme1.xml')),
      ]);
      const styles = parseStyles(stylesXml ? parseXml(stylesXml) : null, theme);
      const strings = stringsXml
        ? childElements(parseXml(stringsXml), 'si').map((item) => parseStringItem(item, styles))
        : [];
      return { strings, styles };
    })());

  const sheets: WorkbookSheet[] = [];
  for (const sheet of childElements(find(workbook, 'sheets'), 'sheet')) {
    const state = attr(sheet, 'state');
    if (state === 'hidden' || state === 'veryHidden') continue;
    const name = attr(sheet, 'name') ?? `Sheet${sheets.length + 1}`;
    const relationship = relationships.get(attr(sheet, 'r:id') ?? '');
    const path = relationship?.target;
    // Chart sheets and macro sheets have no cell grid to show.
    if (!path || !/worksheets\//.test(path)) continue;
    sheets.push({
      name,
      async load(): Promise<GridSource> {
        const [{ strings, styles }, xml] = await Promise.all([loadShared(), zip.readText(path)]);
        if (xml === null) {
          throw new AnyViewError('PARSE_FAILED', `Sheet "${name}" is missing from the workbook.`);
        }
        const root = parseXml(xml);
        const parser = new SheetParser(root, strings, styles, date1904);
        parser.parseCells();
        const layout = parser.parseLayout();
        const sheetRelationships = await readRelationships(zip, path);
        for (const part of sheetRelationships.values()) {
          if (part.external) continue;
          // Tables, pictures and charts are extras: a damaged one must not hide the cells.
          try {
            if (part.type === 'table') parser.addTable(await readXmlPart(zip, part.target));
            else if (part.type === 'drawing') {
              await loadDrawing(zip, part.target, parser, styles, host);
            }
          } catch {
            // Skip the broken part.
          }
        }
        parser.applyConditionalFormats();
        return parser.build(layout);
      },
    });
  }
  if (sheets.length === 0) {
    throw new AnyViewError(
      'PARSE_FAILED',
      'This workbook does not contain any visible worksheets.',
    );
  }
  return sheets;
}
