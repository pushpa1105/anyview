import { clear, h, observeResize, rafThrottle, setStyles } from './dom';
import { svg } from './svg';

export interface GridBorder {
  /** Line width in CSS pixels. */
  width: number;
  style: 'solid' | 'dashed' | 'dotted' | 'double';
  /** CSS colour. */
  color: string;
}

export interface GridCellStyle {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  /** CSS colour, e.g. `#1f2328`. */
  color?: string;
  /** CSS colour or gradient. */
  background?: string;
  align?: 'left' | 'center' | 'right';
  valign?: 'top' | 'middle' | 'bottom';
  wrap?: boolean;
  /** Extra padding on the aligned side, in CSS pixels at 100% zoom. */
  indent?: number;
  /** A plain font family name. */
  fontFamily?: string;
  /** Font size in CSS pixels at 100% zoom. */
  fontSize?: number;
  borderTop?: GridBorder;
  borderRight?: GridBorder;
  borderBottom?: GridBorder;
  borderLeft?: GridBorder;
}

/** A differently formatted part of a cell's text. */
export interface GridRun {
  text: string;
  style?: GridCellStyle;
}

export type GridIconShape =
  | 'check'
  | 'cross'
  | 'exclamation'
  | 'up'
  | 'upRight'
  | 'right'
  | 'downRight'
  | 'down'
  | 'circle'
  | 'triangleUp'
  | 'triangleDown'
  | 'dash'
  | 'flag'
  | 'star';

export interface GridIcon {
  shape: GridIconShape;
  color: string;
  /** Draws the glyph in white on a filled circle. */
  circled?: boolean;
}

export interface GridCell {
  text: string;
  style?: GridCellStyle;
  /** Rich text; when present it is shown instead of `text`. */
  runs?: readonly GridRun[];
  /** Accounting-style layout: `left` hugs the left edge and `right` the right edge. */
  fill?: { left: string; right: string };
  icon?: GridIcon;
  /** Keeps the value out of view (icon-only and bar-only cells). */
  hideText?: boolean;
  /** Data bar behind the text; `start` and `end` are fractions of the cell width. */
  bar?: { color: string; start: number; end: number };
  /** Shows a filter button, as on table header cells. */
  filter?: boolean;
  /** Lets text that does not fit run over empty neighbouring cells, like a spreadsheet. */
  overflow?: boolean;
}

export interface GridMerge {
  row: number;
  col: number;
  rowSpan: number;
  colSpan: number;
}

/** A position inside the grid: a cell plus a pixel offset into it (at 100% zoom). */
export interface GridAnchor {
  row: number;
  col: number;
  x?: number;
  y?: number;
}

/** Something drawn above the cells, such as a picture or a chart. */
export interface GridOverlay {
  from: GridAnchor;
  /** Bottom-right corner; when missing, `width` and `height` are used. */
  to?: GridAnchor;
  width?: number;
  height?: number;
  /** Builds the content for a box of the given size (CSS pixels at 100% zoom). */
  create(width: number, height: number): Element;
}

/** Tabular data for the grid. Rows and columns are zero-based. */
export interface GridSource {
  rowCount: number;
  colCount: number;
  cell(row: number, col: number): GridCell | null;
  /** Column widths in CSS pixels at 100% zoom; missing entries use the default. */
  colWidths?: readonly (number | undefined)[];
  /** Row heights in CSS pixels at 100% zoom; missing entries use the default. */
  rowHeights?: readonly (number | undefined)[];
  defaultColWidth?: number;
  defaultRowHeight?: number;
  merges?: readonly GridMerge[];
  /** Row count before truncation, when the data was cut to stay responsive. */
  totalRows?: number;
  /**
   * Spreadsheet look: cells are drawn on white with their own fonts and colours
   * whatever the viewer theme, and text is clipped rather than ellipsized.
   */
  sheet?: boolean;
  /** Set to `false` to hide the default lines between cells. */
  showGridLines?: boolean;
  /** Default font of the cells. */
  font?: { family?: string; size?: number };
  overlays?: readonly GridOverlay[];
}

export interface Grid {
  readonly element: HTMLElement;
  getZoom(): number;
  setZoom(zoom: number): void;
  /** Builds a plain `<table>` of the first `maxRows` rows, for printing. */
  toTable(maxRows?: number): HTMLTableElement;
  destroy(): void;
}

const ROW_HEIGHT = 24;
const DEFAULT_COL_WIDTH = 100;
const OVERSCAN = 2;
const FALLBACK_VIEWPORT = { width: 1200, height: 700 };
/** Beyond this many covered cells, merges are ignored rather than indexed. */
const MAX_MERGED_CELLS = 250_000;
/** How many neighbouring columns overflowing text may cover, and how far back it is looked for. */
const MAX_SPILL_COLS = 40;
const MAX_COL_WIDTH = 2000;
const MAX_ROW_HEIGHT = 1000;

/** `0` → `A`, `25` → `Z`, `26` → `AA`. */
export function columnName(index: number): string {
  let name = '';
  for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) {
    name = String.fromCharCode(65 + ((n - 1) % 26)) + name;
  }
  return name;
}

const fontStack = (family: string): string =>
  `"${family.replace(/["\\]/g, '')}", Calibri, Carlito, "Segoe UI", Arial, sans-serif`;

const borderCss = (border: GridBorder | undefined): string | undefined =>
  border ? `${border.width}px ${border.style} ${border.color}` : undefined;

function cellStyles(
  style: GridCellStyle | undefined,
  zoom = 1,
): Record<string, string | undefined> {
  if (!style) return {};
  const decoration = [style.underline && 'underline', style.strike && 'line-through']
    .filter(Boolean)
    .join(' ');
  return {
    'font-weight': style.bold ? '600' : undefined,
    'font-style': style.italic ? 'italic' : undefined,
    'text-decoration': decoration || undefined,
    color: style.color,
    background: style.background,
    'text-align': style.align,
    'font-family': style.fontFamily ? fontStack(style.fontFamily) : undefined,
    'font-size': style.fontSize ? `${style.fontSize * zoom}px` : undefined,
  };
}

const ICON_ROTATION: Partial<Record<GridIconShape, number>> = {
  upRight: 45,
  right: 90,
  downRight: 135,
  down: 180,
  triangleDown: 180,
};

/** Builds a 16×16 icon for conditional-formatting icon sets. */
export function createGridIcon(icon: GridIcon, size: number): SVGElement {
  const stroke = icon.circled ? '#ffffff' : icon.color;
  const line = {
    fill: 'none',
    stroke,
    'stroke-width': icon.circled ? 2 : 2.6,
    'stroke-linecap': 'round',
    'stroke-linejoin': 'round',
  };
  const rotation = ICON_ROTATION[icon.shape];
  const transform = rotation ? `rotate(${rotation} 8 8)` : undefined;
  let glyph: SVGElement;
  switch (icon.shape) {
    case 'check':
      glyph = svg('path', {
        ...line,
        d: icon.circled ? 'M4.6 8.4l2.3 2.3 4.5-4.9' : 'M3 8.6l3.3 3.3L13 4.6',
      });
      break;
    case 'cross':
      glyph = svg('path', {
        ...line,
        d: icon.circled ? 'M5.4 5.4l5.2 5.2M10.6 5.4l-5.2 5.2' : 'M4 4l8 8M12 4l-8 8',
      });
      break;
    case 'exclamation':
      glyph = svg('path', { ...line, d: 'M8 4v5M8 11.8v.4' });
      break;
    case 'up':
    case 'upRight':
    case 'right':
    case 'downRight':
    case 'down':
      glyph = svg('path', { ...line, transform, d: 'M8 13V3.5M4 7.2l4-4 4 4' });
      break;
    case 'triangleUp':
    case 'triangleDown':
      glyph = svg('path', { fill: icon.color, transform, d: 'M8 3.5l5.5 9h-11z' });
      break;
    case 'dash':
      glyph = svg('rect', { fill: icon.color, x: 3, y: 6.5, width: 10, height: 3, rx: 1 });
      break;
    case 'flag':
      glyph = svg('path', {
        fill: icon.color,
        stroke: icon.color,
        'stroke-width': 1.4,
        'stroke-linejoin': 'round',
        d: 'M4 14V2.5M4 3h8.5l-2.4 3 2.4 3H4z',
      });
      break;
    case 'star':
      glyph = svg('path', {
        fill: icon.color,
        d: 'M8 1.8l1.9 4 4.3.5-3.2 3 .9 4.3L8 11.4l-3.9 2.2.9-4.3-3.2-3 4.3-.5z',
      });
      break;
    default:
      glyph = svg('circle', { fill: icon.color, cx: 8, cy: 8, r: 6 });
  }
  return svg(
    'svg',
    {
      class: 'av-grid-icon',
      viewBox: '0 0 16 16',
      width: size,
      height: size,
      'aria-hidden': 'true',
    },
    icon.circled && svg('circle', { fill: icon.color, cx: 8, cy: 8, r: 7 }),
    glyph,
  );
}

/** Prefix sums of sizes: `offsets[i]` is where item `i` starts. */
function prefixSums(
  count: number,
  sizes: readonly (number | undefined)[] | undefined,
  fallback: number,
  max: number,
): Float64Array {
  const offsets = new Float64Array(count + 1);
  for (let index = 0; index < count; index++) {
    const size = sizes?.[index];
    offsets[index + 1] =
      offsets[index]! + (size !== undefined && size >= 0 ? Math.min(size, max) : fallback);
  }
  return offsets;
}

/** Index of the item that contains `position`, by binary search over prefix sums. */
function indexAt(offsets: Float64Array, count: number, position: number): number {
  let low = 0;
  let high = count - 1;
  while (low < high) {
    const mid = (low + high) >> 1;
    if (offsets[mid + 1]! <= position) low = mid + 1;
    else high = mid;
  }
  return low;
}

/**
 * A spreadsheet-style grid that only keeps the visible cells in the DOM, so a
 * 100 000-row sheet scrolls as smoothly as a 10-row one.
 */
export function createGrid(parent: HTMLElement, source: GridSource): Grid {
  const { rowCount, colCount } = source;
  const sheet = source.sheet === true;
  const gridLines = source.showGridLines !== false;
  let zoom = 1;

  // Positions at 100% zoom.
  const offsets = prefixSums(
    colCount,
    source.colWidths,
    source.defaultColWidth ?? DEFAULT_COL_WIDTH,
    sheet ? MAX_COL_WIDTH : 800,
  );
  const tops = prefixSums(
    rowCount,
    source.rowHeights,
    source.defaultRowHeight ?? ROW_HEIGHT,
    MAX_ROW_HEIGHT,
  );

  // Index merges: anchor cells know their span, covered cells point at their anchor.
  const anchors = new Map<number, GridMerge>();
  const covered = new Map<number, GridMerge>();
  let mergedCells = 0;
  for (const merge of source.merges ?? []) {
    const rows = Math.min(merge.rowSpan, rowCount - merge.row);
    const cols = Math.min(merge.colSpan, colCount - merge.col);
    if (rows < 1 || cols < 1 || rows * cols < 2) continue;
    mergedCells += rows * cols;
    if (mergedCells > MAX_MERGED_CELLS) break;
    const clipped: GridMerge = { row: merge.row, col: merge.col, rowSpan: rows, colSpan: cols };
    anchors.set(merge.row * colCount + merge.col, clipped);
    for (let r = merge.row; r < merge.row + rows; r++) {
      for (let c = merge.col; c < merge.col + cols; c++) {
        if (r !== merge.row || c !== merge.col) covered.set(r * colCount + c, clipped);
      }
    }
  }
  const isMerged = (row: number, col: number): boolean => {
    const key = row * colCount + col;
    return anchors.has(key) || covered.has(key);
  };

  const rowHeadWidth = 18 + 8 * String(Math.max(rowCount, 1)).length;
  const scroller = h('div', {
    class: sheet ? 'av-grid av-grid-sheet' : 'av-grid',
    attrs: { tabindex: '0' },
  });
  const canvas = h('div', { class: 'av-grid-canvas' });
  const colHead = h('div', { class: 'av-grid-colhead' });
  const corner = h('div', { class: 'av-grid-corner' });
  const colLabels = h('div', { style: { position: 'absolute', top: '0' } });
  const rowHead = h('div', { class: 'av-grid-rowhead' });
  const cells = h('div', {
    class: 'av-grid-cells',
    style: { 'font-family': source.font?.family ? fontStack(source.font.family) : undefined },
  });
  const overlayLayer = h('div', { class: 'av-grid-overlays' });
  colHead.append(corner, colLabels);
  canvas.append(colHead, rowHead, cells, overlayLayer);
  scroller.appendChild(canvas);
  parent.appendChild(scroller);

  // Pictures and charts are few, so they are created once and only repositioned.
  const overlays: { element: HTMLElement; box: [number, number, number, number] }[] = [];
  for (const overlay of source.overlays ?? []) {
    const point = (anchor: GridAnchor): [number, number] => {
      const col = Math.max(0, Math.min(anchor.col, colCount));
      const row = Math.max(0, Math.min(anchor.row, rowCount));
      return [offsets[col]! + (anchor.x ?? 0), tops[row]! + (anchor.y ?? 0)];
    };
    const [left, top] = point(overlay.from);
    const [right, bottom] = overlay.to
      ? point(overlay.to)
      : [left + (overlay.width ?? 0), top + (overlay.height ?? 0)];
    const width = right - left;
    const height = bottom - top;
    if (!(width > 0) || !(height > 0)) continue;
    const element = h('div', { class: 'av-grid-overlay' }, overlay.create(width, height));
    overlayLayer.appendChild(element);
    overlays.push({ element, box: [left, top, width, height] });
  }

  const headHeight = (): number => ROW_HEIGHT * zoom;

  const layout = (): void => {
    const headWidth = rowHeadWidth * zoom;
    const width = offsets[colCount]! * zoom;
    const height = tops[rowCount]! * zoom;
    scroller.style.fontSize = `${13 * zoom}px`;
    if (source.font?.size) cells.style.fontSize = `${source.font.size * zoom}px`;
    setStyles(canvas, { width: `${headWidth + width}px`, height: `${headHeight() + height}px` });
    setStyles(colHead, { height: `${headHeight()}px`, 'line-height': `${headHeight()}px` });
    setStyles(corner, { width: `${headWidth}px`, height: `${headHeight()}px` });
    setStyles(colLabels, { left: `${headWidth}px` });
    setStyles(rowHead, { width: `${headWidth}px`, height: `${height}px` });
    setStyles(cells, { top: `${headHeight()}px`, left: `${headWidth}px` });
    setStyles(overlayLayer, { top: `${headHeight()}px`, left: `${headWidth}px` });
    for (const { element, box } of overlays) {
      setStyles(element, {
        left: `${box[0] * zoom}px`,
        top: `${box[1] * zoom}px`,
        width: `${box[2] * zoom}px`,
        height: `${box[3] * zoom}px`,
      });
    }
  };

  const isBlank = (row: number, col: number): boolean => {
    if (isMerged(row, col)) return false;
    const cell = source.cell(row, col);
    return !cell || (cell.text === '' && !cell.icon && !cell.fill);
  };

  /** Width, in columns' pixels at 100% zoom, that overflowing text may use. */
  const spillWidth = (row: number, col: number, toLeft: boolean): number => {
    let width = offsets[col + 1]! - offsets[col]!;
    const step = toLeft ? -1 : 1;
    for (let c = col + step, n = 0; c >= 0 && c < colCount && n < MAX_SPILL_COLS; c += step, n++) {
      if (!isBlank(row, c)) break;
      width += offsets[c + 1]! - offsets[c]!;
    }
    return width;
  };

  const buildCell = (row: number, col: number, merge: GridMerge | undefined): HTMLElement => {
    const cell = source.cell(row, col);
    const style = cell?.style;
    const endCol = merge ? col + merge.colSpan : col + 1;
    const endRow = merge ? row + merge.rowSpan : row + 1;
    const width = (offsets[endCol]! - offsets[col]!) * zoom;
    const height = (tops[endRow]! - tops[row]!) * zoom;
    const text = cell?.text ?? '';

    // A line shared by two cells is drawn once, by the cell above or to the left.
    const above = row > 0 && style?.borderTop ? source.cell(row - 1, col)?.style : undefined;
    const before = col > 0 && style?.borderLeft ? source.cell(row, col - 1)?.style : undefined;
    const defaultLine = gridLines && !(sheet && style?.background) ? undefined : 'none';
    const element = h('div', {
      class: 'av-grid-c',
      attrs: { title: !cell?.hideText && text.length > 12 ? text : undefined },
      style: {
        top: `${tops[row]! * zoom}px`,
        left: `${offsets[col]! * zoom}px`,
        width: `${width}px`,
        height: `${height}px`,
        ...cellStyles(style, zoom),
        'border-top': above?.borderBottom ? undefined : borderCss(style?.borderTop),
        'border-left': before?.borderRight ? undefined : borderCss(style?.borderLeft),
        'border-right': borderCss(style?.borderRight) ?? defaultLine,
        'border-bottom': borderCss(style?.borderBottom) ?? defaultLine,
        'align-items':
          style?.valign === 'top'
            ? 'flex-start'
            : style?.valign === 'bottom'
              ? 'flex-end'
              : undefined,
      },
    });
    if (!cell) return element;

    if (cell.bar) {
      const start = Math.max(0, Math.min(1, cell.bar.start));
      const end = Math.max(start, Math.min(1, cell.bar.end));
      element.appendChild(
        h('span', {
          class: 'av-grid-bar',
          style: {
            left: `${start * 100}%`,
            width: `${(end - start) * 100}%`,
            background: cell.bar.color,
          },
        }),
      );
    }
    if (cell.icon) {
      element.appendChild(createGridIcon(cell.icon, Math.min(16 * zoom, Math.max(height - 2, 4))));
    }

    if (!cell.hideText && (text !== '' || cell.runs || cell.fill)) {
      const content = h('span', { class: 'av-grid-t' });
      if (cell.fill) {
        content.classList.add('av-grid-fill');
        content.append(h('span', { text: cell.fill.left }), h('span', { text: cell.fill.right }));
      } else if (cell.runs) {
        for (const run of cell.runs) {
          content.appendChild(h('span', { text: run.text, style: cellStyles(run.style, zoom) }));
        }
      } else {
        content.textContent = text;
      }
      if (style?.wrap) content.classList.add('av-grid-wrap');
      if (style?.indent) {
        const side = style.align === 'right' ? 'padding-right' : 'padding-left';
        content.style.setProperty(side, `${style.indent * zoom}px`);
      }
      if (cell.overflow && !merge && !style?.wrap && style?.align !== 'center') {
        const toLeft = style?.align === 'right';
        const available = spillWidth(row, col, toLeft) * zoom;
        if (available > width + 0.5) {
          element.classList.add('av-grid-spill');
          if (toLeft) element.style.setProperty('justify-content', 'flex-end');
          // 7px: the cell's own horizontal padding and border.
          setStyles(content, { flex: 'none', width: `${available - 7}px` });
        }
      }
      element.appendChild(content);
    }
    if (cell.filter) {
      element.appendChild(
        h(
          'span',
          { class: 'av-grid-filter', attrs: { 'aria-hidden': 'true' } },
          svg(
            'svg',
            { viewBox: '0 0 10 10', width: 8 * zoom, height: 8 * zoom },
            svg('path', {
              d: 'M2 3.5l3 3 3-3',
              fill: 'none',
              stroke: 'currentColor',
              'stroke-width': 1.4,
              'stroke-linecap': 'round',
              'stroke-linejoin': 'round',
            }),
          ),
        ),
      );
    }
    return element;
  };

  const render = (): void => {
    clear(cells);
    clear(colLabels);
    clear(rowHead);
    if (rowCount === 0 || colCount === 0) return;
    const viewWidth = scroller.clientWidth || FALLBACK_VIEWPORT.width;
    const viewHeight = scroller.clientHeight || FALLBACK_VIEWPORT.height;
    const firstRow = Math.max(0, indexAt(tops, rowCount, scroller.scrollTop / zoom) - OVERSCAN);
    const lastRow = Math.min(
      rowCount - 1,
      indexAt(tops, rowCount, (scroller.scrollTop + viewHeight) / zoom) + OVERSCAN,
    );
    const firstCol = Math.max(0, indexAt(offsets, colCount, scroller.scrollLeft / zoom) - OVERSCAN);
    const lastCol = Math.min(
      colCount - 1,
      indexAt(offsets, colCount, (scroller.scrollLeft + viewWidth) / zoom) + OVERSCAN,
    );

    for (let col = firstCol; col <= lastCol; col++) {
      const width = (offsets[col + 1]! - offsets[col]!) * zoom;
      if (width <= 0) continue;
      colLabels.appendChild(
        h('div', {
          class: 'av-grid-h',
          text: columnName(col),
          style: {
            left: `${offsets[col]! * zoom}px`,
            width: `${width}px`,
            height: `${headHeight()}px`,
          },
        }),
      );
    }

    const headWidth = rowHeadWidth * zoom;
    const fragment = document.createDocumentFragment();
    const drawnMerges = new Set<GridMerge>();

    for (let row = firstRow; row <= lastRow; row++) {
      const rowHeight = (tops[row + 1]! - tops[row]!) * zoom;
      if (rowHeight <= 0) continue;
      rowHead.appendChild(
        h('div', {
          class: 'av-grid-h',
          text: String(row + 1),
          style: {
            top: `${tops[row]! * zoom}px`,
            width: `${headWidth}px`,
            height: `${rowHeight}px`,
            'line-height': `${rowHeight - 1}px`,
          },
        }),
      );

      // Text that starts left of the visible columns may still run into view.
      if (sheet && firstCol > 0) {
        for (let col = firstCol - 1; col >= Math.max(0, firstCol - MAX_SPILL_COLS); col--) {
          if (isMerged(row, col)) break;
          const cell = source.cell(row, col);
          if (!cell || cell.text === '') continue;
          if (cell.overflow && offsets[col + 1]! > offsets[col]!) {
            fragment.appendChild(buildCell(row, col, undefined));
          }
          break;
        }
      }

      for (let col = firstCol; col <= lastCol; col++) {
        const key = row * colCount + col;
        const merge = anchors.get(key) ?? covered.get(key);
        if (merge) {
          // Draw a merged area once, from its anchor, even if the anchor is off-screen.
          if (!drawnMerges.has(merge)) {
            drawnMerges.add(merge);
            fragment.appendChild(buildCell(merge.row, merge.col, merge));
          }
        } else if (offsets[col + 1]! > offsets[col]!) {
          // Empty cells are drawn too, so the sheet keeps its grid lines.
          fragment.appendChild(buildCell(row, col, undefined));
        }
      }
    }
    cells.appendChild(fragment);
  };

  const scheduleRender = rafThrottle(render);
  scroller.addEventListener('scroll', scheduleRender, { passive: true });
  const stopObserving = observeResize(scroller, scheduleRender);
  layout();
  render();

  return {
    element: scroller,
    getZoom: () => zoom,
    setZoom(value: number): void {
      if (value === zoom || !(value > 0)) return;
      const ratio = value / zoom;
      zoom = value;
      layout();
      scroller.scrollTop *= ratio;
      scroller.scrollLeft *= ratio;
      render();
    },
    toTable(maxRows = 2000): HTMLTableElement {
      const table = h('table', {
        style: { 'border-collapse': 'collapse', 'font-size': '11px', 'font-family': 'sans-serif' },
      });
      // Trailing empty rows and columns are not worth printing.
      let lastRow = -1;
      let lastCol = -1;
      const scanRows = Math.min(rowCount, maxRows);
      for (let row = 0; row < scanRows; row++) {
        for (let col = colCount - 1; col >= 0; col--) {
          const cell = source.cell(row, col);
          if (cell && (cell.text !== '' || cell.style?.background)) {
            lastRow = row;
            if (col > lastCol) lastCol = col;
            break;
          }
        }
      }
      const printStyles = (
        style: GridCellStyle | undefined,
      ): Record<string, string | undefined> => {
        const styles = cellStyles(style);
        delete styles['font-size'];
        return styles;
      };
      for (let row = 0; row <= lastRow; row++) {
        if (tops[row + 1]! <= tops[row]!) continue;
        const tr = h('tr');
        for (let col = 0; col <= lastCol; col++) {
          const key = row * colCount + col;
          if (covered.has(key) || offsets[col + 1]! <= offsets[col]!) continue;
          const merge = anchors.get(key);
          const cell = source.cell(row, col);
          tr.appendChild(
            h('td', {
              text: cell && !cell.hideText ? cell.text : '',
              attrs: { rowspan: merge?.rowSpan, colspan: merge?.colSpan },
              style: { border: '1px solid #999', padding: '2px 4px', ...printStyles(cell?.style) },
            }),
          );
        }
        table.appendChild(tr);
      }
      return table;
    },
    destroy(): void {
      scheduleRender.cancel();
      stopObserving();
      scroller.removeEventListener('scroll', scheduleRender);
      scroller.remove();
    },
  };
}
