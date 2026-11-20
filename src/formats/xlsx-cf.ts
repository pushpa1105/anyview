import type { GridIcon } from '../ui/grid';
import { attr, child, childElements, descendants, textOf, type XmlElement } from '../util/xml';
import { compileFormula, isTruthy, type CellReader, type FormulaValue } from './xlsx-formula';
import { parseInlineDxf, type Dxf, type StyleSheet } from './xlsx-styles';

/**
 * Conditional formatting: evaluates the rules of a sheet against its cached
 * cell values and records what each affected cell should look like.
 */

export interface CfCell {
  text: string;
  /** Numeric value, for number cells. */
  value?: number;
  error?: boolean;
}

/** What conditional formatting adds to one cell. */
export interface CfEffect {
  /** Differential formats, lowest priority first (later entries win). */
  dxfs: Dxf[];
  icon?: GridIcon;
  hideText?: boolean;
  bar?: { color: string; start: number; end: number };
  background?: string;
}

export interface CfSheet {
  /** Extent of the stored cells. */
  rowCount: number;
  colCount: number;
  cell(row: number, col: number): CfCell | undefined;
}

export interface Rect {
  top: number;
  left: number;
  bottom: number;
  right: number;
}

/** Keys of the effect map: `row * CF_STRIDE + col`. */
export const CF_STRIDE = 16384;
/** Upper bound on cell visits per sheet, so pathological rules cannot stall the viewer. */
const MAX_VISITS = 600_000;

const RED = '#d9453c';
const YELLOW = '#e9b21c';
const GREEN = '#3c9d4e';
const GRAY = '#7f7f7f';
const BLACK = '#262626';

const arrows3 = (down: string, side: string, up: string): GridIcon[] => [
  { shape: 'down', color: down },
  { shape: 'right', color: side },
  { shape: 'up', color: up },
];
const arrows4 = (down: string, side: string, up: string): GridIcon[] => [
  { shape: 'down', color: down },
  { shape: 'downRight', color: side },
  { shape: 'upRight', color: side },
  { shape: 'up', color: up },
];
const arrows5 = (down: string, side: string, up: string): GridIcon[] => [
  { shape: 'down', color: down },
  { shape: 'downRight', color: side },
  { shape: 'right', color: side },
  { shape: 'upRight', color: side },
  { shape: 'up', color: up },
];
const circles = (...colors: string[]): GridIcon[] =>
  colors.map((color) => ({ shape: 'circle', color }));

/** Icons of each set, from the lowest value to the highest. */
const ICON_SETS: Readonly<Record<string, GridIcon[]>> = {
  '3Arrows': arrows3(RED, YELLOW, GREEN),
  '3ArrowsGray': arrows3(GRAY, GRAY, GRAY),
  '3Flags': [RED, YELLOW, GREEN].map((color) => ({ shape: 'flag', color })),
  '3TrafficLights1': circles(RED, YELLOW, GREEN),
  '3TrafficLights2': circles(RED, YELLOW, GREEN),
  '3Signs': circles(RED, YELLOW, GREEN),
  '3Symbols': [
    { shape: 'cross', color: RED, circled: true },
    { shape: 'exclamation', color: YELLOW, circled: true },
    { shape: 'check', color: GREEN, circled: true },
  ],
  '3Symbols2': [
    { shape: 'cross', color: RED },
    { shape: 'exclamation', color: YELLOW },
    { shape: 'check', color: GREEN },
  ],
  '3Stars': [
    { shape: 'star', color: '#d9d9d9' },
    { shape: 'star', color: '#f0d58a' },
    { shape: 'star', color: YELLOW },
  ],
  '3Triangles': [
    { shape: 'triangleDown', color: RED },
    { shape: 'dash', color: YELLOW },
    { shape: 'triangleUp', color: GREEN },
  ],
  '4Arrows': arrows4(RED, YELLOW, GREEN),
  '4ArrowsGray': arrows4(GRAY, GRAY, GRAY),
  '4RedToBlack': circles(BLACK, GRAY, '#e8a5a0', RED),
  '4Rating': circles('#d9d9d9', '#a6a6a6', '#737373', BLACK),
  '4TrafficLights': circles(BLACK, RED, YELLOW, GREEN),
  '5Arrows': arrows5(RED, YELLOW, GREEN),
  '5ArrowsGray': arrows5(GRAY, GRAY, GRAY),
  '5Rating': circles('#e6e6e6', '#bfbfbf', '#999999', '#666666', BLACK),
  '5Quarters': circles('#e6e6e6', '#bfbfbf', '#999999', '#666666', BLACK),
  '5Boxes': circles('#e6e6e6', '#bfbfbf', '#999999', '#666666', BLACK),
};

function columnNumber(letters: string): number {
  let col = 0;
  for (const letter of letters.toUpperCase()) col = col * 26 + (letter.charCodeAt(0) - 64);
  return col - 1;
}

/** Parses `"B2:D9"`, `"C4"`, `"A:C"` or `"3:8"`; the result is clipped to the given extent. */
export function parseRange(reference: string, rowCount: number, colCount: number): Rect | null {
  const match = /^\$?([A-Za-z]{0,3})\$?(\d{0,7})(?::\$?([A-Za-z]{0,3})\$?(\d{0,7}))?$/.exec(
    reference.trim(),
  );
  if (!match || (!match[1] && !match[2])) return null;
  const [, col1, row1, col2 = col1, row2 = row1] = match;
  const single = match[3] === undefined && match[4] === undefined;
  const left = col1 ? columnNumber(col1) : 0;
  const top = row1 ? Number(row1) - 1 : 0;
  const right = single ? (col1 ? left : colCount - 1) : col2 ? columnNumber(col2) : colCount - 1;
  const bottom = single ? (row1 ? top : rowCount - 1) : row2 ? Number(row2) - 1 : rowCount - 1;
  const rect = {
    top: Math.max(0, Math.min(top, bottom)),
    left: Math.max(0, Math.min(left, right)),
    bottom: Math.min(Math.max(top, bottom), rowCount - 1),
    right: Math.min(Math.max(left, right), colCount - 1),
  };
  return rect.top > rect.bottom || rect.left > rect.right ? null : rect;
}

export function parseRanges(sqref: string, rowCount: number, colCount: number): Rect[] {
  return sqref
    .split(/\s+/)
    .filter(Boolean)
    .map((reference) => parseRange(reference, rowCount, colCount))
    .filter((rect): rect is Rect => rect !== null);
}

interface Rule {
  ranges: Rect[];
  priority: number;
  element: XmlElement;
  /** Rules from the `x14` extension list keep their values in child elements. */
  extended: boolean;
}

interface Threshold {
  type: string;
  value: number;
  inclusive: boolean;
}

function quantile(sorted: number[], fraction: number): number {
  if (sorted.length === 0) return 0;
  const position = (sorted.length - 1) * Math.max(0, Math.min(1, fraction));
  const low = Math.floor(position);
  const high = Math.ceil(position);
  return sorted[low]! + (sorted[high]! - sorted[low]!) * (position - low);
}

function interpolate(from: string, to: string, ratio: number): string {
  const part = (hex: string, at: number): number => parseInt(hex.slice(at, at + 2), 16);
  return `#${[1, 3, 5]
    .map((at) =>
      Math.round(part(from, at) + (part(to, at) - part(from, at)) * ratio)
        .toString(16)
        .padStart(2, '0'),
    )
    .join('')}`;
}

export function evaluateConditionalFormats(
  root: XmlElement,
  styles: StyleSheet,
  sheet: CfSheet,
): Map<number, CfEffect> {
  const effects = new Map<number, CfEffect>();
  const rules: Rule[] = [];
  for (const block of childElements(root, 'conditionalFormatting')) {
    const ranges = parseRanges(attr(block, 'sqref') ?? '', sheet.rowCount, sheet.colCount);
    for (const element of childElements(block, 'cfRule')) {
      rules.push({
        ranges,
        priority: Number(attr(element, 'priority') ?? 0) || 0,
        element,
        extended: false,
      });
    }
  }
  for (const block of descendants(child(root, 'extLst'), 'conditionalFormatting')) {
    const ranges = parseRanges(textOf(child(block, 'sqref')), sheet.rowCount, sheet.colCount);
    for (const element of childElements(block, 'cfRule')) {
      // Extended data bars repeat a classic rule; only icon sets are unique to this list.
      if (attr(element, 'type') !== 'iconSet') continue;
      rules.push({
        ranges,
        priority: Number(attr(element, 'priority') ?? 0) || 0,
        element,
        extended: true,
      });
    }
  }
  if (rules.length === 0) return effects;
  // Low priority numbers win, so they are applied last.
  rules.sort((a, b) => b.priority - a.priority);

  let visits = 0;
  const effectAt = (row: number, col: number): CfEffect => {
    const key = row * CF_STRIDE + col;
    let effect = effects.get(key);
    if (!effect) effects.set(key, (effect = { dxfs: [] }));
    return effect;
  };
  const each = (
    ranges: Rect[],
    visit: (row: number, col: number, cell: CfCell | undefined) => void,
  ): void => {
    for (const range of ranges) {
      for (let row = range.top; row <= range.bottom; row++) {
        for (let col = range.left; col <= range.right; col++) {
          if (++visits > MAX_VISITS) return;
          visit(row, col, sheet.cell(row, col));
        }
      }
    }
  };
  const numbersIn = (ranges: Rect[]): number[] => {
    const values: number[] = [];
    each(ranges, (_row, _col, cell) => {
      if (cell?.value !== undefined) values.push(cell.value);
    });
    return values.sort((a, b) => a - b);
  };
  const read: CellReader = (row, col): FormulaValue => {
    const cell = sheet.cell(row, col);
    if (!cell) return null;
    if (cell.value !== undefined) return cell.value;
    return cell.text === '' ? null : cell.text;
  };

  const thresholds = (container: XmlElement, rule: Rule, sorted: number[]): Threshold[] =>
    childElements(container, 'cfvo').map((cfvo) => {
      const type = attr(cfvo, 'type') ?? 'num';
      const raw = rule.extended ? textOf(child(cfvo, 'f')) : (attr(cfvo, 'val') ?? '');
      const amount = Number(raw);
      const min = sorted[0] ?? 0;
      const max = sorted[sorted.length - 1] ?? 0;
      let value: number;
      if (type === 'min' || type === 'autoMin') value = type === 'autoMin' ? Math.min(min, 0) : min;
      else if (type === 'max' || type === 'autoMax') value = max;
      else if (type === 'percent') value = min + ((max - min) * amount) / 100;
      else if (type === 'percentile') value = quantile(sorted, amount / 100);
      else value = raw === '' ? NaN : amount;
      return { type, value, inclusive: !['0', 'false'].includes(attr(cfvo, 'gte') ?? '1') };
    });

  for (const rule of rules) {
    const { element, ranges } = rule;
    if (ranges.length === 0) continue;
    const type = attr(element, 'type') ?? '';
    const dxfIndex = attr(element, 'dxfId');
    const inlineDxf = child(element, 'dxf');
    const dxf: Dxf | undefined =
      dxfIndex !== undefined
        ? styles.dxfs[Number(dxfIndex)]
        : inlineDxf
          ? parseInlineDxf(inlineDxf, styles)
          : undefined;
    const base = ranges[0]!;
    const formulas = childElements(element, 'formula').map((formula) =>
      compileFormula(textOf(formula)),
    );
    const operand = (index: number, row: number, col: number): FormulaValue | undefined => {
      const formula = formulas[index];
      return formula ? formula(read, row - base.top, col - base.left, row, col) : undefined;
    };
    const mark = (test: (cell: CfCell | undefined, row: number, col: number) => boolean): void => {
      if (!dxf) return;
      each(ranges, (row, col, cell) => {
        if (test(cell, row, col)) effectAt(row, col).dxfs.push(dxf);
      });
    };
    const textOfRule = (attr(element, 'text') ?? '').toLowerCase();
    const cellText = (cell: CfCell | undefined): string => (cell?.text ?? '').toLowerCase();

    switch (type) {
      case 'expression':
        if (formulas[0]) mark((_cell, row, col) => isTruthy(operand(0, row, col) ?? null));
        break;
      case 'cellIs': {
        const operator = attr(element, 'operator') ?? 'equal';
        if (!formulas[0]) break;
        mark((cell, row, col) => {
          const value: FormulaValue =
            cell?.value !== undefined ? cell.value : cell && cell.text !== '' ? cell.text : null;
          const first = operand(0, row, col);
          if (first === undefined || first === null || value === null) return false;
          const order = (other: FormulaValue): number | null => {
            if (typeof value === 'number' && typeof other === 'number') return value - other;
            if (typeof value === 'string' && typeof other === 'string') {
              const [a, b] = [value.toLowerCase(), other.toLowerCase()];
              return a < b ? -1 : a > b ? 1 : 0;
            }
            return null;
          };
          const a = order(first);
          if (operator === 'between' || operator === 'notBetween') {
            const second = operand(1, row, col);
            const b = second === undefined || second === null ? null : order(second);
            if (a === null || b === null) return false;
            const inside = (a >= 0 && b <= 0) || (a <= 0 && b >= 0);
            return operator === 'between' ? inside : !inside;
          }
          if (a === null) return operator === 'notEqual';
          switch (operator) {
            case 'equal':
              return a === 0;
            case 'notEqual':
              return a !== 0;
            case 'greaterThan':
              return a > 0;
            case 'greaterThanOrEqual':
              return a >= 0;
            case 'lessThan':
              return a < 0;
            case 'lessThanOrEqual':
              return a <= 0;
            default:
              return false;
          }
        });
        break;
      }
      case 'containsText':
        mark((cell) => textOfRule !== '' && cellText(cell).includes(textOfRule));
        break;
      case 'notContainsText':
        mark((cell) => !cellText(cell).includes(textOfRule));
        break;
      case 'beginsWith':
        mark((cell) => textOfRule !== '' && cellText(cell).startsWith(textOfRule));
        break;
      case 'endsWith':
        mark((cell) => textOfRule !== '' && cellText(cell).endsWith(textOfRule));
        break;
      case 'containsBlanks':
        mark((cell) => cellText(cell).trim() === '');
        break;
      case 'notContainsBlanks':
        mark((cell) => cellText(cell).trim() !== '');
        break;
      case 'containsErrors':
        mark((cell) => cell?.error === true);
        break;
      case 'notContainsErrors':
        mark((cell) => cell?.error !== true);
        break;
      case 'duplicateValues':
      case 'uniqueValues': {
        const counts = new Map<string, number>();
        each(ranges, (_row, _col, cell) => {
          const text = cellText(cell);
          if (text !== '') counts.set(text, (counts.get(text) ?? 0) + 1);
        });
        mark((cell) => {
          const count = counts.get(cellText(cell)) ?? 0;
          return type === 'duplicateValues' ? count > 1 : count === 1;
        });
        break;
      }
      case 'top10': {
        const sorted = numbersIn(ranges);
        if (sorted.length === 0) break;
        const rank = Math.max(1, Number(attr(element, 'rank') ?? 10) || 10);
        const count = ['1', 'true'].includes(attr(element, 'percent') ?? '')
          ? Math.max(1, Math.floor((sorted.length * rank) / 100))
          : Math.min(rank, sorted.length);
        const bottom = ['1', 'true'].includes(attr(element, 'bottom') ?? '');
        const limit = bottom ? sorted[count - 1]! : sorted[sorted.length - count]!;
        mark((cell) =>
          cell?.value === undefined ? false : bottom ? cell.value <= limit : cell.value >= limit,
        );
        break;
      }
      case 'aboveAverage': {
        const values = numbersIn(ranges);
        if (values.length === 0) break;
        const average = values.reduce((sum, value) => sum + value, 0) / values.length;
        const above = !['0', 'false'].includes(attr(element, 'aboveAverage') ?? '1');
        const equal = ['1', 'true'].includes(attr(element, 'equalAverage') ?? '');
        mark((cell) => {
          if (cell?.value === undefined) return false;
          if (equal && cell.value === average) return true;
          return above ? cell.value > average : cell.value < average;
        });
        break;
      }
      case 'colorScale': {
        const scale = child(element, 'colorScale');
        if (!scale) break;
        const sorted = numbersIn(ranges);
        const stops = thresholds(scale, rule, sorted);
        const colors = childElements(scale, 'color').map((color) => styles.color(color));
        if (stops.length < 2 || colors.length < stops.length || colors.includes(undefined)) break;
        each(ranges, (row, col, cell) => {
          if (cell?.value === undefined) return;
          let index = 0;
          while (index < stops.length - 2 && cell.value > stops[index + 1]!.value) index++;
          const from = stops[index]!.value;
          const to = stops[index + 1]!.value;
          const ratio = to > from ? Math.max(0, Math.min(1, (cell.value - from) / (to - from))) : 0;
          effectAt(row, col).background = interpolate(colors[index]!, colors[index + 1]!, ratio);
        });
        break;
      }
      case 'dataBar': {
        const bar = child(element, 'dataBar');
        const color = styles.color(child(bar, 'color'));
        if (!bar || !color) break;
        const sorted = numbersIn(ranges);
        const stops = thresholds(bar, rule, sorted);
        if (stops.length < 2) break;
        // An automatic minimum starts the bars at zero when all values are positive.
        const low = stops[0]!.type === 'min' ? Math.min(stops[0]!.value, 0) : stops[0]!.value;
        const high = stops[1]!.value;
        const hide = ['0', 'false'].includes(attr(bar, 'showValue') ?? '1');
        each(ranges, (row, col, cell) => {
          if (cell?.value === undefined || !(high > low)) return;
          const effect = effectAt(row, col);
          const ratio = Math.max(0, Math.min(1, (cell.value - low) / (high - low)));
          effect.bar = { color, start: 0, end: Math.max(ratio, 0.02) };
          if (hide) effect.hideText = true;
        });
        break;
      }
      case 'iconSet': {
        const set = child(element, 'iconSet');
        if (!set) break;
        const name = attr(set, 'iconSet') ?? '3TrafficLights1';
        const custom = childElements(set, 'cfIcon');
        const sorted = numbersIn(ranges);
        const stops = thresholds(set, rule, sorted);
        const icons: (GridIcon | undefined)[] =
          custom.length > 0
            ? custom.map(
                (icon) => ICON_SETS[attr(icon, 'iconSet') ?? '']?.[Number(attr(icon, 'iconId'))],
              )
            : (ICON_SETS[name] ?? []);
        if (stops.length === 0 || icons.length === 0) break;
        const reverse = ['1', 'true'].includes(attr(set, 'reverse') ?? '');
        const hide = ['0', 'false'].includes(attr(set, 'showValue') ?? '1');
        each(ranges, (row, col, cell) => {
          if (cell?.value === undefined) return;
          let index = 0;
          for (let i = 1; i < stops.length; i++) {
            const stop = stops[i]!;
            if (stop.inclusive ? cell.value >= stop.value : cell.value > stop.value) index = i;
          }
          if (reverse) index = stops.length - 1 - index;
          const effect = effectAt(row, col);
          const icon = icons[index];
          if (icon) effect.icon = icon;
          else delete effect.icon;
          if (hide) effect.hideText = true;
        });
        break;
      }
    }
  }
  return effects;
}
