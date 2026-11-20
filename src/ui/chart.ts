import type { ChartAxis, ChartGroup, ChartModel, ChartSeries, ChartText } from '../formats/chart';
import { formatNumber } from '../util/numfmt';
import { round, svg } from './svg';

/**
 * Draws a parsed chart as inline SVG. Bar, column, line, area, pie, doughnut
 * and scatter charts are supported, including stacked variants and combinations.
 */

interface Box {
  x: number;
  y: number;
  width: number;
  height: number;
}

interface Scale {
  min: number;
  max: number;
  step: number;
}

const textWidth = (text: string, size: number): number => text.length * size * 0.54;

const textAttrs = (text: ChartText): Record<string, string | number | undefined> => ({
  'font-size': round(text.size),
  fill: text.color,
  'font-weight': text.bold ? 'bold' : undefined,
  'font-style': text.italic ? 'italic' : undefined,
});

/** A "nice" axis step (1, 2, 2.5 or 5 times a power of ten) for roughly `target` intervals. */
export function niceStep(range: number, target: number): number {
  if (!(range > 0)) return 1;
  const rough = range / Math.max(target, 1);
  const power = 10 ** Math.floor(Math.log10(rough));
  const fraction = rough / power;
  const nice =
    fraction <= 1 ? 1 : fraction <= 2 ? 2 : fraction <= 2.5 ? 2.5 : fraction <= 5 ? 5 : 10;
  return nice * power;
}

/** Axis range and step for the given data extent, honouring fixed axis settings. */
export function axisScale(low: number, high: number, axis: ChartAxis, target: number): Scale {
  let min = Math.min(low, 0);
  let max = Math.max(high, 0);
  if (min === max) max = min + 1;
  const step = axis.majorUnit ?? niceStep((axis.max ?? max) - (axis.min ?? min), target);
  // Leave a little headroom like spreadsheet applications do.
  if (max > 0) max = Math.ceil((max + (max - min) * 0.02) / step) * step;
  if (min < 0) min = Math.floor((min - (max - min) * 0.02) / step) * step;
  if (axis.min !== undefined) min = axis.min;
  if (axis.max !== undefined) max = axis.max;
  if (!(max > min)) max = min + step;
  return { min, max, step };
}

function ticksOf(scale: Scale): number[] {
  const ticks: number[] = [];
  const count = Math.min(Math.round((scale.max - scale.min) / scale.step), 200);
  for (let index = 0; index <= count; index++) {
    // Rounding hides binary noise such as 0.30000000000000004.
    ticks.push(Number((scale.min + index * scale.step).toPrecision(12)));
  }
  return ticks;
}

function formatTick(value: number, code: string | undefined): string {
  if (!code || /^general$/i.test(code)) return formatNumber(value, 'General');
  return formatNumber(value, code).trim().replace(/\s+/g, ' ');
}

const colorAt = (series: ChartSeries, index: number): string =>
  series.pointColors[index] ?? series.color;

/** Colours for charts that give every point its own colour (pies). */
function variedColors(model: ChartModel, count: number): string[] {
  const base = model.groups.flatMap((group) => group.series.map((series) => series.color));
  const palette = ['#4472c4', '#ed7d31', '#a5a5a5', '#ffc000', '#5b9bd5', '#70ad47'];
  const first = base[0] ?? palette[0]!;
  const start = Math.max(0, palette.indexOf(first));
  return Array.from({ length: count }, (_unused, index) =>
    index === 0 ? first : palette[(start + index) % palette.length]!,
  );
}

function drawPie(parent: SVGElement, group: ChartGroup, box: Box, colors: string[]): void {
  const series = group.series[0]!;
  const values = series.values.map((value) => (value !== null && value > 0 ? value : 0));
  const total = values.reduce((sum, value) => sum + value, 0);
  if (!(total > 0)) return;
  const radius = Math.max(Math.min(box.width, box.height) / 2 - 4, 4);
  const inner = group.kind === 'doughnut' ? (radius * group.holeSize) / 100 : 0;
  const cx = box.x + box.width / 2;
  const cy = box.y + box.height / 2;
  const at = (angle: number, r: number): string =>
    `${round(cx + r * Math.sin(angle))} ${round(cy - r * Math.cos(angle))}`;
  let angle = 0;
  values.forEach((value, index) => {
    if (value <= 0) return;
    const sweep = (value / total) * Math.PI * 2;
    const end = angle + sweep;
    const fill = series.pointColors[index] ?? colors[index] ?? series.color;
    const stroke = series.stroke ?? { color: '#ffffff', width: 1.5 };
    const shared = { fill, stroke: stroke.color, 'stroke-width': round(stroke.width) };
    if (sweep >= Math.PI * 2 - 1e-9) {
      // A single slice is a full ring, which an arc cannot describe in one go.
      parent.appendChild(
        svg('circle', { ...shared, cx: round(cx), cy: round(cy), r: round(radius) }),
      );
      if (inner > 0) {
        parent.appendChild(
          svg('circle', { fill: '#ffffff', cx: round(cx), cy: round(cy), r: round(inner) }),
        );
      }
    } else {
      const large = sweep > Math.PI ? 1 : 0;
      const path =
        inner > 0
          ? `M${at(angle, radius)}A${round(radius)} ${round(radius)} 0 ${large} 1 ${at(end, radius)}L${at(end, inner)}A${round(inner)} ${round(inner)} 0 ${large} 0 ${at(angle, inner)}Z`
          : `M${round(cx)} ${round(cy)}L${at(angle, radius)}A${round(radius)} ${round(radius)} 0 ${large} 1 ${at(end, radius)}Z`;
      parent.appendChild(svg('path', { ...shared, d: path }));
    }
    angle = end;
  });
  if (!group.showPercent && !group.showValues) return;
  // Slice labels sit inside the slice, two thirds of the way out.
  angle = 0;
  values.forEach((value, index) => {
    if (value <= 0) return;
    const sweep = (value / total) * Math.PI * 2;
    const middle = angle + sweep / 2;
    const distance = inner > 0 ? (radius + inner) / 2 : radius * 0.68;
    const label = group.showPercent
      ? `${Math.round((value / total) * 100)}%`
      : String(series.values[index] ?? '');
    parent.appendChild(
      svg(
        'text',
        {
          x: round(cx + distance * Math.sin(middle)),
          y: round(cy - distance * Math.cos(middle)),
          'text-anchor': 'middle',
          'dominant-baseline': 'central',
          'font-size': 12,
          fill: '#000000',
        },
        label,
      ),
    );
    angle += sweep;
  });
}

/** Renders a chart into an `<svg>` of the given size (CSS pixels). */
export function renderChart(model: ChartModel, width: number, height: number): SVGElement {
  const root = svg('svg', {
    class: 'av-chart',
    viewBox: `0 0 ${round(width)} ${round(height)}`,
    preserveAspectRatio: 'none',
    role: 'img',
    'aria-label': model.title ?? 'Chart',
    'font-family': model.fontFamily
      ? `"${model.fontFamily}", Calibri, Arial, sans-serif`
      : 'Calibri, Arial, sans-serif',
  });
  if (model.background || model.border) {
    root.appendChild(
      svg('rect', {
        x: 0.5,
        y: 0.5,
        width: round(Math.max(width - 1, 0)),
        height: round(Math.max(height - 1, 0)),
        fill: model.background ?? 'none',
        stroke: model.border?.color,
        'stroke-width': model.border ? round(model.border.width) : undefined,
      }),
    );
  }

  const pad = Math.min(10, width * 0.04, height * 0.08);
  const area: Box = { x: pad, y: pad, width: width - pad * 2, height: height - pad * 2 };

  if (model.title) {
    const size = model.titleText.size;
    root.appendChild(
      svg(
        'text',
        {
          ...textAttrs(model.titleText),
          x: round(width / 2),
          y: round(area.y + size),
          'text-anchor': 'middle',
        },
        model.title,
      ),
    );
    area.y += size * 1.6;
    area.height -= size * 1.6;
  }

  const first = model.groups[0]!;
  const round_ = first.kind === 'pie' || first.kind === 'doughnut';
  const categoryCount = Math.max(
    ...model.groups.flatMap((group) => group.series.map((series) => series.values.length)),
    0,
  );
  const categories: string[] = [];
  for (let index = 0; index < categoryCount; index++) {
    const label = model.groups
      .flatMap((group) => group.series)
      .map((series) => series.categories[index])
      .find((value) => value !== undefined && value !== '');
    categories.push(label ?? String(index + 1));
  }
  const pieColors = round_ ? variedColors(model, categoryCount) : [];

  // Legend: one entry per series, or per category for pies.
  if (model.legend) {
    const entries = round_
      ? categories.map((name, index) => ({
          name,
          color: first.series[0]!.pointColors[index] ?? pieColors[index]!,
        }))
      : model.groups.flatMap((group) =>
          group.series.map((series) => ({ name: series.name, color: series.color })),
        );
    const size = model.legendText.size;
    const mark = size * 0.55;
    const entryWidth = (name: string): number => mark + 6 + textWidth(name, size) + 16;
    const legend = svg('g', { class: 'av-chart-legend', ...textAttrs(model.legendText) });
    const entry = (name: string, color: string, x: number, y: number): void => {
      legend.append(
        svg('rect', {
          x: round(x),
          y: round(y - mark),
          width: round(mark),
          height: round(mark),
          fill: color,
        }),
        svg('text', { x: round(x + mark + 6), y: round(y) }, name),
      );
    };
    if (model.legend === 'b' || model.legend === 't') {
      const total = entries.reduce((sum, item) => sum + entryWidth(item.name), 0);
      const y = model.legend === 'b' ? area.y + area.height - size * 0.3 : area.y + size;
      let x = area.x + Math.max((area.width - total + 16) / 2, 0);
      for (const item of entries) {
        entry(item.name, item.color, x, y);
        x += entryWidth(item.name);
      }
      if (model.legend === 't') area.y += size * 1.9;
      area.height -= size * 1.9;
    } else {
      const column = Math.min(
        Math.max(...entries.map((item) => entryWidth(item.name)), 0),
        area.width * 0.4,
      );
      const x = model.legend === 'l' ? area.x : area.x + area.width - column + 8;
      let y = area.y + Math.max((area.height - entries.length * size * 1.5) / 2, 0) + size;
      for (const item of entries) {
        entry(item.name, item.color, x, y);
        y += size * 1.5;
      }
      if (model.legend === 'l') area.x += column;
      area.width -= column;
    }
    root.appendChild(legend);
  }

  if (round_) {
    drawPie(root, first, area, pieColors);
    return root;
  }

  const horizontal = first.horizontal;
  const scatter = first.kind === 'scatter';
  const valueAxis = model.valueAxis;
  const categoryAxis = model.categoryAxis;

  // Value extent across all groups, taking stacking into account.
  let low = Infinity;
  let high = -Infinity;
  let percent = false;
  for (const group of model.groups) {
    const stacked = group.grouping === 'stacked' || group.grouping === 'percentStacked';
    if (group.grouping === 'percentStacked') percent = true;
    if (stacked && group.kind !== 'scatter') {
      for (let index = 0; index < categoryCount; index++) {
        let positive = 0;
        let negative = 0;
        for (const series of group.series) {
          const value = series.values[index] ?? 0;
          if (value >= 0) positive += value;
          else negative += value;
        }
        low = Math.min(low, negative);
        high = Math.max(high, positive);
      }
    } else {
      for (const series of group.series) {
        for (const value of series.values) {
          if (value === null) continue;
          low = Math.min(low, value);
          high = Math.max(high, value);
        }
      }
    }
  }
  if (!Number.isFinite(low)) [low, high] = [0, 1];
  const valueSpan = horizontal ? area.width : area.height;
  const scale = percent
    ? { min: 0, max: 1, step: 0.2 }
    : axisScale(low, high, valueAxis, Math.max(2, Math.floor(valueSpan / 45)));
  const valueCode = percent
    ? '0%'
    : (valueAxis.formatCode ?? model.groups.flatMap((group) => group.series)[0]?.formatCode);
  const ticks = ticksOf(scale);
  const tickLabels = ticks.map((tick) => formatTick(tick, valueCode));

  // Scatter x axis.
  let xScale: Scale | null = null;
  if (scatter) {
    const xs = first.series.flatMap((series) => series.xValues ?? []).filter((x) => x !== null);
    xScale = axisScale(
      Math.min(...xs, 0),
      Math.max(...xs, 1),
      categoryAxis,
      Math.max(2, Math.floor(area.width / 70)),
    );
  }
  const xTicks = xScale ? ticksOf(xScale) : [];
  const xLabels = xScale
    ? xTicks.map((tick) => formatTick(tick, categoryAxis.formatCode))
    : categories;

  // Space taken by the axis labels.
  const valueSize = valueAxis.text.size;
  const categorySize = categoryAxis.text.size;
  const valueLabelWidth = valueAxis.deleted
    ? 0
    : Math.max(...tickLabels.map((label) => textWidth(label, valueSize)), 0) + 8;
  const categoryLabelWidth = categoryAxis.deleted
    ? 0
    : Math.min(
        Math.max(...xLabels.map((label) => textWidth(label, categorySize)), 0) + 8,
        area.width * 0.35,
      );
  const plot: Box = { ...area };
  if (horizontal) {
    plot.x += categoryLabelWidth;
    plot.width -= categoryLabelWidth;
    if (!valueAxis.deleted) plot.height -= valueSize * 1.7;
    // Half a label sticks out at both ends of the value axis.
    const overhang = valueAxis.deleted
      ? 0
      : textWidth(tickLabels[tickLabels.length - 1] ?? '', valueSize) / 2;
    plot.width -= overhang;
    if (categoryLabelWidth === 0 && !valueAxis.deleted) {
      const lead = textWidth(tickLabels[0] ?? '', valueSize) / 2;
      plot.x += lead;
      plot.width -= lead;
    }
  } else {
    plot.x += valueLabelWidth;
    plot.width -= valueLabelWidth;
    if (!categoryAxis.deleted) plot.height -= categorySize * 1.9;
    if (!valueAxis.deleted) {
      plot.y += valueSize * 0.5;
      plot.height -= valueSize * 0.5;
    }
  }
  plot.width = Math.max(plot.width, 10);
  plot.height = Math.max(plot.height, 10);

  const valueAt = (value: number): number => {
    const ratio = (value - scale.min) / (scale.max - scale.min);
    return horizontal ? plot.x + ratio * plot.width : plot.y + plot.height - ratio * plot.height;
  };
  const zero = valueAt(Math.max(scale.min, Math.min(0, scale.max)));

  if (model.plotBackground) {
    root.appendChild(
      svg('rect', {
        x: round(plot.x),
        y: round(plot.y),
        width: round(plot.width),
        height: round(plot.height),
        fill: model.plotBackground,
      }),
    );
  }

  // Gridlines, value labels and tick marks.
  const axes = svg('g', { class: 'av-chart-axes' });
  ticks.forEach((tick, index) => {
    const position = round(valueAt(tick));
    if (valueAxis.gridlines) {
      axes
        .appendChild(
          svg(
            'line',
            horizontal
              ? { x1: position, x2: position, y1: round(plot.y), y2: round(plot.y + plot.height) }
              : { x1: round(plot.x), x2: round(plot.x + plot.width), y1: position, y2: position },
          ),
        )
        .setAttribute('stroke', valueAxis.gridlines);
    }
    if (valueAxis.deleted) return;
    if (valueAxis.ticks) {
      axes
        .appendChild(
          svg(
            'line',
            horizontal
              ? {
                  x1: position,
                  x2: position,
                  y1: round(plot.y + plot.height),
                  y2: round(plot.y + plot.height + 5),
                }
              : { x1: round(plot.x - 5), x2: round(plot.x), y1: position, y2: position },
          ),
        )
        .setAttribute('stroke', valueAxis.line ?? '#000000');
    }
    axes.appendChild(
      svg(
        'text',
        horizontal
          ? {
              ...textAttrs(valueAxis.text),
              x: position,
              y: round(plot.y + plot.height + valueSize * 1.35),
              'text-anchor': 'middle',
            }
          : {
              ...textAttrs(valueAxis.text),
              x: round(plot.x - 7),
              y: round(position + valueSize * 0.35),
              'text-anchor': 'end',
            },
        tickLabels[index],
      ),
    );
  });

  // Category (or scatter x) labels.
  const bandCount = Math.max(categoryCount, 1);
  const band = (horizontal ? plot.height : plot.width) / bandCount;
  // Horizontal bar charts list the first category at the bottom.
  const bandStart = (index: number): number =>
    horizontal ? plot.y + plot.height - (index + 1) * band : plot.x + index * band;
  if (!categoryAxis.deleted) {
    if (xScale) {
      xTicks.forEach((tick, index) => {
        const x = plot.x + ((tick - xScale.min) / (xScale.max - xScale.min)) * plot.width;
        axes.appendChild(
          svg(
            'text',
            {
              ...textAttrs(categoryAxis.text),
              x: round(x),
              y: round(plot.y + plot.height + categorySize * 1.35),
              'text-anchor': 'middle',
            },
            xLabels[index],
          ),
        );
      });
    } else {
      // Skip labels when they would overlap.
      const widest = Math.max(...categories.map((label) => textWidth(label, categorySize)), 1);
      const every = horizontal
        ? Math.max(1, Math.ceil((categorySize * 1.2) / band))
        : Math.max(1, Math.ceil((widest + 6) / band));
      categories.forEach((label, index) => {
        if (index % every !== 0) return;
        const centre = bandStart(index) + band / 2;
        axes.appendChild(
          svg(
            'text',
            horizontal
              ? {
                  ...textAttrs(categoryAxis.text),
                  x: round(plot.x - 7),
                  y: round(centre + categorySize * 0.35),
                  'text-anchor': 'end',
                }
              : {
                  ...textAttrs(categoryAxis.text),
                  x: round(centre),
                  y: round(plot.y + plot.height + categorySize * 1.35),
                  'text-anchor': 'middle',
                },
            label,
          ),
        );
      });
    }
  }
  root.appendChild(axes);

  // Series.
  const marks = svg('g', { class: 'av-chart-series' });
  const labels = svg('g', { class: 'av-chart-labels', ...textAttrs(valueAxis.text) });
  for (const group of model.groups) {
    const stacked = group.grouping === 'stacked' || group.grouping === 'percentStacked';
    const totals: number[] = [];
    if (group.grouping === 'percentStacked') {
      for (let index = 0; index < categoryCount; index++) {
        totals[index] = group.series.reduce(
          (sum, series) => sum + Math.abs(series.values[index] ?? 0),
          0,
        );
      }
    }
    const normal = (value: number, index: number): number =>
      group.grouping === 'percentStacked' ? (totals[index] ? value / totals[index] : 0) : value;

    if (group.kind === 'bar') {
      const slots = stacked ? 1 : group.series.length;
      const barSize = band / (slots + group.gapWidth / 100);
      const inset = (band - barSize * slots) / 2;
      const positive: number[] = [];
      const negative: number[] = [];
      group.series.forEach((series, order) => {
        for (let index = 0; index < categoryCount; index++) {
          const raw = series.values[index];
          if (raw === null || raw === undefined) continue;
          const value = normal(raw, index);
          let from = 0;
          if (stacked) {
            const stack = value >= 0 ? positive : negative;
            from = stack[index] ?? 0;
            stack[index] = from + value;
          }
          const a = valueAt(Math.max(scale.min, Math.min(from, scale.max)));
          const b = valueAt(Math.max(scale.min, Math.min(from + value, scale.max)));
          const offset = bandStart(index) + inset + (stacked ? 0 : order * barSize);
          // In horizontal charts the first series sits at the bottom of each cluster.
          const cross =
            horizontal && !stacked
              ? bandStart(index) + band - inset - (order + 1) * barSize
              : offset;
          const rect = horizontal
            ? { x: Math.min(a, b), y: cross, width: Math.abs(b - a), height: barSize }
            : { x: cross, y: Math.min(a, b), width: barSize, height: Math.abs(b - a) };
          marks.appendChild(
            svg('rect', {
              x: round(rect.x),
              y: round(rect.y),
              width: round(rect.width),
              height: round(rect.height),
              fill: colorAt(series, index),
              stroke: series.stroke?.color,
              'stroke-width': series.stroke ? round(series.stroke.width) : undefined,
            }),
          );
          if (group.showValues) {
            labels.appendChild(
              svg(
                'text',
                horizontal
                  ? {
                      x: round(Math.max(a, b) + 4),
                      y: round(cross + barSize / 2 + valueAxis.text.size * 0.35),
                    }
                  : {
                      x: round(cross + barSize / 2),
                      y: round(Math.min(a, b) - 4),
                      'text-anchor': 'middle',
                    },
                formatTick(raw, series.formatCode ?? valueCode),
              ),
            );
          }
        }
      });
      continue;
    }

    const stackBase: number[] = [];
    for (const series of group.series) {
      const points: [number, number][] = [];
      const bases: [number, number][] = [];
      series.values.forEach((raw, index) => {
        if (raw === null) return;
        let x: number;
        if (xScale) {
          const xValue = series.xValues?.[index];
          if (xValue === null || xValue === undefined) return;
          x = plot.x + ((xValue - xScale.min) / (xScale.max - xScale.min)) * plot.width;
        } else {
          x = plot.x + (index + 0.5) * band;
        }
        const base = stacked ? (stackBase[index] ?? 0) : 0;
        const value = base + normal(raw, index);
        if (stacked) stackBase[index] = value;
        points.push([x, valueAt(Math.max(scale.min, Math.min(value, scale.max)))]);
        bases.push([x, stacked ? valueAt(Math.max(scale.min, Math.min(base, scale.max))) : zero]);
      });
      if (points.length === 0) continue;
      const path = (list: [number, number][]): string =>
        list.map(([x, y]) => `${round(x)},${round(y)}`).join(' ');
      if (group.kind === 'area') {
        marks.appendChild(
          svg('polygon', {
            points: `${path(points)} ${path(bases.reverse())}`,
            fill: series.color,
            stroke: series.stroke?.color,
          }),
        );
        continue;
      }
      if (series.connected && points.length > 1) {
        marks.appendChild(
          svg('polyline', {
            points: path(points),
            fill: 'none',
            stroke: series.color,
            'stroke-width': round(series.lineWidth),
            'stroke-linejoin': 'round',
            'stroke-linecap': 'round',
          }),
        );
      }
      if (series.markers || !series.connected || points.length === 1) {
        for (const [x, y] of points) {
          marks.appendChild(
            svg('circle', { cx: round(x), cy: round(y), r: 3.5, fill: series.color }),
          );
        }
      }
    }
  }
  root.append(marks, labels);

  // Axis lines and the plot border go on top of the series.
  const lines = svg('g', { class: 'av-chart-lines', fill: 'none' });
  if (model.plotBorder) {
    lines.appendChild(
      svg('rect', {
        x: round(plot.x),
        y: round(plot.y),
        width: round(plot.width),
        height: round(plot.height),
        stroke: model.plotBorder.color,
        'stroke-width': round(model.plotBorder.width),
      }),
    );
  }
  if (!categoryAxis.deleted && categoryAxis.line) {
    lines
      .appendChild(
        svg(
          'line',
          horizontal
            ? {
                x1: round(zero),
                x2: round(zero),
                y1: round(plot.y),
                y2: round(plot.y + plot.height),
              }
            : {
                x1: round(plot.x),
                x2: round(plot.x + plot.width),
                y1: round(zero),
                y2: round(zero),
              },
        ),
      )
      .setAttribute('stroke', categoryAxis.line);
  }
  root.appendChild(lines);
  return root;
}
