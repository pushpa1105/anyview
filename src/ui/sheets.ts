import { toAnyViewError } from '../core/errors';
import { formatLabel } from '../core/labels';
import type { RenderContext, RendererInstance } from '../core/types';
import { h } from './dom';
import { createGrid, type Grid, type GridSource } from './grid';

export interface SheetTab {
  name: string;
  /** Called the first time the tab is shown; the result is cached. */
  load(): GridSource | Promise<GridSource>;
}

/** Shared view for every tabular format: one virtualized grid plus sheet tabs. */
export function createSheetView(
  context: RenderContext,
  tabs: readonly SheetTab[],
): RendererInstance {
  const { container, labels } = context;
  const notice = h('div', { class: 'av-notice', attrs: { hidden: true } });
  const area = h('div', {
    style: { flex: '1 1 auto', 'min-height': '0', display: 'flex', 'flex-direction': 'column' },
  });
  const tabBar = h('div', {
    class: 'av-tabs',
    attrs: { role: 'tablist', 'aria-label': labels.sheets },
  });
  container.append(notice, area);
  if (tabs.length > 1) container.appendChild(tabBar);

  const cache = new Map<number, Promise<GridSource>>();
  const buttons: HTMLButtonElement[] = [];
  let grid: Grid | null = null;
  let zoom = 1;
  let active = -1;
  let destroyed = false;

  const show = async (index: number): Promise<void> => {
    const tab = tabs[index];
    if (!tab || index === active) return;
    active = index;
    buttons.forEach((button, i) => button.setAttribute('aria-selected', String(i === index)));
    grid?.destroy();
    grid = null;
    area.textContent = '';
    notice.hidden = true;
    let source: GridSource;
    try {
      let pending = cache.get(index);
      if (!pending) {
        pending = Promise.resolve().then(() => tab.load());
        cache.set(index, pending);
      }
      source = await pending;
    } catch (error) {
      cache.delete(index);
      if (destroyed || active !== index) return;
      area.appendChild(
        h('p', {
          class: 'av-status-text',
          style: { padding: '24px', margin: '0 auto' },
          text: toAnyViewError(error, 'PARSE_FAILED', `Could not read "${tab.name}"`).message,
        }),
      );
      return;
    }
    if (destroyed || active !== index) return;
    if (source.totalRows !== undefined && source.totalRows > source.rowCount) {
      notice.textContent = formatLabel(labels.truncated, {
        shown: source.rowCount.toLocaleString(),
        total: source.totalRows.toLocaleString(),
      });
      notice.hidden = false;
    }
    grid = createGrid(area, source);
    grid.setZoom(zoom);
  };

  tabs.forEach((tab, index) => {
    const button = h('button', {
      class: 'av-tab',
      text: tab.name,
      attrs: { type: 'button', role: 'tab', 'aria-selected': 'false' },
      on: { click: () => void show(index) },
    });
    buttons.push(button);
    tabBar.appendChild(button);
  });
  void show(0);

  return {
    zoom: {
      get: () => zoom,
      set(value: number): void {
        zoom = value;
        grid?.setZoom(value);
      },
    },
    printable: () => grid?.toTable() ?? null,
    destroy(): void {
      destroyed = true;
      grid?.destroy();
      grid = null;
      cache.clear();
    },
  };
}
