import { formatLabel, type Labels } from '../core/labels';
import type { ToolbarOptions } from '../core/types';
import { h, iconButton } from './dom';

export interface ToolbarHandlers {
  zoomIn: () => void;
  zoomOut: () => void;
  resetZoom: () => void;
  previousPage: () => void;
  nextPage: () => void;
  download: () => void;
  print: () => void;
  fullscreen: () => void;
}

export interface ToolbarState {
  fileName: string;
  formatLabel: string;
  zoom: number | null;
  page: { page: number; pageCount: number } | null;
  hasFile: boolean;
  canPrint: boolean;
}

export interface Toolbar {
  readonly element: HTMLElement;
  update(state: ToolbarState): void;
}

export function createToolbar(
  labels: Readonly<Labels>,
  config: Required<ToolbarOptions>,
  handlers: ToolbarHandlers,
): Toolbar {
  const name = h('span', { class: 'av-file-name' });
  const kind = h('span', { class: 'av-file-kind' });
  const title = h('div', { class: 'av-toolbar-title' }, name, kind);

  const previous = iconButton('previous', labels.previousPage, handlers.previousPage);
  const next = iconButton('next', labels.nextPage, handlers.nextPage);
  const pageInfo = h('span', { class: 'av-toolbar-info', attrs: { 'aria-live': 'polite' } });
  const pages = h('div', { class: 'av-toolbar-group' }, previous, pageInfo, next);

  const zoomOut = iconButton('zoomOut', labels.zoomOut, handlers.zoomOut);
  const zoomIn = iconButton('zoomIn', labels.zoomIn, handlers.zoomIn);
  const zoomInfo = h('button', {
    class: 'av-btn av-toolbar-info',
    attrs: { type: 'button', title: labels.resetZoom, 'aria-label': labels.resetZoom },
    on: { click: handlers.resetZoom },
  });
  const zoom = h('div', { class: 'av-toolbar-group' }, zoomOut, zoomInfo, zoomIn);

  const download = iconButton('download', labels.download, handlers.download);
  const print = iconButton('print', labels.print, handlers.print);
  const fullscreen = iconButton('fullscreen', labels.fullscreen, handlers.fullscreen);
  const actions = h('div', { class: 'av-toolbar-group' }, download, print, fullscreen);

  const element = h(
    'div',
    { class: 'av-toolbar', attrs: { role: 'toolbar' } },
    title,
    pages,
    zoom,
    actions,
  );
  const canFullscreen =
    typeof document !== 'undefined' &&
    typeof document.documentElement.requestFullscreen === 'function';

  return {
    element,
    update(state: ToolbarState): void {
      title.style.visibility = config.fileName ? '' : 'hidden';
      name.textContent = state.fileName;
      name.title = state.fileName;
      kind.textContent = state.formatLabel;

      pages.hidden = !config.pagination || state.page === null || state.page.pageCount < 2;
      if (state.page) {
        pageInfo.textContent = formatLabel(labels.pageOf, {
          page: state.page.page,
          count: state.page.pageCount,
        });
        previous.disabled = state.page.page <= 1;
        next.disabled = state.page.page >= state.page.pageCount;
      }

      zoom.hidden = !config.zoom || state.zoom === null;
      if (state.zoom !== null) zoomInfo.textContent = `${Math.round(state.zoom * 100)}%`;

      download.hidden = !config.download || !state.hasFile;
      print.hidden = !config.print || !state.canPrint;
      fullscreen.hidden = !config.fullscreen || !canFullscreen;
    },
  };
}
