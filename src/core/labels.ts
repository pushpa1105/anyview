/** Every piece of UI text. Override any of them through the `labels` option. */
export interface Labels {
  loading: string;
  empty: string;
  errorTitle: string;
  unsupportedTitle: string;
  unsupportedHint: string;
  download: string;
  print: string;
  fullscreen: string;
  exitFullscreen: string;
  zoomIn: string;
  zoomOut: string;
  resetZoom: string;
  previousPage: string;
  nextPage: string;
  /** `{page}` and `{count}` are replaced. */
  pageOf: string;
  back: string;
  /** `{shown}` and `{total}` are replaced. */
  truncated: string;
  sheets: string;
  attachments: string;
  name: string;
  size: string;
  emptyArchive: string;
  notebookOutput: string;
  from: string;
  to: string;
  cc: string;
  subject: string;
  date: string;
}

export const DEFAULT_LABELS: Readonly<Labels> = Object.freeze({
  loading: 'Loading…',
  empty: 'No file selected',
  errorTitle: 'This file could not be displayed',
  unsupportedTitle: 'Preview is not available for this file',
  unsupportedHint: 'You can download the file and open it with another application.',
  download: 'Download',
  print: 'Print',
  fullscreen: 'Full screen',
  exitFullscreen: 'Exit full screen',
  zoomIn: 'Zoom in',
  zoomOut: 'Zoom out',
  resetZoom: 'Reset zoom',
  previousPage: 'Previous',
  nextPage: 'Next',
  pageOf: '{page} / {count}',
  back: 'Back',
  truncated: 'Showing the first {shown} of {total}.',
  sheets: 'Sheets',
  attachments: 'Attachments',
  name: 'Name',
  size: 'Size',
  emptyArchive: 'This archive is empty.',
  notebookOutput: 'Output',
  from: 'From',
  to: 'To',
  cc: 'Cc',
  subject: 'Subject',
  date: 'Date',
});

/** Replaces `{key}` placeholders in a label. */
export function formatLabel(template: string, values: Record<string, string | number>): string {
  return template.replace(/\{(\w+)\}/g, (match, key: string) =>
    key in values ? String(values[key]) : match,
  );
}
