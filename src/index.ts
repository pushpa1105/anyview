/**
 * AnyView — a framework-agnostic, zero-dependency universal file viewer.
 *
 * @example
 * import { createViewer } from '@buffmomo/anyview';
 *
 * const viewer = createViewer('#preview', { src: '/files/report.xlsx' });
 * viewer.on('error', (error) => console.warn(error.code, error.message));
 *
 * @packageDocumentation
 */
export { createViewer } from './core/viewer';
export { AnyViewError, ERROR_CODES, isAnyViewError, toAnyViewError } from './core/errors';
export { DEFAULT_LABELS } from './core/labels';
export {
  FORMATS,
  detectFormat,
  getFormatByExtension,
  getFormatById,
  getFormatByMimeType,
  sniffFormat,
} from './core/formats';
export { SUPPORTED_FORMAT_IDS } from './core/registry';

export type { AnyViewErrorCode, AnyViewErrorOptions } from './core/errors';
export type { Labels } from './core/labels';
export type { DetectInput } from './core/formats';
export type {
  FormatCategory,
  FormatInfo,
  LoadedFile,
  LoadEventDetail,
  LoadOptions,
  PageCapability,
  PdfJsDocument,
  PdfJsLike,
  PdfJsLoadingTask,
  PdfJsPage,
  PdfJsViewport,
  RenderContext,
  Renderer,
  RendererDefinition,
  RendererInstance,
  ResolvedViewerOptions,
  ToolbarOptions,
  Viewer,
  ViewerEventHandler,
  ViewerEventMap,
  ViewerEventName,
  ViewerOptions,
  ViewerSource,
  ViewerState,
  ViewerTheme,
  ZoomCapability,
} from './core/types';

/** The package version, replaced at build time. */
export const VERSION: string = __ANYVIEW_VERSION__;
