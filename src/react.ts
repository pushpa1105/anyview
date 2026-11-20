import {
  createElement,
  forwardRef,
  useEffect,
  useImperativeHandle,
  useRef,
  type CSSProperties,
  type ForwardRefExoticComponent,
  type RefAttributes,
} from 'react';
import { createViewer } from './core/viewer';
import type { AnyViewError } from './core/errors';
import type { Viewer, ViewerEventMap, ViewerOptions, ViewerSource } from './core/types';

export interface AnyViewProps extends Omit<ViewerOptions, 'src' | 'className'> {
  /** File to show: a URL, `File`, `Blob` or bytes. `null`/`undefined` shows the empty state. */
  src?: ViewerSource | null;
  /** Class name of the wrapper element. */
  className?: string;
  /** Inline style of the wrapper element. Give it a height, e.g. `{ height: 600 }`. */
  style?: CSSProperties;
  id?: string;
  onLoadStart?: (detail: ViewerEventMap['loadstart']) => void;
  onLoad?: (detail: ViewerEventMap['load']) => void;
  onError?: (error: AnyViewError) => void;
  onPageChange?: (detail: ViewerEventMap['pagechange']) => void;
  onZoomChange?: (detail: ViewerEventMap['zoomchange']) => void;
}

/** What a `ref` on `<AnyView>` gives access to. */
export interface AnyViewHandle {
  /** The underlying viewer (`null` before mount and after unmount). */
  readonly viewer: Viewer | null;
}

/**
 * React component for AnyView.
 *
 * `src`, `name`, `type` and `theme` are reactive. The remaining options
 * (`toolbar`, `renderers`, `pdfjs`, `labels`, `fetchOptions`, `maxFileSize`)
 * are read when the component mounts; change the component's `key` to apply
 * new values.
 *
 * @example
 * <AnyView src={file} style={{ height: 600 }} onError={(e) => console.warn(e.code)} />
 */
export const AnyView: ForwardRefExoticComponent<AnyViewProps & RefAttributes<AnyViewHandle>> =
  forwardRef<AnyViewHandle, AnyViewProps>(function AnyView(props, ref) {
    const { src, name, type, theme, className, style, id } = props;
    const host = useRef<HTMLDivElement | null>(null);
    const viewer = useRef<Viewer | null>(null);
    // Always call the latest callbacks without re-subscribing on every render.
    const latest = useRef(props);
    latest.current = props;

    useImperativeHandle(
      ref,
      () => ({
        get viewer() {
          return viewer.current;
        },
      }),
      [],
    );

    useEffect(() => {
      if (!host.current) return undefined;
      const options = latest.current;
      const instance = createViewer(host.current, {
        toolbar: options.toolbar ?? true,
        theme: options.theme ?? 'auto',
        ...(options.maxFileSize !== undefined ? { maxFileSize: options.maxFileSize } : {}),
        ...(options.fetchOptions ? { fetchOptions: options.fetchOptions } : {}),
        ...(options.renderers ? { renderers: options.renderers } : {}),
        ...(options.pdfjs ? { pdfjs: options.pdfjs } : {}),
        ...(options.labels ? { labels: options.labels } : {}),
        ...(options.styleNonce ? { styleNonce: options.styleNonce } : {}),
      });
      viewer.current = instance;
      instance.on('loadstart', (detail) => latest.current.onLoadStart?.(detail));
      instance.on('load', (detail) => latest.current.onLoad?.(detail));
      instance.on('error', (error) => latest.current.onError?.(error));
      instance.on('pagechange', (detail) => latest.current.onPageChange?.(detail));
      instance.on('zoomchange', (detail) => latest.current.onZoomChange?.(detail));
      return () => {
        instance.destroy();
        viewer.current = null;
      };
    }, []);

    useEffect(() => {
      viewer.current?.setTheme(theme ?? 'auto');
    }, [theme]);

    useEffect(() => {
      const instance = viewer.current;
      if (!instance) return;
      if (src === null || src === undefined || src === '') {
        instance.clear();
        return;
      }
      // Failures are reported through onError and the viewer's own UI.
      instance
        .load(src, { ...(name ? { name } : {}), ...(type ? { type } : {}) })
        .catch(() => undefined);
    }, [src, name, type]);

    return createElement('div', {
      ref: host,
      id,
      className,
      style: { height: '100%', ...style },
    });
  });
