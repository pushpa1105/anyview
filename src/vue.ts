import {
  defineComponent,
  h,
  onBeforeUnmount,
  onMounted,
  ref,
  shallowRef,
  watch,
  type PropType,
} from 'vue';
import { createViewer } from './core/viewer';
import type { AnyViewError } from './core/errors';
import type {
  ToolbarOptions,
  Viewer,
  ViewerEventMap,
  ViewerOptions,
  ViewerSource,
  ViewerTheme,
} from './core/types';

/**
 * Vue 3 component for AnyView.
 *
 * `src`, `name`, `type` and `theme` are reactive. `toolbar` and `options` are
 * read when the component mounts; change the component's `key` to apply new values.
 *
 * @example
 * <AnyView :src="file" style="height: 600px" @error="onError" />
 */
export const AnyView = defineComponent({
  name: 'AnyView',
  props: {
    /** File to show: a URL, `File`, `Blob` or bytes. */
    src: { type: [String, Object] as PropType<ViewerSource | null>, default: null },
    /** File name override. */
    name: { type: String, default: undefined },
    /** Force a format: a format id, an extension or a MIME type. */
    type: { type: String, default: undefined },
    theme: { type: String as PropType<ViewerTheme>, default: 'auto' },
    toolbar: { type: [Boolean, Object] as PropType<boolean | ToolbarOptions>, default: true },
    /** Remaining viewer options (renderers, pdfjs, labels, fetchOptions, maxFileSize…). */
    options: { type: Object as PropType<ViewerOptions>, default: undefined },
  },
  emits: {
    loadstart: (_detail: ViewerEventMap['loadstart']) => true,
    load: (_detail: ViewerEventMap['load']) => true,
    error: (_error: AnyViewError) => true,
    pagechange: (_detail: ViewerEventMap['pagechange']) => true,
    zoomchange: (_detail: ViewerEventMap['zoomchange']) => true,
  },
  setup(props, { emit, expose }) {
    const host = ref<HTMLElement | null>(null);
    const viewer = shallowRef<Viewer | null>(null);

    const open = (): void => {
      const instance = viewer.value;
      if (!instance) return;
      const source = props.src;
      if (source === null || source === undefined || source === '') {
        instance.clear();
        return;
      }
      // Failures are reported through the "error" event and the viewer's own UI.
      instance
        .load(source, {
          ...(props.name ? { name: props.name } : {}),
          ...(props.type ? { type: props.type } : {}),
        })
        .catch(() => undefined);
    };

    onMounted(() => {
      if (!host.value) return;
      const instance = createViewer(host.value, {
        ...props.options,
        src: null,
        theme: props.theme,
        toolbar: props.toolbar,
      });
      viewer.value = instance;
      instance.on('loadstart', (detail) => emit('loadstart', detail));
      instance.on('load', (detail) => emit('load', detail));
      instance.on('error', (error) => emit('error', error));
      instance.on('pagechange', (detail) => emit('pagechange', detail));
      instance.on('zoomchange', (detail) => emit('zoomchange', detail));
      open();
    });

    watch(() => [props.src, props.name, props.type], open);
    watch(
      () => props.theme,
      (theme) => viewer.value?.setTheme(theme),
    );

    onBeforeUnmount(() => {
      viewer.value?.destroy();
      viewer.value = null;
    });

    expose({ viewer });
    return () => h('div', { ref: host, style: { height: '100%' } });
  },
});
