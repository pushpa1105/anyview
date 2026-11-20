import { extensionOf } from '../util/url';
import { getFormatById } from './formats';
import type { FormatInfo, Renderer, RendererDefinition } from './types';

function lazy(
  id: string,
  formats: string[],
  loader: () => Promise<{ renderer: Renderer }>,
): RendererDefinition {
  return { id, formats, load: async () => (await loader()).renderer };
}

/**
 * Built-in renderers. Each one lives in its own module behind a dynamic
 * `import()`, so bundlers emit a separate chunk per format and a page that only
 * ever shows images never downloads the spreadsheet or document code.
 */
export const BUILTIN_RENDERERS: readonly RendererDefinition[] = Object.freeze([
  lazy('image', ['image', 'svg'], () => import('../renderers/image')),
  lazy('media', ['video', 'audio'], () => import('../renderers/media')),
  lazy('pdf', ['pdf'], () => import('../renderers/pdf')),
  lazy('text', ['text', 'json', 'xml'], () => import('../renderers/text')),
  lazy('csv', ['csv', 'tsv'], () => import('../renderers/csv')),
  lazy('markdown', ['markdown'], () => import('../renderers/markdown')),
  lazy('html', ['html'], () => import('../renderers/html')),
  lazy('xlsx', ['xlsx'], () => import('../renderers/xlsx')),
  lazy('ods', ['ods'], () => import('../renderers/ods')),
  lazy('docx', ['docx'], () => import('../renderers/docx')),
  lazy('odt', ['odt'], () => import('../renderers/odt')),
  lazy('pptx', ['pptx'], () => import('../renderers/pptx')),
  lazy('odp', ['odp'], () => import('../renderers/odp')),
  lazy('rtf', ['rtf'], () => import('../renderers/rtf')),
  lazy('epub', ['epub'], () => import('../renderers/epub')),
  lazy('zip', ['zip'], () => import('../renderers/archive')),
  lazy('ipynb', ['ipynb'], () => import('../renderers/notebook')),
  lazy('eml', ['eml'], () => import('../renderers/email')),
  lazy('font', ['font'], () => import('../renderers/font')),
]);

/** Format ids that have a built-in renderer. */
export const SUPPORTED_FORMAT_IDS: readonly string[] = Object.freeze(
  BUILTIN_RENDERERS.flatMap((definition) => definition.formats ?? []),
);

export interface CustomMatch {
  definition: RendererDefinition;
  format: FormatInfo;
}

/**
 * Finds a custom renderer that claims a file by extension or MIME type. This
 * runs before format detection so applications can add formats AnyView has
 * never heard of.
 */
export function matchCustomRenderer(
  custom: readonly RendererDefinition[],
  name: string,
  mimeType: string,
): CustomMatch | null {
  const extension = extensionOf(name);
  for (const definition of custom) {
    const byExtension =
      extension !== '' &&
      definition.extensions?.some((item) => item.replace(/^\./, '').toLowerCase() === extension);
    const byMime =
      mimeType !== '' && definition.mimeTypes?.some((item) => item.toLowerCase() === mimeType);
    if (!byExtension && !byMime) continue;
    const known = definition.formats?.map(getFormatById).find((info) => info !== null);
    const format: FormatInfo = known ?? {
      id: definition.formats?.[0] ?? definition.id,
      label: (extension || definition.id).toUpperCase(),
      category: 'other',
      extensions: definition.extensions ?? [],
      mimeTypes: definition.mimeTypes ?? [],
    };
    return { definition, format };
  }
  return null;
}

/** Picks the renderer for a detected format: custom definitions win over built-ins. */
export function findRenderer(
  formatId: string,
  custom: readonly RendererDefinition[],
): RendererDefinition | null {
  for (const definition of custom) {
    // A custom renderer reusing a built-in id inherits that renderer's formats.
    const formats =
      definition.formats ??
      BUILTIN_RENDERERS.find((builtin) => builtin.id === definition.id)?.formats;
    if (formats?.includes(formatId)) return definition;
  }
  return BUILTIN_RENDERERS.find((definition) => definition.formats?.includes(formatId)) ?? null;
}
