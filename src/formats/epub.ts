import { AnyViewError } from '../core/errors';
import { bytesToBinaryString, decodeText } from '../util/text';
import { attr, child, childElements, find, parseXml, textOf } from '../util/xml';
import { resolveZipPath, type ZipArchive } from '../util/zip';
import { imageMimeType } from './ooxml';

export interface EpubBook {
  title: string;
  chapterCount: number;
  /**
   * Returns a self-contained HTML document for a chapter: scripts and other
   * active content are removed, and images and stylesheets are inlined so the
   * chapter can be shown in a sandboxed frame with no network access to the book.
   */
  chapterHtml(index: number, zoom?: number): Promise<string>;
}

/** Images larger than this are left out rather than inlined as data URLs. */
const MAX_INLINE_IMAGE = 8 * 1024 * 1024;
const REMOVED_ELEMENTS =
  'script, iframe, object, embed, base, form, meta[http-equiv], link:not([rel~="stylesheet"])';
const READER_CSS =
  'html{color-scheme:light dark}body{max-width:44em;margin:0 auto;padding:24px 28px;line-height:1.6;font-family:Georgia,"Times New Roman",serif;overflow-wrap:break-word}img,svg{max-width:100%;height:auto}';

const VOID_ELEMENTS = new Set([
  'area',
  'base',
  'br',
  'col',
  'embed',
  'hr',
  'img',
  'input',
  'link',
  'meta',
  'source',
  'track',
  'wbr',
]);

/**
 * EPUB chapters are XHTML, where `<title/>` or `<a id="x"/>` are complete
 * elements. An HTML parser would treat them as unclosed start tags and swallow
 * the rest of the page, so they are rewritten as explicit open/close pairs.
 */
export function expandSelfClosing(xhtml: string): string {
  return xhtml.replace(
    /<([a-zA-Z][\w:-]*)((?:\s[^<>]*?)?)\/>/g,
    (match, name: string, attributes: string) =>
      VOID_ELEMENTS.has(name.toLowerCase()) ? match : `<${name}${attributes}></${name}>`,
  );
}

function decodePath(path: string): string {
  const clean = path.split('#')[0]!.split('?')[0]!;
  try {
    return decodeURIComponent(clean);
  } catch {
    return clean;
  }
}

/** Opens an EPUB package and reads its reading order. */
export async function openEpub(zip: ZipArchive): Promise<EpubBook> {
  const container = await zip.readText('META-INF/container.xml');
  if (!container) {
    throw new AnyViewError(
      'PARSE_FAILED',
      'This is not a valid EPUB: "META-INF/container.xml" is missing.',
    );
  }
  const packagePath = attr(find(parseXml(container), 'rootfiles', 'rootfile'), 'full-path');
  const packageXml = packagePath ? await zip.readText(packagePath) : null;
  if (!packagePath || !packageXml) {
    throw new AnyViewError(
      'PARSE_FAILED',
      'This is not a valid EPUB: the package document is missing.',
    );
  }
  const packageDocument = parseXml(packageXml);
  const manifest = new Map<string, string>();
  for (const item of childElements(child(packageDocument, 'manifest'), 'item')) {
    const id = attr(item, 'id');
    const href = attr(item, 'href');
    if (id && href) manifest.set(id, resolveZipPath(packagePath, decodePath(href)));
  }
  const chapters: string[] = [];
  for (const reference of childElements(child(packageDocument, 'spine'), 'itemref')) {
    const path = manifest.get(attr(reference, 'idref') ?? '');
    if (path && zip.has(path)) chapters.push(path);
  }
  if (chapters.length === 0) {
    throw new AnyViewError('PARSE_FAILED', 'This EPUB does not contain any readable chapters.');
  }
  const title = textOf(find(packageDocument, 'metadata', 'title')).trim();

  const inlineImage = async (
    chapterPath: string,
    reference: string | null,
  ): Promise<string | null> => {
    if (!reference || /^[a-z][\w+.-]*:/i.test(reference)) return null;
    const path = resolveZipPath(chapterPath, decodePath(reference));
    const mimeType = imageMimeType(path);
    const entry = zip.get(path);
    if (!mimeType || !entry || entry.size > MAX_INLINE_IMAGE) return null;
    return `data:${mimeType};base64,${btoa(bytesToBinaryString(await zip.readEntry(entry)))}`;
  };

  return {
    title,
    chapterCount: chapters.length,
    async chapterHtml(index: number, zoom = 1): Promise<string> {
      const path = chapters[index];
      const data = path ? await zip.read(path) : null;
      if (!path || !data) {
        throw new AnyViewError('PARSE_FAILED', `Chapter ${index + 1} is missing from this EPUB.`);
      }
      // DOMParser produces an inert document: nothing in it loads or runs.
      const chapter = new DOMParser().parseFromString(
        expandSelfClosing(decodeText(data)),
        'text/html',
      );
      for (const element of Array.from(chapter.querySelectorAll(REMOVED_ELEMENTS)))
        element.remove();

      for (const element of Array.from(chapter.querySelectorAll('*'))) {
        for (const attribute of Array.from(element.attributes)) {
          const name = attribute.name.toLowerCase();
          // Event handlers and in-book links (which would navigate the frame away) are dropped.
          if (name.startsWith('on') || (element.localName === 'a' && name === 'href')) {
            element.removeAttribute(attribute.name);
          }
        }
      }

      for (const image of Array.from(chapter.querySelectorAll('img'))) {
        const inlined = await inlineImage(path, image.getAttribute('src'));
        image.removeAttribute('srcset');
        if (inlined) image.setAttribute('src', inlined);
        else image.removeAttribute('src');
      }
      for (const image of Array.from(chapter.querySelectorAll('image'))) {
        const reference = image.getAttribute('xlink:href') ?? image.getAttribute('href');
        const inlined = await inlineImage(path, reference);
        image.removeAttribute('xlink:href');
        if (inlined) image.setAttribute('href', inlined);
        else image.removeAttribute('href');
      }
      for (const link of Array.from(chapter.querySelectorAll('link'))) {
        const href = link.getAttribute('href');
        const css = href ? await zip.readText(resolveZipPath(path, decodePath(href))) : null;
        if (css !== null) {
          const style = chapter.createElement('style');
          // External references inside the book's CSS (fonts, images) are not resolved.
          style.textContent = css.replace(/@import[^;]+;/g, '');
          link.replaceWith(style);
        } else {
          link.remove();
        }
      }

      const base = chapter.createElement('style');
      base.textContent = zoom === 1 ? READER_CSS : `${READER_CSS}body{zoom:${zoom}}`;
      (chapter.head ?? chapter.documentElement).prepend(base);
      return `<!DOCTYPE html>${chapter.documentElement.outerHTML}`;
    },
  };
}
