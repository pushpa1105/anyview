import { toAnyViewError } from '../core/errors';
import { formatLabel } from '../core/labels';
import type { Renderer } from '../core/types';
import { h } from '../ui/dom';
import { formatBytes, toBlobPart } from '../util/text';
import { openZip, type ZipEntry } from '../util/zip';

/** Rows drawn at most; larger archives are truncated with a notice. */
const MAX_LISTED = 5000;

export const renderer: Renderer = {
  async render(context) {
    const { container, labels, file } = context;
    const zip = openZip(await file.bytes());
    const entries = zip.entries
      .filter((entry) => !entry.isDirectory)
      .sort((a, b) => a.name.localeCompare(b.name));

    if (entries.length === 0) {
      container.appendChild(
        h('p', {
          class: 'av-status-text',
          style: { padding: '24px', margin: '0 auto' },
          text: labels.emptyArchive,
        }),
      );
      return;
    }
    if (entries.length > MAX_LISTED) {
      container.appendChild(
        h('div', {
          class: 'av-notice',
          text: formatLabel(labels.truncated, {
            shown: MAX_LISTED.toLocaleString(),
            total: entries.length.toLocaleString(),
          }),
        }),
      );
    }

    const status = h('div', { class: 'av-notice', attrs: { hidden: true, role: 'alert' } });
    const open = async (entry: ZipEntry): Promise<void> => {
      status.hidden = true;
      try {
        const data = await zip.readEntry(entry);
        if (context.signal.aborted) return;
        context.openChild(new Blob([toBlobPart(data)]), entry.name.split('/').pop() ?? entry.name);
      } catch (error) {
        status.textContent = toAnyViewError(
          error,
          'PARSE_FAILED',
          `Could not extract "${entry.name}"`,
        ).message;
        status.hidden = false;
      }
    };

    const body = h('tbody');
    for (const entry of entries.slice(0, MAX_LISTED)) {
      body.appendChild(
        h(
          'tr',
          null,
          h(
            'td',
            null,
            h('button', {
              class: 'av-link',
              text: entry.name,
              attrs: { type: 'button' },
              on: { click: () => void open(entry) },
            }),
          ),
          h('td', { text: formatBytes(entry.size) }),
        ),
      );
    }
    const table = h(
      'table',
      { class: 'av-list' },
      h(
        'thead',
        null,
        h('tr', null, h('th', { text: labels.name }), h('th', { text: labels.size })),
      ),
      body,
    );
    container.append(status, h('div', { class: 'av-scroll' }, table));
    return { printable: () => table };
  },
};
