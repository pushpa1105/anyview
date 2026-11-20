import type { Renderer } from '../core/types';
import { parseEmail } from '../formats/eml';
import { h } from '../ui/dom';
import { createSandboxedFrame } from '../ui/frame';
import { formatBytes, toBlobPart } from '../util/text';

export const renderer: Renderer = {
  async render(context) {
    const { container, labels } = context;
    const message = parseEmail(await context.file.bytes());

    const head = h('dl', { class: 'av-mail-head' });
    const rows: [string, string][] = [
      [labels.subject, 'subject'],
      [labels.from, 'from'],
      [labels.to, 'to'],
      [labels.cc, 'cc'],
      [labels.date, 'date'],
    ];
    for (const [label, name] of rows) {
      const value = message.headers.get(name);
      if (value) head.append(h('dt', { text: label }), h('dd', { text: value }));
    }
    const view = h('div', { class: 'av-mail' }, head);

    if (message.attachments.length > 0) {
      const list = h('div', {
        class: 'av-mail-attachments',
        attrs: { 'aria-label': labels.attachments },
      });
      for (const attachment of message.attachments) {
        list.appendChild(
          h('button', {
            class: 'av-link',
            text: `${attachment.name} (${formatBytes(attachment.data.length)})`,
            attrs: { type: 'button' },
            on: {
              click: () =>
                context.openChild(
                  new Blob([toBlobPart(attachment.data)], { type: attachment.mimeType }),
                  attachment.name,
                ),
            },
          }),
        );
      }
      view.appendChild(list);
    }

    if (message.html !== null) {
      // Email HTML is untrusted: it is shown in a frame that cannot run scripts.
      view.appendChild(
        createSandboxedFrame(message.html, message.headers.get('subject') ?? 'Email'),
      );
    } else {
      view.appendChild(h('pre', { class: 'av-mail-text', text: message.text ?? '' }));
    }
    container.appendChild(view);
    return { printable: () => view };
  },
};
