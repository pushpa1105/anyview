import type { Renderer } from '../core/types';
import { createSandboxedFrame } from '../ui/frame';

export const renderer: Renderer = {
  async render(context) {
    const { file, container } = context;
    const html = await file.text();
    container.appendChild(createSandboxedFrame(html, file.name || 'HTML document', file.url));
  },
};
