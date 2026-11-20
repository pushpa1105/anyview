import { createViewer, isAnyViewError, type Viewer } from '@buffmomo/anyview';

interface Attachment {
  id: number;
  name: string;
  url: string;
}

/** Opens an attachment in a preview pane and reports what happened. */
export async function preview(pane: HTMLElement, attachment: Attachment): Promise<Viewer> {
  const viewer = createViewer(pane, { theme: 'auto' });

  viewer.on('pagechange', ({ page, pageCount }) => {
    console.log(`Page ${page} of ${pageCount}`);
  });

  try {
    await viewer.load(attachment.url, { name: attachment.name });
  } catch (error) {
    if (isAnyViewError(error) && error.code === 'UNSUPPORTED_FORMAT') {
      console.warn(`No preview for ${attachment.name}`);
    } else {
      throw error;
    }
  }
  return viewer;
}
