// React example. Copy this file into a React 17+ project that has `@buffmomo/anyview` installed.
import { useRef, useState } from 'react';
import { AnyView, type AnyViewHandle } from '@buffmomo/anyview/react';
import type { AnyViewError, LoadEventDetail, ViewerSource } from '@buffmomo/anyview';

export function FilePreview() {
  const [source, setSource] = useState<ViewerSource | null>(null);
  const [status, setStatus] = useState('Choose a file to preview.');
  const [dark, setDark] = useState(false);
  // The ref exposes the underlying viewer for imperative calls.
  const preview = useRef<AnyViewHandle>(null);

  const handleLoad = ({ format, file }: LoadEventDetail) => {
    setStatus(`${file.name || 'File'}: ${format.label}`);
  };
  const handleError = (error: AnyViewError) => {
    setStatus(`${error.code}: ${error.message}`);
  };

  return (
    <section style={{ display: 'flex', flexDirection: 'column', height: '100vh' }}>
      <header style={{ display: 'flex', gap: 12, alignItems: 'center', padding: 12 }}>
        <input type="file" onChange={(event) => setSource(event.target.files?.[0] ?? null)} />
        <button type="button" onClick={() => setSource('/files/sample.pdf')}>
          Open a URL
        </button>
        <button type="button" onClick={() => setDark((value) => !value)}>
          {dark ? 'Light' : 'Dark'} theme
        </button>
        <button type="button" onClick={() => preview.current?.viewer?.zoomIn()}>
          Zoom in
        </button>
        <button type="button" onClick={() => void preview.current?.viewer?.download()}>
          Download
        </button>
        <span>{status}</span>
      </header>

      {/* src, name, type and theme are reactive. Give the component a height. */}
      <AnyView
        ref={preview}
        src={source}
        theme={dark ? 'dark' : 'light'}
        toolbar={{ print: false }}
        style={{ flex: '1 1 auto', minHeight: 0 }}
        onLoad={handleLoad}
        onError={handleError}
        onPageChange={({ page, pageCount }) => setStatus(`Page ${page} of ${pageCount}`)}
      />
    </section>
  );
}
