import { readFileSync } from 'node:fs';
import { defineConfig } from 'tsup';

const { version } = JSON.parse(readFileSync('./package.json', 'utf8')) as { version: string };

export default defineConfig({
  entry: {
    index: 'src/index.ts',
    element: 'src/element.ts',
    react: 'src/react.ts',
    vue: 'src/vue.ts',
  },
  format: ['esm', 'cjs'],
  target: 'es2022',
  platform: 'browser',
  dts: true,
  // Renderers are loaded with dynamic import(); splitting keeps each one in its own chunk.
  splitting: true,
  sourcemap: true,
  clean: true,
  treeshake: true,
  external: ['react', 'vue'],
  define: { __ANYVIEW_VERSION__: JSON.stringify(version) },
});
