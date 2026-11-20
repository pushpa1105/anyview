<!-- Vue 3 example. Copy this file into a Vue 3.2+ project that has `@buffmomo/anyview` installed. -->
<script setup lang="ts">
import { ref } from 'vue';
import { AnyView } from '@buffmomo/anyview/vue';
import type { AnyViewError, LoadEventDetail, Viewer, ViewerSource } from '@buffmomo/anyview';

const source = ref<ViewerSource | null>(null);
const status = ref('Choose a file to preview.');
const dark = ref(false);
// The template ref exposes the underlying viewer for imperative calls.
const preview = ref<{ viewer: Viewer | null } | null>(null);

function pick(event: Event): void {
  source.value = (event.target as HTMLInputElement).files?.[0] ?? null;
}

function onLoad({ format, file }: LoadEventDetail): void {
  status.value = `${file.name || 'File'}: ${format.label}`;
}

function onError(error: AnyViewError): void {
  status.value = `${error.code}: ${error.message}`;
}
</script>

<template>
  <section class="preview">
    <header>
      <input type="file" @change="pick" />
      <button type="button" @click="source = '/files/sample.pdf'">Open a URL</button>
      <button type="button" @click="dark = !dark">{{ dark ? 'Light' : 'Dark' }} theme</button>
      <button type="button" @click="preview?.viewer?.zoomIn()">Zoom in</button>
      <span>{{ status }}</span>
    </header>

    <!-- src, name, type and theme are reactive. Give the component a height. -->
    <AnyView
      ref="preview"
      class="viewer"
      :src="source"
      :theme="dark ? 'dark' : 'light'"
      :toolbar="{ print: false }"
      @load="onLoad"
      @error="onError"
      @pagechange="({ page, pageCount }) => (status = `Page ${page} of ${pageCount}`)"
    />
  </section>
</template>

<style scoped>
.preview {
  display: flex;
  flex-direction: column;
  height: 100vh;
}
header {
  display: flex;
  gap: 12px;
  align-items: center;
  padding: 12px;
}
.viewer {
  flex: 1 1 auto;
  min-height: 0;
}
</style>
