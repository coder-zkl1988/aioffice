import react from '@vitejs/plugin-react'
import { resolve } from 'node:path'
import { defineConfig } from 'vite'

const TIPTAP_DEDUPE = [
  '@tiptap/core',
  '@tiptap/pm',
  '@tiptap/react',
  '@tiptap/extensions',
  '@tiptap/extension-list',
  '@tiptap/extension-table',
  '@tiptap/extension-image',
  '@tiptap/suggestion',
  '@tiptap/markdown',
  '@tiptap/extension-highlight',
  '@tiptap/extension-code-block',
]

export default defineConfig({
  base: './',
  plugins: [react()],
  resolve: { dedupe: TIPTAP_DEDUPE },
  server: {
    port: 5180,
    strictPort: true,
    fs: { allow: [resolve(__dirname, '../..')] },
  },
  build: {
    rollupOptions: {
      input: {
        workspace: resolve(__dirname, 'index.html'),
        docs: resolve(__dirname, 'docs.html'),
        markdown: resolve(__dirname, 'markdown.html'),
      },
    },
  },
})
