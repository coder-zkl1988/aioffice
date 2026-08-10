import react from '@vitejs/plugin-react'
import { createRequire } from 'node:module'
import { dirname, join, resolve } from 'node:path'
import { defineConfig, normalizePath } from 'vite'
import { viteStaticCopy } from 'vite-plugin-static-copy'

const require = createRequire(import.meta.url)
const pdfjsRoot = dirname(dirname(require.resolve('pdfjs-dist/package.json')))
const pdfjsDir = (subpath: string) => normalizePath(join(pdfjsRoot, 'pdfjs-dist', subpath))

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
  plugins: [
    react(),
    viteStaticCopy({
      targets: [
        { src: pdfjsDir('cmaps'), dest: 'pdfjs' },
        { src: pdfjsDir('standard_fonts'), dest: 'pdfjs' },
        { src: pdfjsDir('wasm'), dest: 'pdfjs' },
      ],
    }),
  ],
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
        pdf: resolve(__dirname, 'pdf.html'),
        sheets: resolve(__dirname, 'sheets.html'),
      },
    },
  },
})
