import react from '@vitejs/plugin-react'
import { createRequire } from 'node:module'
import { dirname, join, resolve } from 'node:path'
import { defineConfig, normalizePath } from 'vite'
import { viteStaticCopy } from 'vite-plugin-static-copy'

const require = createRequire(import.meta.url)
const pdfjsRoot = dirname(dirname(require.resolve('pdfjs-dist/package.json')))
const pdfjsDir = (subpath: string) => normalizePath(join(pdfjsRoot, 'pdfjs-dist', subpath))
const tesseractDir = dirname(require.resolve('tesseract.js/package.json'))
const tesseractCoreDir = dirname(require.resolve('tesseract.js-core/package.json'))
const tesseractEnglishDir = dirname(require.resolve('@tesseract.js-data/eng/package.json'))
const tesseractChineseDir = dirname(require.resolve('@tesseract.js-data/chi_sim/package.json'))

const ocrAssets = [
  { src: normalizePath(join(tesseractDir, 'dist/worker.min.js')), dest: 'ocr' },
  {
    src: normalizePath(join(tesseractCoreDir, 'tesseract-core*-lstm.wasm*')),
    dest: 'ocr/core',
  },
  {
    src: normalizePath(join(tesseractEnglishDir, '4.0.0_best_int/eng.traineddata.gz')),
    dest: 'ocr/lang',
  },
  {
    src: normalizePath(join(tesseractChineseDir, '4.0.0_best_int/chi_sim.traineddata.gz')),
    dest: 'ocr/lang',
  },
]

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
        ...ocrAssets,
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
        slides: resolve(__dirname, 'slides.html'),
      },
    },
  },
})
