import { createRequire } from 'node:module'
import { dirname, join } from 'node:path'
import react from '@vitejs/plugin-react'
import { defineConfig, normalizePath } from 'vite'
import { viteStaticCopy } from 'vite-plugin-static-copy'

// renderer-only dev server (embedded by shell via PDF_RENDERER_URL for HMR; no standalone Electron)
const require = createRequire(import.meta.url)
const pdfjsRoot = dirname(dirname(require.resolve('pdfjs-dist/package.json')))
// vite-plugin-static-copy globs require POSIX separators; join() breaks on Windows
const pdfjsDir = (sub: string) => normalizePath(join(pdfjsRoot, 'pdfjs-dist', sub))
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

export default defineConfig({
  root: 'src/renderer',
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
  server: {
    port: Number(process.env.PDF_DEV_PORT) || 5176,
    strictPort: true,
  },
})
