import { createRequire } from 'node:module'
import { dirname, join } from 'node:path'
import react from '@vitejs/plugin-react'
import { defineConfig, externalizeDepsPlugin } from 'electron-vite'
import { normalizePath } from 'vite'
import { viteStaticCopy } from 'vite-plugin-static-copy'

// Non-embedded CMaps/standard fonts (e.g. CJK) need pdfjs data dirs, shipped with renderer output
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
  // @genoffice/i18n ships as TS source; pdf-lib's package only includes out/** — both must be bundled
  main: {
    plugins: [
      externalizeDepsPlugin({
        exclude: ['@genoffice/i18n', 'pdf-lib', '@genoffice/electron-utils'],
      }),
    ],
  },
  preload: {
    plugins: [externalizeDepsPlugin({ exclude: ['@genoffice/i18n'] })],
  },
  renderer: {
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
      strictPort: Boolean(process.env.PDF_DEV_PORT),
    },
  },
})
