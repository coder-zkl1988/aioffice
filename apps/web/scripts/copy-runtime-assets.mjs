import { copyFile, mkdir } from 'node:fs/promises'
import { createRequire } from 'node:module'
import { dirname, join } from 'node:path'

const require = createRequire(import.meta.url)
const output = new URL('../dist-server/wasm/', import.meta.url)
await mkdir(output, { recursive: true })

await copyFile(require.resolve('@embedpdf/pdfium/pdfium.wasm'), new URL('pdfium.wasm', output))

const harfbuzzEntry = require.resolve('harfbuzzjs')
await copyFile(
  join(dirname(harfbuzzEntry), 'harfbuzz-subset.wasm'),
  new URL('harfbuzz-subset.wasm', output),
)
