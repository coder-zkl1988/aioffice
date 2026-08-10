import { access, chmod, copyFile, mkdir } from 'node:fs/promises'
import { createRequire } from 'node:module'
import { fileURLToPath } from 'node:url'
import { dirname, join, resolve } from 'node:path'

const require = createRequire(import.meta.url)
const output = new URL('../dist-server/wasm/', import.meta.url)
await mkdir(output, { recursive: true })

await copyFile(require.resolve('@embedpdf/pdfium/pdfium.wasm'), new URL('pdfium.wasm', output))

const harfbuzzEntry = require.resolve('harfbuzzjs')
await copyFile(
  join(dirname(harfbuzzEntry), 'harfbuzz-subset.wasm'),
  new URL('harfbuzz-subset.wasm', output),
)

const nativeOutput = new URL('../dist-server/native/', import.meta.url)
await mkdir(nativeOutput, { recursive: true })
const sidecarName = process.platform === 'win32' ? 'xlsx-sidecar.exe' : 'xlsx-sidecar'
const sidecarPath = resolve(
  fileURLToPath(new URL('../../sheets/native/xlsx-engine/target/release/', import.meta.url)),
  sidecarName,
)
await access(sidecarPath)
const sidecarOutput = new URL('xlsx-sidecar', nativeOutput)
await copyFile(sidecarPath, sidecarOutput)
if (process.platform !== 'win32') await chmod(sidecarOutput, 0o755)
