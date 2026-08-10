import { existsSync } from 'node:fs'
import { createRequire } from 'node:module'
import { dirname, join } from 'node:path'

/**
 * Runtime wasm assets live in node_modules during dev/tests but the packaged app
 * ships no node_modules (everything is bundled) — electron-builder copies them
 * into Resources/wasm instead (see apps/shell/electron-builder.cjs extraResources).
 */
const packagedPath = (fileName: string) => {
  const resourcesPath = (process as NodeJS.Process & { resourcesPath?: string }).resourcesPath
  return join(resourcesPath || process.cwd(), 'wasm', fileName)
}
const webRuntimePath = (fileName: string) =>
  process.env.GENOFFICE_WASM_DIR ? join(process.env.GENOFFICE_WASM_DIR, fileName) : null

const req = () => createRequire(import.meta.url)

export function pdfiumWasmPath(): string {
  const webPath = webRuntimePath('pdfium.wasm')
  if (webPath) return webPath
  try {
    return req().resolve('@embedpdf/pdfium/pdfium.wasm')
  } catch {
    return packagedPath('pdfium.wasm')
  }
}

export function hbSubsetWasmPath(): string {
  const webPath = webRuntimePath('harfbuzz-subset.wasm')
  if (webPath) return webPath
  const r = req()
  try {
    // harfbuzzjs ≤0.10 ships hb-subset.wasm at the package root with no exports map
    return r.resolve('harfbuzzjs/hb-subset.wasm')
  } catch {
    /* fall through */
  }
  try {
    // harfbuzzjs ≥1.x seals subpaths; the wasm sits next to the exported entry point
    const p = join(dirname(r.resolve('harfbuzzjs')), 'harfbuzz-subset.wasm')
    if (existsSync(p)) return p
  } catch {
    /* fall through */
  }
  return packagedPath('hb-subset.wasm')
}
