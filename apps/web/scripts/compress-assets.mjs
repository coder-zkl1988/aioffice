import { readdir, readFile, stat, writeFile } from 'node:fs/promises'
import { fileURLToPath } from 'node:url'
import { extname, join } from 'node:path'
import { promisify } from 'node:util'
import { gzip } from 'node:zlib'

const compress = promisify(gzip)
const root = fileURLToPath(new URL('../dist/', import.meta.url))
const compressible = new Set(['.css', '.html', '.js', '.json', '.mjs', '.svg'])

async function visit(directory) {
  for (const entry of await readdir(directory, { withFileTypes: true })) {
    const path = join(directory, entry.name)
    if (entry.isDirectory()) {
      await visit(path)
      continue
    }
    if (!compressible.has(extname(entry.name)) || (await stat(path)).size < 1024) continue
    await writeFile(`${path}.gz`, await compress(await readFile(path)))
  }
}

await visit(root)
