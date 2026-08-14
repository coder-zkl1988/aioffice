import type { PdfJsonImportFonts } from '@genoffice/pdf-tools'

import boldUrl from '../../../docs/src/renderer/fonts/LiberationSans-Bold.ttf?url'
import boldItalicUrl from '../../../docs/src/renderer/fonts/LiberationSans-BoldItalic.ttf?url'
import italicUrl from '../../../docs/src/renderer/fonts/LiberationSans-Italic.ttf?url'
import regularUrl from '../../../docs/src/renderer/fonts/LiberationSans-Regular.ttf?url'
import unicodeUrl from '../../../docs/src/renderer/fonts/NotoSansSC-Regular-subset.ttf?url'

async function loadFont(url: string): Promise<Uint8Array> {
  const response = await fetch(url)
  if (!response.ok) throw new Error(`Unable to load the bundled PDF font (${response.status})`)
  return new Uint8Array(await response.arrayBuffer())
}

let cachedFonts: Promise<PdfJsonImportFonts> | undefined

export function loadPdfJsonImportFonts(): Promise<PdfJsonImportFonts> {
  cachedFonts ??= Promise.all([
    loadFont(regularUrl),
    loadFont(boldUrl),
    loadFont(italicUrl),
    loadFont(boldItalicUrl),
    loadFont(unicodeUrl),
  ]).then(([regular, bold, italic, boldItalic, unicode]) => ({
    regular,
    bold,
    italic,
    boldItalic,
    unicode,
  }))
  return cachedFonts
}
