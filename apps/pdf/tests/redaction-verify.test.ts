import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { PDFDocument, StandardFonts } from 'pdf-lib'
import { redactPdf, verifyRedactionRegions } from '../src/main/redaction'
import { readPdfText } from '../src/main/read-text'
import { savePdfToPath } from '../src/main/save-pdf'
import type { SavePdfRequest } from '../src/shared/ipc'

async function twoPageFixture(): Promise<Uint8Array> {
  const doc = await PDFDocument.create()
  const font = await doc.embedFont(StandardFonts.Helvetica)
  const first = doc.addPage([300, 200])
  first.drawText('PUBLIC NOTES', { x: 20, y: 130, size: 18, font })
  const second = doc.addPage([300, 200])
  second.drawText('SECRET BUDGET', { x: 20, y: 130, size: 18, font })
  return doc.save({ useObjectStreams: false })
}

describe('verifyRedactionRegions', () => {
  it('passes once the native redaction has removed the covered text', async () => {
    const before = await twoPageFixture()
    const region = { pageIndex: 1, rect: [18, 125, 200, 155] }
    const out = await redactPdf(before, [region])
    expect((await readPdfText(out)).pages[1]?.text).not.toContain('SECRET')
    await expect(verifyRedactionRegions(out, [region])).resolves.toEqual([])
  })

  it('fails with the one-based page when the text survived', async () => {
    const before = await twoPageFixture()
    // The unredacted bytes are exactly the "baking was skipped" case.
    const failures = await verifyRedactionRegions(before, [
      { pageIndex: 1, rect: [18, 125, 200, 155] },
    ])
    expect(failures).toHaveLength(1)
    expect(failures[0]?.pageIndex).toBe(1)
    expect(failures[0]?.reason).toMatch(/text still overlaps/i)
    // The check must not echo the redacted string back to the caller.
    expect(JSON.stringify(failures)).not.toContain('SECRET')
  })

  it('lists every page whose area still contains text, sorted', async () => {
    const before = await twoPageFixture()
    const failures = await verifyRedactionRegions(before, [
      { pageIndex: 0, rect: [18, 125, 200, 155] },
      { pageIndex: 1, rect: [18, 125, 200, 155] },
    ])
    expect(failures.map((failure) => failure.pageIndex)).toEqual([0, 1])
  })

  it('accepts an area that never contained text', async () => {
    const before = await twoPageFixture()
    await expect(
      verifyRedactionRegions(before, [{ pageIndex: 0, rect: [220, 20, 280, 60] }]),
    ).resolves.toEqual([])
  })

  it('reports a redacted page that is missing from the output', async () => {
    const before = await twoPageFixture()
    await expect(
      verifyRedactionRegions(before, [{ pageIndex: 6, rect: [18, 125, 200, 155] }]),
    ).resolves.toEqual([{ pageIndex: 6, reason: 'redacted page is missing from the saved output' }])
  })

  it('rejects an empty request instead of silently passing', async () => {
    const before = await twoPageFixture()
    await expect(verifyRedactionRegions(before, [])).rejects.toThrow(/at least one/i)
  })
})

describe('the save pipeline withholds an unverified redaction copy', () => {
  it('fails with the page-prefixed error and writes no file', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'redact-verify-'))
    try {
      const source = join(dir, 'source.pdf')
      const target = join(dir, 'copy.pdf')
      const before = await twoPageFixture()
      await writeFile(source, before)
      const request: SavePdfRequest = {
        path: source,
        markups: [],
        drawings: [],
        formValues: [],
        stamps: [],
        // This rectangle stops mid-glyph: PDFium leaves a text fragment inside it,
        // so the copy must be withheld rather than presented as redacted.
        redactions: [{ pageIndex: 1, rect: [20, 130, 160, 144] }],
      }
      await expect(savePdfToPath(source, target, request)).rejects.toThrow(
        /redact-verify-failed pages=2/,
      )
      await expect(readFile(target)).rejects.toThrow()
      // The source itself was never touched.
      expect(await readFile(source)).toEqual(Buffer.from(before))
    } finally {
      await rm(dir, { recursive: true, force: true })
    }
  })
})
