/**
 * Compression feedback: the review step tells the user when a compress run
 * could not make the file smaller (Stirling-PDF #8482).
 */
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, beforeAll, describe, expect, it } from 'vitest'
import { compressionShrank } from '../src/renderer/compress-feedback'
import { PdfToolsDialog } from '../src/renderer/PdfToolsDialog'
import type { TFunc } from '../src/renderer/i18n/locale'

const t = ((key: string) => key) as unknown as TFunc

let root: Root | null = null
let container: HTMLDivElement | null = null

beforeAll(() => {
  ;(globalThis as Record<string, unknown>).IS_REACT_ACT_ENVIRONMENT = true
})

afterEach(async () => {
  if (root) await act(async () => root?.unmount())
  container?.remove()
  root = null
  container = null
})

async function renderCompressDialog(compressNoShrink: boolean) {
  container = document.createElement('div')
  document.body.appendChild(container)
  root = createRoot(container)
  await act(async () => {
    root!.render(
      createElement(PdfToolsDialog, {
        initialKind: 'compress',
        filePath: '/tmp/example.pdf',
        lang: 'en',
        currentPage: 0,
        pageCount: 1,
        t,
        busy: false,
        encrypted: false,
        openingPassword: '',
        ocrProgress: null,
        videoProgress: null,
        compressNoShrink,
        onCancel: () => {},
        onApply: () => {},
        onBulkReplaceText: () => {},
      }),
    )
  })
  return container
}

describe('compressionShrank', () => {
  it('treats a smaller output as a real saving', () => {
    expect(compressionShrank(2_000_000, 800_000)).toBe(true)
  })

  it('treats an unchanged size as no saving', () => {
    expect(compressionShrank(1_000_000, 1_000_000)).toBe(false)
  })

  it('treats a larger output as no saving', () => {
    expect(compressionShrank(1_000_000, 1_200_000)).toBe(false)
  })
})

describe('PdfToolsDialog compress review', () => {
  it('shows the no-shrink notice when the flag is set', async () => {
    await renderCompressDialog(true)
    // The dialog renders through a portal on document.body
    expect(document.body.textContent).toContain(
      'Compression ran successfully, but this file cannot be compressed any further.',
    )
    expect(document.body.querySelector('.pdf-preflight-status.warning')).not.toBeNull()
  })

  it('keeps the compress settings free of the notice otherwise', async () => {
    await renderCompressDialog(false)
    expect(document.body.textContent).not.toContain('cannot be compressed any further')
    expect(document.body.querySelector('.pdf-preflight-status.warning')).toBeNull()
  })
})
