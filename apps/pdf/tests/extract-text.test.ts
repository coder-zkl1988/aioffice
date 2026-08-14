import { describe, expect, it, vi } from 'vitest'
import type { PDFDocumentProxy } from 'pdfjs-dist'
import {
  extractPdfTextPageData,
  extractPdfJsonPages,
  extractPdfTextPages,
  groupPdfTextRuns,
  pdfJsonTextRuns,
  pdfTextLinesToBlocks,
} from '../src/renderer/extract-text'

function item(text: string, x: number, y: number, size = 12, width = text.length * 7) {
  return { str: text, transform: [size, 0, 0, size, x, y], width, height: size, fontName: 'body' }
}

describe('PDF text extraction', () => {
  it('preserves text layout and font traits for structured export', () => {
    expect(
      pdfJsonTextRuns(
        [
          {
            str: 'Layout',
            transform: [16, 0, 0, 16, 40, 700],
            width: 54,
            height: 16,
            fontName: 'HeadingBoldItalic',
          },
        ],
        { HeadingBoldItalic: { fontFamily: 'Inter' } },
      ),
    ).toEqual([
      {
        text: 'Layout',
        x: 40,
        y: 700,
        width: 54,
        height: 16,
        fontSize: 16,
        fontFamily: 'Inter',
        bold: true,
        italic: true,
      },
    ])
  })

  it('groups text runs with natural English and CJK spacing', () => {
    const lines = groupPdfTextRuns([
      { text: 'Hello', x: 10, y: 100, width: 30, height: 12, bold: false },
      { text: 'world', x: 45, y: 100, width: 32, height: 12, bold: false },
      { text: '山', x: 10, y: 80, width: 12, height: 12, bold: false },
      { text: '东', x: 22, y: 80, width: 12, height: 12, bold: false },
    ])
    expect(lines.map((line) => line.text)).toEqual(['Hello world', '山东'])
  })

  it('detects headings, joins wrapped lines, repairs hyphenation, and keeps bullets', () => {
    const page = extractPdfTextPageData(
      1,
      [
        item('Launch Plan', 50, 740, 24, 130),
        item('Local genera-', 50, 700, 12, 82),
        item('tion keeps data private.', 50, 684, 12, 132),
        item('• Editable elements', 50, 650, 12, 110),
      ],
      { body: { fontFamily: 'Arial' } },
      [],
    )
    expect(page.blocks).toEqual([
      { kind: 'heading', text: 'Launch Plan', level: 1 },
      { kind: 'paragraph', text: 'Local generation keeps data private.' },
      { kind: 'listItem', text: '• Editable elements' },
    ])
  })

  it('orders genuine two-column prose by column', () => {
    const lines = [
      ...[700, 680, 660, 640].map((y, index) => ({
        text: `Left ${index + 1}.`,
        x: 50,
        y,
        width: 110,
        height: 12,
        bold: false,
      })),
      ...[700, 680, 660, 640].map((y, index) => ({
        text: `Right ${index + 1}.`,
        x: 340,
        y,
        width: 110,
        height: 12,
        bold: false,
      })),
    ]
    expect(pdfTextLinesToBlocks(lines).map((block) => block.text)).toEqual([
      'Left 1. Left 2. Left 3. Left 4.',
      'Right 1. Right 2. Right 3. Right 4.',
    ])
  })

  it('associates visible link text and deduplicates external URLs', () => {
    const page = extractPdfTextPageData(
      2,
      [item('Open docs', 50, 700, 12, 65)],
      { body: { fontFamily: 'Arial' } },
      [
        { subtype: 'Link', rect: [48, 698, 118, 715], url: 'https://example.com/docs' },
        { subtype: 'Link', rect: [48, 698, 118, 715], url: 'https://example.com/docs' },
        { subtype: 'Link', rect: [48, 698, 118, 715] },
      ],
    )
    expect(page.links).toEqual([{ url: 'https://example.com/docs', label: 'Open docs' }])
  })

  it('extracts only requested pages and cleans page resources', async () => {
    const cleanup = vi.fn()
    const document = {
      numPages: 2,
      getPage: vi.fn(async (pageNumber: number) => ({
        getTextContent: async () => ({
          items: [item(`Page ${pageNumber}`, 50, 700)],
          styles: { body: { fontFamily: 'Arial' } },
        }),
        getAnnotations: async () => [],
        cleanup,
      })),
    } as unknown as PDFDocumentProxy
    const pages = await extractPdfTextPages(document, [1])
    expect(document.getPage).toHaveBeenCalledWith(2)
    expect(pages).toMatchObject([{ pageNumber: 2, text: 'Page 2' }])
    expect(cleanup).toHaveBeenCalledOnce()
    await expect(extractPdfTextPages(document, [2])).rejects.toThrow('invalid page')
  })

  it('exports page geometry and only includes layout runs in full mode', async () => {
    const cleanup = vi.fn()
    const document = {
      numPages: 1,
      getPage: vi.fn(async () => ({
        getViewport: () => ({ width: 792, height: 612, rotation: 90 }),
        getTextContent: async () => ({
          items: [item('Structured', 50, 700)],
          styles: { body: { fontFamily: 'Arial' } },
        }),
        getAnnotations: async () => [],
        cleanup,
      })),
    } as unknown as PDFDocumentProxy
    const full = await extractPdfJsonPages(document, [0], false)
    expect(full).toMatchObject([
      { pageNumber: 1, width: 792, height: 612, rotation: 90, text: 'Structured' },
    ])
    expect(full[0]!.textRuns).toHaveLength(1)
    const semantic = await extractPdfJsonPages(document, [0], true)
    expect(semantic[0]).not.toHaveProperty('textRuns')
    expect(cleanup).toHaveBeenCalledTimes(2)
  })

  it('cleans page resources when annotation extraction fails', async () => {
    const cleanup = vi.fn()
    const document = {
      numPages: 1,
      getPage: async () => ({
        getTextContent: async () => ({ items: [], styles: {} }),
        getAnnotations: async () => {
          throw new Error('annotation failure')
        },
        cleanup,
      }),
    } as unknown as PDFDocumentProxy
    await expect(extractPdfTextPages(document, [0])).rejects.toThrow('annotation failure')
    expect(cleanup).toHaveBeenCalledOnce()
  })
})
