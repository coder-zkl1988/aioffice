import { describe, expect, it } from 'vitest'
import { localHtmlPageBreaks } from '../src/renderer/local-html-pages'

describe('local HTML pagination', () => {
  it('prefers forced chapter starts inside the current page', () => {
    const body = document.createElement('main')
    const first = document.createElement('p')
    const chapter = document.createElement('article')
    chapter.setAttribute('data-pdf-page-break-before', '')
    body.append(first, chapter)
    document.body.appendChild(body)
    Object.defineProperty(body, 'getBoundingClientRect', {
      value: () => ({ top: 0, height: 1800 }),
    })
    Object.defineProperty(first, 'getBoundingClientRect', {
      value: () => ({ top: 0, bottom: 700 }),
    })
    Object.defineProperty(chapter, 'getBoundingClientRect', {
      value: () => ({ top: 900, bottom: 1700 }),
    })

    expect(localHtmlPageBreaks(body, 1800, 1123, 10)).toEqual([0, 900])
    body.remove()
  })

  it('guards against excessive output pages', () => {
    const body = document.createElement('main')
    document.body.appendChild(body)
    Object.defineProperty(body, 'getBoundingClientRect', { value: () => ({ top: 0 }) })
    expect(() => localHtmlPageBreaks(body, 10_000, 1000, 2)).toThrow('no more than 2 pages')
    body.remove()
  })
})
