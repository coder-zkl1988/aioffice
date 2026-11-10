// @vitest-environment node
/**
 * CSV export of the comments panel: RFC 4180 escaping, the spreadsheet
 * formula-injection guard, one row per reply, and the empty-text rules.
 */
import { describe, expect, it } from 'vitest'
import {
  COMMENTS_CSV_HEADERS,
  commentsCsvFileName,
  commentsToCsv,
  toCsvCell,
  type CommentsCsvNote,
} from '../src/renderer/comments-csv'

const BOM = '﻿'

const note = (over: Partial<CommentsCsvNote> = {}): CommentsCsvNote => ({
  author: 'Ann',
  contents: '',
  timeMs: null,
  replies: [],
  ...over,
})

const entry = (pageNo: number, root: CommentsCsvNote) => ({ pageNo, root })

describe('toCsvCell', () => {
  it('leaves plain text unquoted', () => {
    expect(toCsvCell('hello world')).toBe('hello world')
  })

  it('quotes fields with commas, quotes or newlines and doubles inner quotes', () => {
    expect(toCsvCell('a,b')).toBe('"a,b"')
    expect(toCsvCell('say "hi"')).toBe('"say ""hi"""')
    expect(toCsvCell('line1\nline2')).toBe('"line1\nline2"')
    expect(toCsvCell('line1\r\nline2')).toBe('"line1\r\nline2"')
  })

  it.each(['=SUM(A1)', '+1', '-1', '@cmd', '\tx'])(
    'neutralises a leading formula character in %j',
    (value) => {
      expect(toCsvCell(value)).toBe(`'${value}`)
    },
  )

  it('neutralises and quotes a leading carriage return', () => {
    expect(toCsvCell('\r=1')).toBe('"\'\r=1"')
  })

  it('neutralises a formula hidden behind leading spaces or tabs', () => {
    expect(toCsvCell(' =1+1')).toBe("' =1+1")
    expect(toCsvCell('\t@cmd')).toBe("'\t@cmd")
  })

  it('neutralises a formula that also needs quoting', () => {
    expect(toCsvCell('=HYPERLINK("x","y")')).toBe('"\'=HYPERLINK(""x"",""y"")"')
  })
})

describe('commentsToCsv', () => {
  it('writes a BOM and header, then each comment followed by its replies', () => {
    const csv = commentsToCsv([
      entry(
        1,
        note({
          author: 'Ann',
          contents: 'Check this',
          timeMs: Date.UTC(2026, 2, 11, 18, 5),
          replies: [note({ author: 'Bob', contents: 'Done, thanks' })],
        }),
      ),
      entry(3, note({ author: 'Bob', contents: 'Typo' })),
    ])
    expect(csv.startsWith(BOM)).toBe(true)
    expect(csv.slice(BOM.length).split('\r\n')).toEqual([
      'Page,Thread,Author,Last modified (UTC),Comment,In reply to',
      '1,1,Ann,2026-03-11T18:05:00.000Z,Check this,',
      '1,1,Bob,,"Done, thanks",Ann',
      '3,2,Bob,,Typo,',
    ])
  })

  it('writes only the header when there are no comments', () => {
    expect(commentsToCsv([])).toBe(
      `${BOM}Page,Thread,Author,Last modified (UTC),Comment,In reply to`,
    )
  })

  it('orders threads by display page, not by input order', () => {
    const csv = commentsToCsv([
      entry(10, note({ contents: 'Late' })),
      entry(2, note({ contents: 'Early' })),
    ])
    expect(csv.slice(BOM.length).split('\r\n').slice(1)).toEqual([
      '2,1,Ann,,Early,',
      '10,2,Ann,,Late,',
    ])
  })

  it('leaves out comments that have not been posted', () => {
    const csv = commentsToCsv([
      entry(1, note({ contents: '   ', replies: [note({ contents: '  ' })] })),
      entry(1, note({ contents: 'Kept' })),
    ])
    expect(csv.slice(BOM.length).split('\r\n').slice(1)).toEqual(['1,1,Ann,,Kept,'])
  })

  it("falls back to the first reply's text for the root row", () => {
    const csv = commentsToCsv([
      entry(1, note({ author: 'Ann', replies: [note({ author: 'Bob', contents: 'Reply' })] })),
    ])
    expect(csv.slice(BOM.length).split('\r\n').slice(1)).toEqual([
      '1,1,Ann,,Reply,',
      '1,1,Bob,,Reply,Ann',
    ])
  })

  it('writes one row per nested reply and points each at its parent author', () => {
    const csv = commentsToCsv([
      entry(
        2,
        note({
          author: 'Ann',
          contents: 'Root',
          replies: [
            note({
              author: 'Bob',
              contents: 'First',
              replies: [
                note({ author: 'Cid', contents: 'Nested' }),
                note({ author: '', contents: '' }),
              ],
            }),
            note({ author: 'Dee', contents: 'Second' }),
          ],
        }),
      ),
    ])
    expect(csv.slice(BOM.length).split('\r\n').slice(1)).toEqual([
      '2,1,Ann,,Root,',
      '2,1,Bob,,First,Ann',
      '2,1,Cid,,Nested,Bob',
      '2,1,,,,Bob',
      '2,1,Dee,,Second,Ann',
    ])
  })

  it('escapes text and authors that contain separators or formulas', () => {
    const csv = commentsToCsv([
      entry(
        1,
        note({
          author: '=Ann',
          contents: '=SUM(A1), "quoted"',
          replies: [note({ author: 'Bob', contents: 'line1\nline2' })],
        }),
      ),
    ])
    expect(csv.slice(BOM.length).split('\r\n').slice(1)).toEqual([
      `1,1,'=Ann,,"'=SUM(A1), ""quoted""",`,
      '1,1,Bob,,"line1\nline2",\'=Ann',
    ])
  })

  it('trims authors like the sidebar does', () => {
    const csv = commentsToCsv([
      entry(
        1,
        note({
          author: ' Guest ',
          contents: 'Hi',
          replies: [note({ author: '', contents: 'Yo' })],
        }),
      ),
    ])
    expect(csv.slice(BOM.length).split('\r\n').slice(1)).toEqual([
      '1,1,Guest,,Hi,',
      '1,1,,,Yo,Guest',
    ])
  })

  it('accepts custom headers', () => {
    const csv = commentsToCsv([entry(1, note({ contents: 'Hi' }))], {
      page: 'Seite',
      thread: 'Thread',
      author: 'Autor',
      modified: 'Geändert',
      comment: 'Kommentar',
      inReplyTo: 'Antwort auf',
    })
    expect(csv.slice(BOM.length).split('\r\n')[0]).toBe(
      'Seite,Thread,Autor,Geändert,Kommentar,Antwort auf',
    )
  })

  it('uses the default headers', () => {
    expect(COMMENTS_CSV_HEADERS.page).toBe('Page')
    expect(COMMENTS_CSV_HEADERS.inReplyTo).toBe('In reply to')
  })
})

describe('commentsCsvFileName', () => {
  it('replaces the pdf extension', () => {
    expect(commentsCsvFileName('Report.PDF')).toBe('Report-comments.csv')
  })

  it('keeps names without a pdf extension', () => {
    expect(commentsCsvFileName('scan.v2')).toBe('scan.v2-comments.csv')
  })

  it('falls back when the name is empty', () => {
    expect(commentsCsvFileName('')).toBe('document-comments.csv')
  })
})
