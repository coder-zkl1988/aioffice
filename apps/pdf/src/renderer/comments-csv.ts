/**
 * CSV export of the all-document comments list. One row per comment and per
 * reply, RFC 4180 escaped, with a spreadsheet formula-injection guard.
 * Ported from Stirling-PDF's commentsCsv.ts (PR #8487) and adapted to our
 * NoteThreadItem tree; dates come from the thread's single `timeMs` stamp.
 */

/** Structural view of a note thread item — NoteThreadItem satisfies it */
export interface CommentsCsvNote {
  author: string
  contents: string
  timeMs: number | null
  replies: CommentsCsvNote[]
}

/** Structural view of a comments-panel entry — CommentEntry satisfies it */
export interface CommentsCsvEntry {
  pageNo: number
  root: CommentsCsvNote
}

export interface CommentsCsvHeaders {
  page: string
  thread: string
  author: string
  modified: string
  comment: string
  inReplyTo: string
}

/** Column labels; a CSV is plain data, so these stay untranslated */
export const COMMENTS_CSV_HEADERS: CommentsCsvHeaders = {
  page: 'Page',
  thread: 'Thread',
  author: 'Author',
  modified: 'Last modified (UTC)',
  comment: 'Comment',
  inReplyTo: 'In reply to',
}

// Without it Excel reads the CSV in the system codepage and mangles non-ASCII text.
const EXCEL_UTF8_BOM = '﻿'

const FORMULA_TRIGGERS = new Set(['=', '+', '-', '@', '\t', '\r'])

/**
 * One RFC 4180 field. A leading formula character is prefixed with a quote so
 * spreadsheets show the text instead of evaluating it (CSV injection); leading
 * spaces and tabs are inspected as well, since spreadsheets skip them.
 */
export function toCsvCell(value: string): string {
  const stripped = value.replace(/^[ \t]+/, '')
  const safe =
    FORMULA_TRIGGERS.has(value.charAt(0)) || FORMULA_TRIGGERS.has(stripped.charAt(0))
      ? `'${value}`
      : value
  return /[",\r\n]/.test(safe) ? `"${safe.replace(/"/g, '""')}"` : safe
}

/** ISO 8601 UTC, or an empty cell when the note carries no usable date */
function isoDate(timeMs: number | null): string {
  if (timeMs === null || !Number.isFinite(timeMs)) return ''
  const d = new Date(timeMs)
  return Number.isNaN(d.getTime()) ? '' : d.toISOString()
}

/**
 * The text a thread shows: the root's own contents, else its first reply's,
 * else "" for a comment that has not been posted yet.
 */
function displayText(note: CommentsCsvNote): string {
  const own = (note.contents ?? '').trim()
  if (own) return own
  return (note.replies[0]?.contents ?? '').trim()
}

interface ReplyRow {
  note: CommentsCsvNote
  parentAuthor: string
}

/** Every descendant reply, in thread order; the parent author is the reply-to */
function collectReplies(parent: CommentsCsvNote, out: ReplyRow[]): void {
  for (const reply of parent.replies) {
    out.push({ note: reply, parentAuthor: parent.author })
    collectReplies(reply, out)
  }
}

/**
 * One row per comment and per reply, numbered by thread in page order,
 * prefixed with a UTF-8 BOM and CRLF-terminated. A comment with no text yet is
 * still being composed, so its thread is left out; every reply is its own row.
 */
export function commentsToCsv(
  entries: CommentsCsvEntry[],
  headers: CommentsCsvHeaders = COMMENTS_CSV_HEADERS,
): string {
  const rows: string[][] = [
    [
      headers.page,
      headers.thread,
      headers.author,
      headers.modified,
      headers.comment,
      headers.inReplyTo,
    ],
  ]
  const ordered = [...entries].sort((a, b) => a.pageNo - b.pageNo)
  let threadNo = 0
  for (const entry of ordered) {
    const text = displayText(entry.root)
    if (!text) continue
    threadNo++
    rows.push([
      String(entry.pageNo),
      String(threadNo),
      (entry.root.author ?? '').trim(),
      isoDate(entry.root.timeMs),
      text,
      '',
    ])
    const replies: ReplyRow[] = []
    collectReplies(entry.root, replies)
    for (const { note, parentAuthor } of replies) {
      rows.push([
        String(entry.pageNo),
        String(threadNo),
        (note.author ?? '').trim(),
        isoDate(note.timeMs),
        (note.contents ?? '').trim(),
        (parentAuthor ?? '').trim(),
      ])
    }
  }
  return EXCEL_UTF8_BOM + rows.map((row) => row.map(toCsvCell).join(',')).join('\r\n')
}

/** `<document base>-comments.csv`, dropping a .pdf extension */
export function commentsCsvFileName(pdfName: string): string {
  const base = pdfName.replace(/\.pdf$/i, '') || 'document'
  return `${base}-comments.csv`
}
