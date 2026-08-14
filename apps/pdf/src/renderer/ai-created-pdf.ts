import type { PdfRasterPage } from '@genoffice/pdf-tools'
import { renderLocalHtmlPages } from './local-html-pages'

export type AiCreatedPdfSectionType =
  'text' | 'key_value' | 'line_items' | 'bullet_list' | 'signature'

export interface AiCreatedPdfPair {
  label: string
  value: string
}

export interface AiCreatedPdfSection {
  type: AiCreatedPdfSectionType
  heading?: string
  body?: string
  pairs?: AiCreatedPdfPair[]
  columns?: string[]
  rows?: string[][]
  totalRow?: string[]
  items?: string[]
  signatories?: string[]
}

export interface AiCreatedPdfDocument {
  title: string
  subtitle?: string
  referenceNumber?: string
  fileName?: string
  primaryColor: string
  sections: AiCreatedPdfSection[]
}

export interface PreparedAiCreatedPdf {
  title: string
  baseName: string
  pages: PdfRasterPage[]
}

const MAX_TOTAL_TEXT = 100_000
const MAX_SECTIONS = 40

function record(value: unknown, name: string): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new Error(`${name} must be an object`)
  }
  return value as Record<string, unknown>
}

function text(value: unknown, name: string, maxLength: number, required = false): string {
  if (value === undefined || value === null) {
    if (required) throw new Error(`${name} is required`)
    return ''
  }
  if (typeof value !== 'string') throw new Error(`${name} must be text`)
  const normalized = value.replace(/\r\n?/g, '\n').trim()
  if (required && !normalized) throw new Error(`${name} is required`)
  if (normalized.length > maxLength) throw new Error(`${name} is too long`)
  return normalized
}

function array(value: unknown, name: string, maxLength: number, required = false): unknown[] {
  if (value === undefined || value === null) {
    if (required) throw new Error(`${name} is required`)
    return []
  }
  if (!Array.isArray(value)) throw new Error(`${name} must be an array`)
  if (required && value.length === 0) throw new Error(`${name} is required`)
  if (value.length > maxLength) throw new Error(`${name} has too many entries`)
  return value
}

function textArray(
  value: unknown,
  name: string,
  maxItems: number,
  maxItemLength: number,
  required = false,
  itemRequired = required,
): string[] {
  return array(value, name, maxItems, required).map((item, index) =>
    text(item, `${name}[${index}]`, maxItemLength, itemRequired),
  )
}

function optional(value: string): string | undefined {
  return value || undefined
}

export function normalizeAiCreatedPdfDocument(input: unknown): AiCreatedPdfDocument {
  const source = record(input, 'document')
  const title = text(source.title, 'title', 200, true)
  const rawSections = array(source.sections, 'sections', MAX_SECTIONS, true)
  const sections = rawSections.map((entry, sectionIndex): AiCreatedPdfSection => {
    const section = record(entry, `sections[${sectionIndex}]`)
    const type = text(section.type, `sections[${sectionIndex}].type`, 30, true)
    if (!['text', 'key_value', 'line_items', 'bullet_list', 'signature'].includes(type)) {
      throw new Error(`sections[${sectionIndex}].type is invalid`)
    }
    const heading = optional(text(section.heading, `sections[${sectionIndex}].heading`, 200))
    if (type === 'text') {
      return {
        type,
        heading,
        body: text(section.body, `sections[${sectionIndex}].body`, 20_000, true),
      }
    }
    if (type === 'key_value') {
      const pairs = array(section.pairs, `sections[${sectionIndex}].pairs`, 100, true).map(
        (entry, pairIndex) => {
          const pair = record(entry, `sections[${sectionIndex}].pairs[${pairIndex}]`)
          return {
            label: text(
              pair.label,
              `sections[${sectionIndex}].pairs[${pairIndex}].label`,
              200,
              true,
            ),
            value: text(
              pair.value,
              `sections[${sectionIndex}].pairs[${pairIndex}].value`,
              2_000,
              true,
            ),
          }
        },
      )
      return { type, heading, pairs }
    }
    if (type === 'line_items') {
      const columns = textArray(section.columns, `sections[${sectionIndex}].columns`, 12, 200, true)
      const rows = array(section.rows, `sections[${sectionIndex}].rows`, 300, true).map(
        (row, rowIndex) => {
          const cells = textArray(
            row,
            `sections[${sectionIndex}].rows[${rowIndex}]`,
            columns.length,
            2_000,
            true,
            false,
          )
          if (cells.length !== columns.length) {
            throw new Error(`sections[${sectionIndex}].rows[${rowIndex}] must match columns`)
          }
          return cells
        },
      )
      const totalRow = textArray(
        section.total_row,
        `sections[${sectionIndex}].total_row`,
        columns.length,
        2_000,
        false,
        false,
      )
      if (totalRow.length > 0 && totalRow.length !== columns.length) {
        throw new Error(`sections[${sectionIndex}].total_row must match columns`)
      }
      return { type, heading, columns, rows, totalRow: totalRow.length ? totalRow : undefined }
    }
    if (type === 'bullet_list') {
      return {
        type,
        heading,
        items: textArray(section.items, `sections[${sectionIndex}].items`, 100, 2_000, true),
      }
    }
    return {
      type: 'signature',
      heading,
      signatories: textArray(
        section.signatories,
        `sections[${sectionIndex}].signatories`,
        12,
        200,
        true,
      ),
    }
  })
  const primaryColor = text(source.primary_color, 'primary_color', 7) || '#2563EB'
  if (!/^#[0-9a-f]{6}$/i.test(primaryColor)) {
    throw new Error('primary_color must use #RRGGBB format')
  }
  const document: AiCreatedPdfDocument = {
    title,
    subtitle: optional(text(source.subtitle, 'subtitle', 500)),
    referenceNumber: optional(text(source.reference_number, 'reference_number', 200)),
    fileName: optional(text(source.file_name, 'file_name', 200)),
    primaryColor: primaryColor.toUpperCase(),
    sections,
  }
  const totalText = JSON.stringify(document).length
  if (totalText > MAX_TOTAL_TEXT) throw new Error('Document content is too long')
  return document
}

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;')
}

function headingHtml(heading?: string): string {
  return heading ? `<h2>${escapeHtml(heading)}</h2>` : ''
}

function sectionHtml(section: AiCreatedPdfSection): string {
  if (section.type === 'text') {
    return `<article>${headingHtml(section.heading)}<div class="body-copy">${escapeHtml(section.body ?? '')}</div></article>`
  }
  if (section.type === 'key_value') {
    return `<article>${headingHtml(section.heading)}<dl>${(section.pairs ?? [])
      .map(
        (pair) => `<div><dt>${escapeHtml(pair.label)}</dt><dd>${escapeHtml(pair.value)}</dd></div>`,
      )
      .join('')}</dl></article>`
  }
  if (section.type === 'line_items') {
    const header = (section.columns ?? [])
      .map((column) => `<th>${escapeHtml(column)}</th>`)
      .join('')
    const rows = (section.rows ?? [])
      .map((row) => `<tr>${row.map((cell) => `<td>${escapeHtml(cell)}</td>`).join('')}</tr>`)
      .join('')
    const total = section.totalRow?.length
      ? `<tfoot><tr>${section.totalRow.map((cell) => `<td>${escapeHtml(cell)}</td>`).join('')}</tr></tfoot>`
      : ''
    return `<article>${headingHtml(section.heading)}<table><thead><tr>${header}</tr></thead><tbody>${rows}</tbody>${total}</table></article>`
  }
  if (section.type === 'bullet_list') {
    return `<article>${headingHtml(section.heading)}<ul>${(section.items ?? [])
      .map((item) => `<li>${escapeHtml(item)}</li>`)
      .join('')}</ul></article>`
  }
  return `<article>${headingHtml(section.heading)}<div class="signatures">${(
    section.signatories ?? []
  )
    .map(
      (signatory) =>
        `<div class="signature"><div class="signature-line"></div><div>${escapeHtml(signatory)}</div></div>`,
    )
    .join('')}</div></article>`
}

export function aiCreatedPdfHtml(document: AiCreatedPdfDocument): string {
  const meta = [document.subtitle, document.referenceNumber].filter(Boolean)
  return `<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<style>
:root { --accent: ${document.primaryColor}; --ink: #17191d; --muted: #667085; --line: #dfe3e8; }
* { box-sizing: border-box; }
html, body { margin: 0; width: 794px; background: #fff; color: var(--ink); }
body { min-height: 1123px; padding: 72px 70px 82px; font-family: Aptos, "Segoe UI", "PingFang SC", "Microsoft YaHei", sans-serif; font-size: 15px; line-height: 1.65; letter-spacing: 0; }
header { margin-bottom: 44px; padding-top: 20px; border-top: 7px solid var(--accent); }
h1 { margin: 0; max-width: 620px; font-size: 34px; line-height: 1.22; font-weight: 720; letter-spacing: 0; overflow-wrap: anywhere; }
.meta { display: flex; flex-wrap: wrap; gap: 8px 24px; margin-top: 14px; color: var(--muted); font-size: 13px; }
.meta span + span::before { content: ""; display: inline-block; width: 4px; height: 4px; margin: 0 14px 3px 0; border-radius: 50%; background: var(--accent); }
main { display: grid; gap: 34px; }
article { break-inside: avoid; page-break-inside: avoid; }
h2 { margin: 0 0 14px; padding-left: 13px; border-left: 4px solid var(--accent); font-size: 19px; line-height: 1.35; font-weight: 700; letter-spacing: 0; }
.body-copy { white-space: pre-wrap; overflow-wrap: anywhere; }
dl { margin: 0; border-top: 1px solid var(--line); }
dl > div { display: grid; grid-template-columns: minmax(120px, 0.8fr) minmax(0, 1.8fr); gap: 20px; padding: 11px 2px; border-bottom: 1px solid var(--line); break-inside: avoid; }
dt { color: var(--muted); font-weight: 600; }
dd { margin: 0; white-space: pre-wrap; overflow-wrap: anywhere; }
table { width: 100%; border-collapse: collapse; table-layout: fixed; font-size: 13px; }
th, td { padding: 10px 9px; border: 1px solid var(--line); text-align: left; vertical-align: top; overflow-wrap: anywhere; }
th { background: #f4f6f8; font-weight: 700; }
tfoot td { border-top: 2px solid var(--accent); font-weight: 700; }
tr { break-inside: avoid; page-break-inside: avoid; }
ul { margin: 0; padding-left: 22px; }
li { margin: 7px 0; padding-left: 5px; break-inside: avoid; }
li::marker { color: var(--accent); }
.signatures { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 42px 34px; padding-top: 28px; }
.signature { color: var(--muted); font-size: 13px; }
.signature-line { height: 44px; margin-bottom: 8px; border-bottom: 1px solid #7b8490; }
</style>
</head>
<body>
<header><h1>${escapeHtml(document.title)}</h1>${
    meta.length
      ? `<div class="meta">${meta.map((item) => `<span>${escapeHtml(item!)}</span>`).join('')}</div>`
      : ''
  }</header>
<main>${document.sections.map(sectionHtml).join('')}</main>
</body>
</html>`
}

export async function prepareAiCreatedPdf(
  document: AiCreatedPdfDocument,
): Promise<PreparedAiCreatedPdf> {
  return {
    title: document.title,
    baseName: document.fileName || document.title,
    pages: await renderLocalHtmlPages(aiCreatedPdfHtml(document), {
      includePageNumbers: true,
      maxPages: 100,
      rasterScale: 1.5,
    }),
  }
}
