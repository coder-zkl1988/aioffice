import { describe, expect, it } from 'vitest'
import { auditDocumentInSearchIndex } from '../src/renderer/document-audit'
import type { SearchIndex } from '../src/renderer/search'

function indexFor(...pages: string[]): SearchIndex {
  return pages.map((text) => {
    const items: SearchIndex[number]['items'] = []
    let from = 0
    for (const line of text.split('\n')) {
      if (line) {
        items.push({
          start: from,
          end: from + line.length,
          x: 20,
          y: 700,
          w: line.length * 6,
          h: 12,
        })
      }
      from += line.length + 1
    }
    return { text, lower: text.toLocaleLowerCase(), items }
  })
}

describe('auditDocumentInSearchIndex', () => {
  it('classifies only from the bounded first and last page window', () => {
    const report = auditDocumentInSearchIndex(
      indexFor('Invoice', 'body', 'Meeting Agenda', 'body', 'Annual Report'),
    )
    expect(report.classificationPageIndexes).toEqual([0, 1, 3, 4])
    expect(report.classifications.map(({ id }) => id)).toEqual(['invoice', 'annual-report'])
  })

  it('adds stable parent labels for specialized contracts and invoices', () => {
    const report = auditDocumentInSearchIndex(
      indexFor('Mutual Non-Disclosure Agreement\nProforma Invoice'),
    )
    expect(report.classifications.map(({ id }) => id)).toEqual([
      'proforma-invoice',
      'invoice',
      'nda',
      'contract',
    ])
  })

  it('does not classify from ordinary keyword prose', () => {
    const report = auditDocumentInSearchIndex(
      indexFor('We discussed the invoice, contract, meeting agenda, and annual report.'),
    )
    expect(report.classifications).toEqual([])
  })

  it('detects and masks email plus labeled phone numbers', () => {
    const report = auditDocumentInSearchIndex(
      indexFor('Resume\nEmail: alex.chen@example.com\nPhone: +1 (415) 555-2671'),
    )
    expect(report.classifications[0]?.id).toBe('resume')
    expect(report.sensitivity).toBe('confidential')
    expect(report.sensitiveFindings).toEqual([
      expect.objectContaining({ kind: 'email', maskedValue: 'a***@example.com' }),
      expect.objectContaining({ kind: 'phone', maskedValue: '*******2671' }),
    ])
  })

  it('validates Chinese identity and payment-card checksums', () => {
    const report = auditDocumentInSearchIndex(
      indexFor(
        '身份证号：11010519491231002X\nCard Number: 4111 1111 1111 1111\nCard Number: 4111 1111 1111 1112',
      ),
    )
    expect(report.sensitivity).toBe('restricted')
    expect(report.sensitiveFindings.map(({ kind }) => kind)).toEqual(['cnIdentity', 'paymentCard'])
  })

  it('validates labeled SSN, passport, and IBAN values', () => {
    const report = auditDocumentInSearchIndex(
      indexFor('SSN: 123-45-6789\nPassport No: E12345678\nIBAN: GB82 WEST 1234 5698 7654 32'),
    )
    expect(report.sensitiveFindings.map(({ kind }) => kind)).toEqual(['ssn', 'passport', 'iban'])
    expect(report.sensitiveFindings[0]?.maskedValue).toBe('***-**-6789')
  })

  it('ignores unlabeled phone-like values and invalid identifiers', () => {
    const report = auditDocumentInSearchIndex(
      indexFor('Order 1234567890\nSSN: 000-12-3456\n110105194912310021\n4111111111111112'),
    )
    expect(report.sensitiveFindings).toEqual([])
  })

  it('honors explicit sensitivity markings and repeated occurrences', () => {
    const email = 'team@example.com'
    const report = auditDocumentInSearchIndex(indexFor(`INTERNAL USE ONLY\n${email}\n${email}`))
    expect(report.explicitSensitivityMarkers).toBe(1)
    expect(report.sensitivity).toBe('confidential')
    expect(report.sensitiveFindings.map(({ occurrence }) => occurrence)).toEqual([1, 2])
  })

  it('caps findings and validates options', () => {
    const report = auditDocumentInSearchIndex(indexFor('a@example.com\nb@example.com'), {
      maxFindings: 1,
    })
    expect(report.sensitiveFindings).toHaveLength(1)
    expect(report.truncated).toBe(true)
    expect(() => auditDocumentInSearchIndex([], { maxFindings: 0 })).toThrow('positive integer')
  })
})
