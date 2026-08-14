import { describe, expect, it } from 'vitest'
import { aiCreatedPdfHtml, normalizeAiCreatedPdfDocument } from '../src/renderer/ai-created-pdf'

describe('AI-created PDF documents', () => {
  it('normalizes every supported structured section', () => {
    const document = normalizeAiCreatedPdfDocument({
      title: 'Project brief',
      subtitle: 'Internal review',
      reference_number: 'BR-2026-08',
      file_name: 'project-brief.pdf',
      primary_color: '#0f766e',
      sections: [
        { type: 'text', heading: 'Summary', body: 'First line\nSecond line' },
        {
          type: 'key_value',
          pairs: [
            { label: 'Owner', value: 'Operations' },
            { label: 'Status', value: 'Draft' },
          ],
        },
        {
          type: 'line_items',
          columns: ['Item', 'Amount'],
          rows: [
            ['Design', '$800'],
            ['Build', '$1,200'],
          ],
          total_row: ['Total', '$2,000'],
        },
        { type: 'bullet_list', items: ['Confirm scope', 'Approve budget'] },
        { type: 'signature', signatories: ['Prepared by', 'Approved by'] },
      ],
    })

    expect(document.primaryColor).toBe('#0F766E')
    expect(document.sections.map((section) => section.type)).toEqual([
      'text',
      'key_value',
      'line_items',
      'bullet_list',
      'signature',
    ])
    expect(document.sections[2]?.totalRow).toEqual(['Total', '$2,000'])
  })

  it('allows intentionally empty cells in table rows and totals', () => {
    const document = normalizeAiCreatedPdfDocument({
      title: 'Schedule',
      sections: [
        {
          type: 'line_items',
          columns: ['Phase', 'Owner', 'Notes'],
          rows: [['Review', 'Operations', '']],
          total_row: ['Total', 'One week', ''],
        },
      ],
    })

    expect(document.sections[0]?.rows?.[0]).toEqual(['Review', 'Operations', ''])
    expect(document.sections[0]?.totalRow).toEqual(['Total', 'One week', ''])
  })

  it('escapes all model-provided content instead of rendering arbitrary HTML', () => {
    const document = normalizeAiCreatedPdfDocument({
      title: '<img src=x onerror=alert(1)>',
      sections: [{ type: 'text', body: '<script>alert("x")</script> & approved' }],
    })
    const html = aiCreatedPdfHtml(document)

    expect(html).toContain('&lt;img src=x onerror=alert(1)&gt;')
    expect(html).toContain('&lt;script&gt;alert(&quot;x&quot;)&lt;/script&gt; &amp; approved')
    expect(html).not.toContain('<script>alert')
    expect(html).not.toContain('<img src=x')
  })

  it('rejects invalid colors and inconsistent table rows', () => {
    expect(() =>
      normalizeAiCreatedPdfDocument({
        title: 'Bad color',
        primary_color: 'blue',
        sections: [{ type: 'text', body: 'Content' }],
      }),
    ).toThrow('primary_color')

    expect(() =>
      normalizeAiCreatedPdfDocument({
        title: 'Bad table',
        sections: [{ type: 'line_items', columns: ['A', 'B'], rows: [['one cell only']] }],
      }),
    ).toThrow('must match columns')
  })
})
