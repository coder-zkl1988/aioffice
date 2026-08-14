import { describe, expect, it } from 'vitest'
import {
  formWidgetFromAnnotation,
  normalizeChoiceValue,
  signedPdfFieldNames,
} from '../src/renderer/FormLayer'

describe('PDF form layer choice fields', () => {
  it('keeps every selected value for multi-select fields', () => {
    expect(normalizeChoiceValue(['Design', 'Research', 'Design'], true)).toEqual([
      'Design',
      'Research',
    ])
    expect(normalizeChoiceValue(['Design', 'Research'], false)).toBe('Design')
  })

  it('recognizes editable combo boxes from PDF field flags', () => {
    const widget = formWidgetFromAnnotation({
      subtype: 'Widget',
      fieldType: 'Ch',
      fieldName: 'department',
      fieldValue: ['Customer Success'],
      fieldFlags: 0x0040000,
      combo: true,
      rect: [20, 30, 180, 54],
      options: [
        { exportValue: 'Sales', displayValue: 'Sales team' },
        { exportValue: 'Support', displayValue: 'Support team' },
      ],
    })

    expect(widget).toMatchObject({
      kind: 'choice',
      value: 'Customer Success',
      combo: true,
      editable: true,
    })
  })

  it('maps PDF.js multi-select annotations without collapsing their values', () => {
    const widget = formWidgetFromAnnotation({
      subtype: 'Widget',
      fieldType: 'Ch',
      fieldName: 'topics',
      fieldValue: ['Design', 'Research'],
      multiSelect: true,
      rect: [20, 80, 180, 160],
      options: [
        { exportValue: 'Design', displayValue: 'Design' },
        { exportValue: 'Engineering', displayValue: 'Engineering' },
        { exportValue: 'Research', displayValue: 'Research' },
      ],
    })

    expect(widget).toMatchObject({
      kind: 'choice',
      value: ['Design', 'Research'],
      multiSelect: true,
      combo: false,
      editable: false,
    })
  })

  it('recognizes AcroForm signature widgets', () => {
    expect(
      formWidgetFromAnnotation({
        subtype: 'Widget',
        fieldType: 'Sig',
        fieldName: 'approval.signature',
        rect: [40, 50, 200, 110],
      }),
    ).toMatchObject({
      kind: 'signature',
      fieldName: 'approval.signature',
      value: '',
    })
  })

  it('collects only named digital signature fields from PDF.js metadata', () => {
    expect([
      ...signedPdfFieldNames([{ fieldName: 'approval.signature' }, { fieldName: '' }, null]),
    ]).toEqual(['approval.signature'])
  })
})
