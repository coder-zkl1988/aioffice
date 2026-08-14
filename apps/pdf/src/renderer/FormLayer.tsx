import { Fragment, useEffect, useState } from 'react'
import type { ReactElement } from 'react'
import type { PDFDocumentProxy } from 'pdfjs-dist'
import { pdfRectToCss } from './annotations'
import type { PageGeom } from './annotations'
import type { FormValueInput } from '../shared/ipc'

export interface FormWidget {
  fieldName: string
  kind: 'text' | 'checkbox' | 'radio' | 'choice' | 'signature'
  rect: [number, number, number, number]
  value: string | string[]
  checked: boolean
  multiLine: boolean
  multiSelect: boolean
  combo: boolean
  editable: boolean
  readOnly: boolean
  /** radio: this button's exportValue */
  buttonValue: string
  /** choice: option list */
  options: { exportValue: string; displayValue: string }[]
}

export interface RawFormAnnotation {
  subtype?: string
  fieldType?: string
  fieldName?: string
  fieldValue?: unknown
  fieldFlags?: number
  exportValue?: string
  buttonValue?: string
  rect?: number[]
  readOnly?: boolean
  checkBox?: boolean
  radioButton?: boolean
  pushButton?: boolean
  multiLine?: boolean
  multiSelect?: boolean
  combo?: boolean
  options?: { exportValue?: unknown; displayValue?: unknown }[]
}

const CHOICE_EDIT_FLAG = 0x0040000

interface PdfSignatureMetadata {
  fieldName?: unknown
}

export function signedPdfFieldNames(
  signatures: readonly unknown[] | null | undefined,
): Set<string> {
  const names = new Set<string>()
  for (const signature of signatures ?? []) {
    if (!signature || typeof signature !== 'object') continue
    const fieldName = (signature as PdfSignatureMetadata).fieldName
    if (typeof fieldName === 'string' && fieldName.length > 0) names.add(fieldName)
  }
  return names
}

function fieldValueStr(v: unknown): string {
  const s = Array.isArray(v) ? v[0] : v
  return typeof s === 'string' ? s : ''
}

export function normalizeChoiceValue(value: unknown, multiSelect: boolean): string | string[] {
  const values = (Array.isArray(value) ? value : [value]).filter(
    (item): item is string => typeof item === 'string',
  )
  return multiSelect ? [...new Set(values)] : (values[0] ?? '')
}

export function formWidgetFromAnnotation(annotation: RawFormAnnotation): FormWidget | null {
  if (
    annotation.subtype !== 'Widget' ||
    !annotation.fieldName ||
    !annotation.rect ||
    annotation.rect.length !== 4
  ) {
    return null
  }
  const base = {
    fieldName: annotation.fieldName,
    rect: annotation.rect as FormWidget['rect'],
    checked: false,
    multiLine: false,
    multiSelect: false,
    combo: false,
    editable: false,
    readOnly: !!annotation.readOnly,
    buttonValue: '',
    options: [] as FormWidget['options'],
  }
  if (annotation.fieldType === 'Tx') {
    return {
      ...base,
      kind: 'text',
      value: fieldValueStr(annotation.fieldValue),
      multiLine: !!annotation.multiLine,
    }
  }
  if (annotation.fieldType === 'Sig') {
    return { ...base, kind: 'signature', value: fieldValueStr(annotation.fieldValue) }
  }
  if (annotation.fieldType === 'Btn' && annotation.checkBox) {
    return {
      ...base,
      kind: 'checkbox',
      value: '',
      checked:
        typeof annotation.fieldValue === 'string' &&
        annotation.fieldValue !== 'Off' &&
        annotation.fieldValue !== '',
    }
  }
  if (annotation.fieldType === 'Btn' && annotation.radioButton) {
    return {
      ...base,
      kind: 'radio',
      value: fieldValueStr(annotation.fieldValue),
      buttonValue: typeof annotation.buttonValue === 'string' ? annotation.buttonValue : '',
    }
  }
  if (annotation.fieldType !== 'Ch') return null
  const multiSelect = !!annotation.multiSelect
  const combo = !!annotation.combo
  return {
    ...base,
    kind: 'choice',
    value: normalizeChoiceValue(annotation.fieldValue, multiSelect),
    multiSelect,
    combo,
    editable: combo && !!((annotation.fieldFlags ?? 0) & CHOICE_EDIT_FLAG),
    options: (annotation.options ?? [])
      .filter(
        (option) =>
          typeof option.exportValue === 'string' || typeof option.displayValue === 'string',
      )
      .map((option) => ({
        exportValue: String(option.exportValue ?? option.displayValue),
        displayValue: String(option.displayValue ?? option.exportValue),
      })),
  }
}

/** AcroForm overlay: text fields/checkboxes absolutely positioned on the page; App aggregates values, ⌘S writes back */
export function FormLayer({
  doc,
  pageNo,
  geom,
  scale,
  readOnly = false,
  edits,
  completedSignatureFields = new Set(),
  signatureLabel = 'Sign',
  onEdit,
  onSignField,
}: {
  doc: PDFDocumentProxy
  pageNo: number
  geom: PageGeom
  scale: number
  readOnly?: boolean
  edits: ReadonlyMap<string, FormValueInput>
  completedSignatureFields?: ReadonlySet<string>
  signatureLabel?: string
  onEdit: (value: FormValueInput) => void
  onSignField?: (field: { fieldName: string; rect: FormWidget['rect'] }) => void
}): ReactElement | null {
  const [widgets, setWidgets] = useState<FormWidget[] | null>(null)

  useEffect(() => {
    let cancelled = false
    void (async () => {
      const [page, signatures] = await Promise.all([
        doc.getPage(pageNo),
        doc.getSignatures().catch(() => null),
      ])
      const annots = (await page.getAnnotations()) as RawFormAnnotation[]
      if (cancelled) return
      const signedFields = signedPdfFieldNames(signatures)
      const out = annots.flatMap((annotation) => {
        const widget = formWidgetFromAnnotation(annotation)
        return widget && !(widget.kind === 'signature' && signedFields.has(widget.fieldName))
          ? [widget]
          : []
      })
      setWidgets(out)
    })()
    return () => {
      cancelled = true
    }
  }, [doc, pageNo])

  if (!widgets || widgets.length === 0) return null

  return (
    <div className="pdf-form-layer">
      {widgets.map((w, i) => {
        const style = pdfRectToCss(geom, w.rect, scale)
        const edit = edits.get(w.fieldName)
        if (w.kind === 'signature') {
          if (w.value || completedSignatureFields.has(w.fieldName)) return null
          return (
            <button
              key={i}
              type="button"
              className="pdf-form-signature"
              style={style}
              title={w.fieldName}
              aria-label={`${signatureLabel}: ${w.fieldName}`}
              disabled={readOnly || w.readOnly || !onSignField}
              onClick={() => onSignField?.({ fieldName: w.fieldName, rect: w.rect })}
            >
              <span aria-hidden="true">+</span>
              <span>{signatureLabel}</span>
            </button>
          )
        }
        if (w.kind === 'checkbox') {
          return (
            <input
              key={i}
              type="checkbox"
              className="pdf-form-checkbox"
              style={style}
              aria-label={w.fieldName}
              title={w.fieldName}
              disabled={readOnly || w.readOnly}
              checked={edit ? !!edit.checked : w.checked}
              onChange={(e) =>
                onEdit({ name: w.fieldName, kind: 'checkbox', checked: e.target.checked })
              }
            />
          )
        }
        const editedValue = edit?.value
        const scalarValue = Array.isArray(editedValue)
          ? (editedValue[0] ?? '')
          : (editedValue ?? fieldValueStr(w.value))
        if (w.kind === 'radio') {
          return (
            <input
              key={i}
              type="radio"
              className="pdf-form-checkbox"
              style={style}
              aria-label={w.fieldName}
              title={w.fieldName}
              disabled={readOnly || w.readOnly}
              checked={scalarValue === w.buttonValue}
              onChange={() => onEdit({ name: w.fieldName, kind: 'radio', value: w.buttonValue })}
            />
          )
        }
        if (w.kind === 'choice') {
          const value = edit ? normalizeChoiceValue(edit.value, w.multiSelect) : w.value
          const selectedValues = Array.isArray(value) ? value : value ? [value] : []
          const unknownValues = selectedValues.filter(
            (selected) => !w.options.some((option) => option.exportValue === selected),
          )
          const fontSize = Math.max(9, Math.min(14, style.height * 0.55))
          if (w.editable) {
            const listId = `pdf-form-choice-${pageNo}-${i}`
            return (
              <Fragment key={i}>
                <input
                  className="pdf-form-input"
                  style={{ ...style, fontSize }}
                  aria-label={w.fieldName}
                  title={w.fieldName}
                  list={listId}
                  disabled={readOnly || w.readOnly}
                  value={selectedValues[0] ?? ''}
                  onChange={(event) =>
                    onEdit({ name: w.fieldName, kind: 'choice', value: event.target.value })
                  }
                />
                <datalist id={listId}>
                  {w.options.map((option, optionIndex) => (
                    <option
                      key={optionIndex}
                      value={option.exportValue}
                      label={option.displayValue}
                    />
                  ))}
                </datalist>
              </Fragment>
            )
          }
          return (
            <select
              key={i}
              className="pdf-form-select"
              style={{ ...style, fontSize }}
              aria-label={w.fieldName}
              title={w.fieldName}
              disabled={readOnly || w.readOnly}
              multiple={w.multiSelect}
              size={w.combo ? undefined : Math.min(Math.max(w.options.length, 2), 6)}
              value={w.multiSelect ? selectedValues : (selectedValues[0] ?? '')}
              onChange={(event) =>
                onEdit({
                  name: w.fieldName,
                  kind: 'choice',
                  value: w.multiSelect
                    ? Array.from(event.target.selectedOptions, (option) => option.value)
                    : event.target.value,
                })
              }
            >
              {unknownValues.map((unknown) => (
                <option key={`unknown:${unknown}`} value={unknown} hidden />
              ))}
              {w.options.map((o, j) => (
                <option key={j} value={o.exportValue}>
                  {o.displayValue}
                </option>
              ))}
            </select>
          )
        }
        const fontSize = Math.max(9, Math.min(14, style.height * 0.55))
        if (w.multiLine) {
          return (
            <textarea
              key={i}
              className="pdf-form-input"
              style={{ ...style, fontSize }}
              aria-label={w.fieldName}
              title={w.fieldName}
              disabled={readOnly || w.readOnly}
              value={scalarValue}
              onChange={(e) => onEdit({ name: w.fieldName, kind: 'text', value: e.target.value })}
            />
          )
        }
        return (
          <input
            key={i}
            type="text"
            className="pdf-form-input"
            style={{ ...style, fontSize }}
            aria-label={w.fieldName}
            title={w.fieldName}
            disabled={readOnly || w.readOnly}
            value={scalarValue}
            onChange={(e) => onEdit({ name: w.fieldName, kind: 'text', value: e.target.value })}
          />
        )
      })}
    </div>
  )
}
