import { useLayoutEffect, useRef, useState } from 'react'
import type { PointerEvent as ReactPointerEvent, ReactElement } from 'react'
import {
  DEFAULT_HEADER_FOOTER,
  DEFAULT_WATERMARK,
  WATERMARK_MARGIN_FACTORS,
  createWatermarkUuid,
  parseStampPageRange,
  resolveWatermarkText,
} from './stamps'
import type {
  HeaderFooterFont,
  HeaderFooterConfig,
  WatermarkConfig,
  WatermarkMargin,
  WatermarkPosition,
} from './stamps'
import type { MetadataInput } from '../shared/ipc'
import type { TFunc } from './i18n/locale'

const WM_COLORS = ['#d0342c', '#8a8a8a', '#2b66ff', '#217346']
const MAX_WATERMARK_IMAGE_BYTES = 20 * 1024 * 1024
const MAX_WATERMARK_IMAGE_PIXELS = 40_000_000
const MAX_WATERMARK_IMAGE_DIMENSION = 16_384
const MAX_WATERMARK_OUTPUT_DIMENSION = 2_048
const WATERMARK_POSITIONS: WatermarkPosition[] = [
  'topLeft',
  'topCenter',
  'topRight',
  'middleLeft',
  'center',
  'middleRight',
  'bottomLeft',
  'bottomCenter',
  'bottomRight',
]
const WATERMARK_MARGINS: { value: WatermarkMargin; label: string }[] = [
  { value: 'small', label: 'S · 2%' },
  { value: 'medium', label: 'M · 3.5%' },
  { value: 'large', label: 'L · 5%' },
  { value: 'x-large', label: 'XL · 7.5%' },
]
const HEADER_FOOTER_FONTS: { value: HeaderFooterFont; label: string }[] = [
  { value: 'sans', label: 'Sans' },
  { value: 'serif', label: 'Serif' },
  { value: 'mono', label: 'Mono' },
]

async function prepareWatermarkImage(file: File): Promise<{ image: string; aspectRatio: number }> {
  if (file.size > MAX_WATERMARK_IMAGE_BYTES) throw new Error('invalid-image')
  let bitmap: ImageBitmap
  try {
    bitmap = await createImageBitmap(file, { imageOrientation: 'from-image' })
  } catch {
    throw new Error('invalid-image')
  }
  try {
    if (
      bitmap.width <= 0 ||
      bitmap.height <= 0 ||
      bitmap.width > MAX_WATERMARK_IMAGE_DIMENSION ||
      bitmap.height > MAX_WATERMARK_IMAGE_DIMENSION ||
      bitmap.width * bitmap.height > MAX_WATERMARK_IMAGE_PIXELS
    ) {
      throw new Error('invalid-image')
    }
    const scale = Math.min(
      1,
      MAX_WATERMARK_OUTPUT_DIMENSION / Math.max(bitmap.width, bitmap.height),
    )
    const canvas = document.createElement('canvas')
    canvas.width = Math.max(1, Math.round(bitmap.width * scale))
    canvas.height = Math.max(1, Math.round(bitmap.height * scale))
    const context = canvas.getContext('2d')
    if (!context) throw new Error('invalid-image')
    context.drawImage(bitmap, 0, 0, canvas.width, canvas.height)
    const image = canvas.toDataURL('image/png').split(',')[1]
    if (!image) throw new Error('invalid-image')
    return { image, aspectRatio: bitmap.width / bitmap.height }
  } finally {
    bitmap.close()
  }
}

/** Watermark / header-footer config dialog; on confirm App generates stamps and marks unsaved changes */
export function StampDialog({
  fileName,
  totalPages,
  metadata,
  t,
  onCancel,
  onApply,
}: {
  fileName: string
  totalPages: number
  metadata: MetadataInput
  t: TFunc
  onCancel: () => void
  onApply: (
    wm: WatermarkConfig | null,
    hf: HeaderFooterConfig | null,
    watermarkPageNumbers: number[] | null,
    headerFooterPageNumbers: number[] | null,
  ) => void
}): ReactElement {
  const [tab, setTab] = useState<'watermark' | 'hf'>('watermark')
  const [wm, setWm] = useState<WatermarkConfig>(() => ({
    ...DEFAULT_WATERMARK,
    uuid: createWatermarkUuid(),
  }))
  const [hf, setHf] = useState<HeaderFooterConfig>(DEFAULT_HEADER_FOOTER)
  const [hfEnabled, setHfEnabled] = useState(false)
  const [imageName, setImageName] = useState('')
  const [imageBusy, setImageBusy] = useState(false)
  const [imageError, setImageError] = useState(false)
  const [pageRange, setPageRange] = useState(totalPages > 1 ? `1-${totalPages}` : '1')
  const [headerFooterPageRange, setHeaderFooterPageRange] = useState(
    totalPages > 1 ? `1-${totalPages}` : '1',
  )
  const [previewLayout, setPreviewLayout] = useState({ left: 50, top: 50, scale: 1 })
  const imageInputRef = useRef<HTMLInputElement>(null)
  const previewRef = useRef<HTMLDivElement>(null)
  const previewItemRef = useRef<HTMLSpanElement>(null)
  const previewDate = useRef(new Date())
  const previewDrag = useRef<{ pointerId: number; offsetX: number; offsetY: number } | null>(null)

  const hfUsed =
    hfEnabled &&
    (hf.pageNumber ||
      [
        hf.headerLeft,
        hf.headerCenter,
        hf.headerRight,
        hf.footerLeft,
        hf.footerCenter,
        hf.footerRight,
      ].some((s) => s.trim()))
  const wmUsed = wm.type === 'text' ? wm.text.trim().length > 0 : wm.image.length > 0
  const selectedPages = parseStampPageRange(pageRange, totalPages)
  const selectedHeaderFooterPages = parseStampPageRange(headerFooterPageRange, totalPages)
  const canApply =
    (wmUsed || hfUsed) &&
    (!wmUsed || selectedPages !== null) &&
    (!hfUsed || selectedHeaderFooterPages !== null)
  const previewText = resolveWatermarkText(wm.text, {
    pageNumber: selectedPages?.[0] ?? 1,
    totalPages,
    fileName,
    dateTime: previewDate.current,
    ...metadata,
    uuid: wm.uuid,
  })

  useLayoutEffect(() => {
    if (wm.layout !== 'single') return
    const preview = previewRef.current
    const item = previewItemRef.current
    if (!preview || !item) return

    const updatePreviewLayout = () => {
      const width = item.offsetWidth
      const height = item.offsetHeight
      if (!width || !height) return
      const radians = (wm.angle * Math.PI) / 180
      const rotatedWidth =
        Math.abs(width * Math.cos(radians)) + Math.abs(height * Math.sin(radians))
      const rotatedHeight =
        Math.abs(width * Math.sin(radians)) + Math.abs(height * Math.cos(radians))
      const fitPadding = 10
      const scale = Math.min(
        1,
        (preview.clientWidth - fitPadding * 2) / rotatedWidth,
        (preview.clientHeight - fitPadding * 2) / rotatedHeight,
      )
      const boundedScale = Math.max(0.18, scale * 0.96)
      const boundWidth = rotatedWidth * boundedScale
      const boundHeight = rotatedHeight * boundedScale
      const positionMargin = wm.customPosition
        ? 0
        : ((preview.clientWidth + preview.clientHeight) / 2) * WATERMARK_MARGIN_FACTORS[wm.margin]
      const horizontal = wm.position.endsWith('Left')
        ? 'left'
        : wm.position.endsWith('Right')
          ? 'right'
          : 'center'
      const vertical = wm.position.startsWith('top')
        ? 'top'
        : wm.position.startsWith('bottom')
          ? 'bottom'
          : 'middle'
      const left = wm.customPosition
        ? Math.max(
            boundWidth / 2,
            Math.min(
              preview.clientWidth - boundWidth / 2,
              wm.customPosition.xRatio * preview.clientWidth,
            ),
          )
        : horizontal === 'left'
          ? positionMargin + boundWidth / 2
          : horizontal === 'right'
            ? preview.clientWidth - positionMargin - boundWidth / 2
            : preview.clientWidth / 2
      const top = wm.customPosition
        ? Math.max(
            boundHeight / 2,
            Math.min(
              preview.clientHeight - boundHeight / 2,
              (1 - wm.customPosition.yRatio) * preview.clientHeight,
            ),
          )
        : vertical === 'top'
          ? positionMargin + boundHeight / 2
          : vertical === 'bottom'
            ? preview.clientHeight - positionMargin - boundHeight / 2
            : preview.clientHeight / 2
      setPreviewLayout({ left, top, scale: boundedScale })
    }

    updatePreviewLayout()
    const observer = new ResizeObserver(updatePreviewLayout)
    observer.observe(preview)
    return () => observer.disconnect()
  }, [
    previewText,
    wm.angle,
    wm.customPosition,
    wm.image,
    wm.layout,
    wm.margin,
    wm.position,
    wm.sizeRatio,
    wm.type,
  ])

  const startPreviewDrag = (event: ReactPointerEvent<HTMLSpanElement>) => {
    if (wm.layout !== 'single') return
    const preview = previewRef.current
    if (!preview) return
    const rect = preview.getBoundingClientRect()
    previewDrag.current = {
      pointerId: event.pointerId,
      offsetX: event.clientX - rect.left - previewLayout.left,
      offsetY: event.clientY - rect.top - previewLayout.top,
    }
    event.currentTarget.setPointerCapture(event.pointerId)
    event.preventDefault()
  }

  const movePreview = (event: ReactPointerEvent<HTMLSpanElement>) => {
    const drag = previewDrag.current
    const preview = previewRef.current
    if (!drag || drag.pointerId !== event.pointerId || !preview) return
    const rect = preview.getBoundingClientRect()
    const xRatio = Math.max(
      0,
      Math.min(1, (event.clientX - rect.left - drag.offsetX) / Math.max(1, rect.width)),
    )
    const yRatio = Math.max(
      0,
      Math.min(1, 1 - (event.clientY - rect.top - drag.offsetY) / Math.max(1, rect.height)),
    )
    setWm((current) => ({ ...current, customPosition: { xRatio, yRatio } }))
  }

  const stopPreviewDrag = (event: ReactPointerEvent<HTMLSpanElement>) => {
    if (previewDrag.current?.pointerId !== event.pointerId) return
    previewDrag.current = null
    if (event.currentTarget.hasPointerCapture(event.pointerId)) {
      event.currentTarget.releasePointerCapture(event.pointerId)
    }
  }

  const pickImage = async (file: File | undefined) => {
    if (!file) return
    setImageBusy(true)
    setImageError(false)
    try {
      const prepared = await prepareWatermarkImage(file)
      setWm((current) => ({
        ...current,
        type: 'image',
        image: prepared.image,
        imageAspectRatio: prepared.aspectRatio,
      }))
      setImageName(file.name)
    } catch {
      setImageError(true)
    } finally {
      setImageBusy(false)
    }
  }

  const field = (key: keyof HeaderFooterConfig, label: string): ReactElement => (
    <label className="pdf-field">
      <span>{label}</span>
      <input
        className="pdf-modal-input"
        value={String(hf[key])}
        onChange={(e) => setHf({ ...hf, [key]: e.target.value })}
      />
    </label>
  )

  return (
    <div className="pdf-modal-mask" onClick={onCancel}>
      <div className="pdf-modal pdf-modal-wide" onClick={(e) => e.stopPropagation()}>
        <div className="pdf-modal-title">{t('stampTitle')}</div>
        <div className="pdf-sign-tabs">
          <button
            className={`pdf-sign-tab${tab === 'watermark' ? ' active' : ''}`}
            onClick={() => setTab('watermark')}
          >
            {t('watermark')}
          </button>
          <button
            className={`pdf-sign-tab${tab === 'hf' ? ' active' : ''}`}
            onClick={() => {
              setTab('hf')
              setHfEnabled(true)
            }}
          >
            {t('headerFooter')}
          </button>
        </div>

        {tab === 'watermark' ? (
          <div className="pdf-wm-fields">
            <div className="pdf-field pdf-wm-segment-field">
              <span>{t('watermarkContent')}</span>
              <div className="pdf-sign-tabs pdf-wm-segments">
                <button
                  className={`pdf-sign-tab${wm.type === 'text' ? ' active' : ''}`}
                  onClick={() =>
                    setWm({ ...wm, type: 'text', sizeRatio: Math.min(wm.sizeRatio, 0.22) })
                  }
                >
                  {t('watermarkText')}
                </button>
                <button
                  className={`pdf-sign-tab${wm.type === 'image' ? ' active' : ''}`}
                  onClick={() => setWm({ ...wm, type: 'image' })}
                >
                  {t('signImage')}
                </button>
              </div>
            </div>
            <div className="pdf-field pdf-wm-segment-field">
              <span>{t('watermarkLayout')}</span>
              <div className="pdf-sign-tabs pdf-wm-segments">
                <button
                  className={`pdf-sign-tab${wm.layout === 'single' ? ' active' : ''}`}
                  onClick={() => setWm({ ...wm, layout: 'single' })}
                >
                  {t('watermarkSingle')}
                </button>
                <button
                  className={`pdf-sign-tab${wm.layout === 'tiled' ? ' active' : ''}`}
                  onClick={() => setWm({ ...wm, layout: 'tiled' })}
                >
                  {t('watermarkTiled')}
                </button>
              </div>
            </div>
            {wm.type === 'text' ? (
              <label className="pdf-field pdf-tools-field-column">
                <span>{t('watermarkText')}</span>
                <textarea
                  className="pdf-modal-input pdf-wm-textarea"
                  value={wm.text}
                  placeholder={t('watermarkPlaceholder')}
                  autoFocus
                  onChange={(e) => setWm({ ...wm, text: e.target.value })}
                />
              </label>
            ) : (
              <div className="pdf-field pdf-wm-image-field">
                <span>{t('watermarkImage')}</span>
                <input
                  ref={imageInputRef}
                  className="pdf-wm-file-input"
                  type="file"
                  accept="image/png,image/jpeg,image/webp,.png,.jpg,.jpeg,.webp"
                  disabled={imageBusy}
                  onChange={(event) => {
                    void pickImage(event.target.files?.[0])
                    event.target.value = ''
                  }}
                />
                <button
                  className="pdf-modal-btn pdf-wm-file-button"
                  disabled={imageBusy}
                  onClick={() => imageInputRef.current?.click()}
                >
                  {imageBusy
                    ? t('loading')
                    : wm.image
                      ? t('watermarkReplaceImage')
                      : t('watermarkChooseImage')}
                </button>
                {imageName && <em className="pdf-wm-file-name">{imageName}</em>}
              </div>
            )}
            {imageError && (
              <div className="pdf-sign-hint pdf-wm-error" role="alert">
                {t('watermarkImageError')}
              </div>
            )}
            <label className="pdf-field">
              <span>{t('watermarkPages')}</span>
              <input
                className={`pdf-modal-input${selectedPages ? '' : ' invalid'}`}
                value={pageRange}
                aria-invalid={!selectedPages}
                title={t('extractRangeHint', { total: totalPages })}
                placeholder={totalPages > 1 ? `1-${totalPages}` : '1'}
                onChange={(event) => setPageRange(event.target.value)}
              />
            </label>
            <label className="pdf-field">
              <span>{t('watermarkAngle')}</span>
              <input
                type="range"
                min={-180}
                max={180}
                value={wm.angle}
                onChange={(e) => setWm({ ...wm, angle: Number(e.target.value) })}
              />
              <em>{wm.angle}°</em>
            </label>
            <label className="pdf-field">
              <span>{t('watermarkOpacity')}</span>
              <input
                type="range"
                min={5}
                max={100}
                value={Math.round(wm.opacity * 100)}
                onChange={(e) => setWm({ ...wm, opacity: Number(e.target.value) / 100 })}
              />
              <em>{Math.round(wm.opacity * 100)}%</em>
            </label>
            <label className="pdf-field">
              <span>{t('watermarkSize')}</span>
              <input
                type="range"
                min={4}
                max={wm.type === 'image' ? 50 : 22}
                value={Math.round(wm.sizeRatio * 100)}
                onChange={(e) => setWm({ ...wm, sizeRatio: Number(e.target.value) / 100 })}
              />
              <em>{Math.round(wm.sizeRatio * 100)}%</em>
            </label>
            {wm.layout === 'single' && (
              <div className="pdf-field">
                <span>{t('watermarkPosition')}</span>
                <div className="pdf-tools-position-grid">
                  {WATERMARK_POSITIONS.map((position, index) => (
                    <button
                      key={position}
                      type="button"
                      className={`pdf-tools-position-button${!wm.customPosition && wm.position === position ? ' active' : ''}`}
                      title={`${t('watermarkPosition')} ${index + 1}`}
                      aria-label={`${t('watermarkPosition')} ${index + 1}`}
                      onClick={() => setWm({ ...wm, position, customPosition: null })}
                    >
                      <span className={`pdf-tools-position-dot ${position}`} />
                    </button>
                  ))}
                </div>
              </div>
            )}
            {wm.layout === 'single' && (
              <label className="pdf-field">
                <span>{t('watermarkMargin')}</span>
                <select
                  className="pdf-modal-input pdf-input-sm pdf-wm-margin-select"
                  value={wm.margin}
                  disabled={wm.customPosition !== null}
                  onChange={(event) =>
                    setWm({ ...wm, margin: event.target.value as WatermarkMargin })
                  }
                >
                  {WATERMARK_MARGINS.map((margin) => (
                    <option key={margin.value} value={margin.value}>
                      {margin.label}
                    </option>
                  ))}
                </select>
              </label>
            )}
            {wm.layout === 'single' && wm.customPosition && (
              <div className="pdf-field-grid pdf-wm-coordinate-grid">
                <label className="pdf-field">
                  <span>X</span>
                  <input
                    className="pdf-modal-input pdf-input-sm"
                    type="number"
                    min={0}
                    max={100}
                    value={Math.round(wm.customPosition.xRatio * 100)}
                    onChange={(event) =>
                      setWm({
                        ...wm,
                        customPosition: {
                          ...wm.customPosition!,
                          xRatio: Math.max(0, Math.min(100, Number(event.target.value) || 0)) / 100,
                        },
                      })
                    }
                  />
                  <em>%</em>
                </label>
                <label className="pdf-field">
                  <span>Y</span>
                  <input
                    className="pdf-modal-input pdf-input-sm"
                    type="number"
                    min={0}
                    max={100}
                    value={Math.round(wm.customPosition.yRatio * 100)}
                    onChange={(event) =>
                      setWm({
                        ...wm,
                        customPosition: {
                          ...wm.customPosition!,
                          yRatio: Math.max(0, Math.min(100, Number(event.target.value) || 0)) / 100,
                        },
                      })
                    }
                  />
                  <em>%</em>
                </label>
              </div>
            )}
            {wm.layout === 'tiled' && (
              <div className="pdf-field-grid pdf-wm-spacing-grid">
                <label className="pdf-field">
                  <span>{t('watermarkHorizontalSpacing')}</span>
                  <input
                    className="pdf-modal-input pdf-input-sm"
                    type="number"
                    min={0}
                    max={300}
                    value={wm.horizontalSpacing}
                    onChange={(e) =>
                      setWm({
                        ...wm,
                        horizontalSpacing: Math.min(300, Math.max(0, Number(e.target.value) || 0)),
                      })
                    }
                  />
                </label>
                <label className="pdf-field">
                  <span>{t('watermarkVerticalSpacing')}</span>
                  <input
                    className="pdf-modal-input pdf-input-sm"
                    type="number"
                    min={0}
                    max={300}
                    value={wm.verticalSpacing}
                    onChange={(e) =>
                      setWm({
                        ...wm,
                        verticalSpacing: Math.min(300, Math.max(0, Number(e.target.value) || 0)),
                      })
                    }
                  />
                </label>
              </div>
            )}
            {wm.type === 'text' && (
              <>
                <label className="pdf-field">
                  <span>{t('drawColor')}</span>
                  <span className="pdf-color-row">
                    {WM_COLORS.map((c) => (
                      <button
                        key={c}
                        className={`pdf-color-dot${wm.color === c ? ' active' : ''}`}
                        style={{ background: c }}
                        aria-label={c}
                        onClick={() => setWm({ ...wm, color: c })}
                      />
                    ))}
                    <input
                      className="pdf-tools-color-input"
                      type="color"
                      value={wm.color}
                      aria-label={t('drawColor')}
                      onChange={(event) => setWm({ ...wm, color: event.target.value })}
                    />
                  </span>
                </label>
                <div className="pdf-sign-hint">
                  @date · @date{'{yyyy/MM/dd}'} · @page · @filename · @title · @author · @uuid · @@
                </div>
              </>
            )}
            <div
              className={`pdf-wm-preview${wm.layout === 'tiled' ? ' tiled' : ` ${wm.position}`}`}
              ref={previewRef}
              style={{ color: wm.color }}
            >
              {Array.from({ length: wm.layout === 'tiled' ? 9 : 1 }, (_, index) => (
                <span
                  className="pdf-wm-preview-item"
                  key={index}
                  ref={wm.layout === 'single' ? previewItemRef : undefined}
                  onPointerDown={wm.layout === 'single' ? startPreviewDrag : undefined}
                  onPointerMove={wm.layout === 'single' ? movePreview : undefined}
                  onPointerUp={wm.layout === 'single' ? stopPreviewDrag : undefined}
                  onPointerCancel={wm.layout === 'single' ? stopPreviewDrag : undefined}
                  style={{
                    opacity: Math.max(wm.opacity, 0.25),
                    ...(wm.layout === 'single'
                      ? {
                          left: previewLayout.left,
                          top: previewLayout.top,
                          transform: `translate(-50%, -50%) rotate(${-wm.angle}deg) scale(${previewLayout.scale})`,
                        }
                      : { transform: `rotate(${-wm.angle}deg)` }),
                  }}
                >
                  {wm.type === 'image' && wm.image ? (
                    <img src={`data:image/png;base64,${wm.image}`} alt="" />
                  ) : (
                    previewText ||
                    (wm.type === 'image' ? t('watermarkChooseImage') : t('watermarkPlaceholder'))
                  )}
                </span>
              ))}
            </div>
          </div>
        ) : (
          <>
            <div className="pdf-field-grid">
              {field('headerLeft', t('hfHeaderLeft'))}
              {field('headerCenter', t('hfHeaderCenter'))}
              {field('headerRight', t('hfHeaderRight'))}
              {field('footerLeft', t('hfFooterLeft'))}
              {field('footerCenter', t('hfFooterCenter'))}
              {field('footerRight', t('hfFooterRight'))}
            </div>
            <label className="pdf-field">
              <input
                type="checkbox"
                checked={hf.pageNumber}
                onChange={(e) => setHf({ ...hf, pageNumber: e.target.checked })}
              />
              <span>{t('hfPageNumber')}</span>
            </label>
            {hf.pageNumber && (
              <>
                <label className="pdf-field">
                  <span>{t('hfPageNumberFormat')}</span>
                  <input
                    className="pdf-modal-input"
                    value={hf.pageNumberFormat}
                    placeholder="{page} / {total}"
                    onChange={(event) => setHf({ ...hf, pageNumberFormat: event.target.value })}
                  />
                </label>
                <div className="pdf-field-grid pdf-hf-number-grid">
                  <label className="pdf-field">
                    <span>{t('hfStartAt')}</span>
                    <input
                      className="pdf-modal-input pdf-input-sm"
                      type="number"
                      min={1}
                      value={hf.startAt}
                      onChange={(event) =>
                        setHf({
                          ...hf,
                          startAt: Math.max(1, Number(event.target.value) || 1),
                        })
                      }
                    />
                  </label>
                  <label className="pdf-field">
                    <span>{t('hfZeroPad')}</span>
                    <input
                      className="pdf-modal-input pdf-input-sm"
                      type="number"
                      min={0}
                      max={12}
                      value={hf.pageNumberZeroPad}
                      onChange={(event) =>
                        setHf({
                          ...hf,
                          pageNumberZeroPad: Math.max(
                            0,
                            Math.min(12, Number(event.target.value) || 0),
                          ),
                        })
                      }
                    />
                  </label>
                </div>
                <div className="pdf-field">
                  <span>{t('watermarkPosition')}</span>
                  <div className="pdf-tools-position-grid">
                    {WATERMARK_POSITIONS.map((position, index) => (
                      <button
                        key={position}
                        type="button"
                        className={`pdf-tools-position-button${hf.pageNumberPosition === position ? ' active' : ''}`}
                        title={`${t('watermarkPosition')} ${index + 1}`}
                        aria-label={`${t('watermarkPosition')} ${index + 1}`}
                        onClick={() => setHf({ ...hf, pageNumberPosition: position })}
                      >
                        <span className={`pdf-tools-position-dot ${position}`} />
                      </button>
                    ))}
                  </div>
                </div>
                <label className="pdf-field">
                  <span>{t('watermarkMargin')}</span>
                  <select
                    className="pdf-modal-input pdf-input-sm pdf-wm-margin-select"
                    value={hf.pageNumberMargin}
                    onChange={(event) =>
                      setHf({ ...hf, pageNumberMargin: event.target.value as WatermarkMargin })
                    }
                  >
                    {WATERMARK_MARGINS.map((margin) => (
                      <option key={margin.value} value={margin.value}>
                        {margin.label}
                      </option>
                    ))}
                  </select>
                </label>
              </>
            )}
            <label className="pdf-field">
              <span>{t('watermarkPages')}</span>
              <input
                className={`pdf-modal-input${selectedHeaderFooterPages ? '' : ' invalid'}`}
                value={headerFooterPageRange}
                aria-invalid={!selectedHeaderFooterPages}
                title={t('extractRangeHint', { total: totalPages })}
                placeholder={totalPages > 1 ? `1-${totalPages}` : '1'}
                onChange={(event) => setHeaderFooterPageRange(event.target.value)}
              />
            </label>
            <label className="pdf-field">
              <span>{t('texteditFont')}</span>
              <select
                className="pdf-modal-input pdf-input-sm"
                value={hf.fontFamily}
                onChange={(event) =>
                  setHf({ ...hf, fontFamily: event.target.value as HeaderFooterFont })
                }
              >
                {HEADER_FOOTER_FONTS.map((font) => (
                  <option key={font.value} value={font.value}>
                    {font.label}
                  </option>
                ))}
              </select>
            </label>
            <label className="pdf-field">
              <span>{t('watermarkSize')}</span>
              <input
                type="range"
                min={6}
                max={36}
                value={hf.fontSize}
                onChange={(event) => setHf({ ...hf, fontSize: Number(event.target.value) })}
              />
              <em>{hf.fontSize} pt</em>
            </label>
            <label className="pdf-field">
              <span>{t('drawColor')}</span>
              <span className="pdf-color-row">
                {WM_COLORS.map((color) => (
                  <button
                    key={color}
                    className={`pdf-color-dot${hf.color === color ? ' active' : ''}`}
                    style={{ background: color }}
                    aria-label={color}
                    onClick={() => setHf({ ...hf, color })}
                  />
                ))}
                <input
                  className="pdf-tools-color-input"
                  type="color"
                  value={hf.color}
                  aria-label={t('drawColor')}
                  onChange={(event) => setHf({ ...hf, color: event.target.value })}
                />
              </span>
            </label>
            <div className="pdf-sign-hint">{t('hfTokenHint')}</div>
          </>
        )}

        <div className="pdf-modal-actions">
          <button className="pdf-modal-btn" onClick={onCancel}>
            {t('cancel')}
          </button>
          <button
            className="pdf-modal-btn primary"
            disabled={!canApply}
            onClick={() =>
              onApply(
                wmUsed ? wm : null,
                hfUsed ? hf : null,
                wmUsed ? selectedPages : null,
                hfUsed ? selectedHeaderFooterPages : null,
              )
            }
          >
            {t('ok')}
          </button>
        </div>
      </div>
    </div>
  )
}
