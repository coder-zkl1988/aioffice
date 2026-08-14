import { useEffect, useRef, useState } from 'react'
import type { PointerEvent as ReactPointerEvent, ReactElement } from 'react'
import type { PdfRedactionArea } from '@genoffice/pdf-tools'
import { AnnotationMode, getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs'
import type { PDFDocumentLoadingTask, RenderTask } from 'pdfjs-dist'

type Corner = 'nw' | 'ne' | 'se' | 'sw'
type Interaction = {
  type: 'move' | 'resize' | 'draw'
  corner?: Corner
  startX: number
  startY: number
  area: PdfRedactionArea
}

const MIN_SIZE = 0.01

function clamp(value: number, minimum: number, maximum: number): number {
  return Math.min(maximum, Math.max(minimum, value))
}

export function moveRedactionArea(
  area: PdfRedactionArea,
  deltaX: number,
  deltaY: number,
): PdfRedactionArea {
  return {
    ...area,
    x: clamp(area.x + deltaX, 0, 1 - area.width),
    y: clamp(area.y + deltaY, 0, 1 - area.height),
  }
}

export function resizeRedactionArea(
  area: PdfRedactionArea,
  corner: Corner,
  x: number,
  y: number,
): PdfRedactionArea {
  const right = area.x + area.width
  const bottom = area.y + area.height
  const left = corner === 'nw' || corner === 'sw' ? clamp(x, 0, right - MIN_SIZE) : area.x
  const top = corner === 'nw' || corner === 'ne' ? clamp(y, 0, bottom - MIN_SIZE) : area.y
  const nextRight = corner === 'ne' || corner === 'se' ? clamp(x, area.x + MIN_SIZE, 1) : right
  const nextBottom = corner === 'sw' || corner === 'se' ? clamp(y, area.y + MIN_SIZE, 1) : bottom
  return { ...area, x: left, y: top, width: nextRight - left, height: nextBottom - top }
}

export function drawRedactionArea(
  pageIndex: number,
  startX: number,
  startY: number,
  x: number,
  y: number,
): PdfRedactionArea {
  const left = clamp(Math.min(startX, x), 0, 1 - MIN_SIZE)
  const top = clamp(Math.min(startY, y), 0, 1 - MIN_SIZE)
  return {
    pageIndex,
    x: left,
    y: top,
    width: Math.max(MIN_SIZE, clamp(Math.max(startX, x), MIN_SIZE, 1) - left),
    height: Math.max(MIN_SIZE, clamp(Math.max(startY, y), MIN_SIZE, 1) - top),
  }
}

export function RedactionAreaPreview({
  filePath,
  areas,
  selectedIndex,
  color,
  label,
  resizeLabel,
  disabled,
  onChange,
}: {
  filePath: string
  areas: PdfRedactionArea[]
  selectedIndex: number
  color: string
  label: string
  resizeLabel: string
  disabled: boolean
  onChange: (area: PdfRedactionArea) => void
}): ReactElement {
  const area = areas[selectedIndex]
  const pageIndex = area?.pageIndex
  const canvasRef = useRef<HTMLCanvasElement>(null)
  const interactionRef = useRef<Interaction | null>(null)
  const [size, setSize] = useState({ width: 612, height: 792 })
  const [error, setError] = useState('')

  useEffect(() => {
    if (pageIndex === undefined) return
    let canceled = false
    let loadingTask: PDFDocumentLoadingTask | null = null
    let renderTask: RenderTask | null = null
    void (async () => {
      try {
        const data = await window.pdfApi.readFile(filePath)
        const assetBase = new URL('pdfjs/', document.baseURI).href
        loadingTask = getDocument({
          data: new Uint8Array(data),
          cMapUrl: `${assetBase}cmaps/`,
          cMapPacked: true,
          standardFontDataUrl: `${assetBase}standard_fonts/`,
          wasmUrl: `${assetBase}wasm/`,
        })
        const source = await loadingTask.promise
        if (canceled || pageIndex >= source.numPages) return
        const page = await source.getPage(pageIndex + 1)
        const baseViewport = page.getViewport({ scale: 1 })
        const scale = Math.min(1.5, 720 / Math.max(baseViewport.width, baseViewport.height))
        const viewport = page.getViewport({ scale })
        const canvas = canvasRef.current
        if (!canvas || canceled) return
        canvas.width = Math.max(1, Math.round(viewport.width))
        canvas.height = Math.max(1, Math.round(viewport.height))
        setSize({ width: viewport.width, height: viewport.height })
        const context = canvas.getContext('2d')
        if (!context) return
        renderTask = page.render({
          canvas,
          viewport,
          intent: 'display',
          annotationMode: AnnotationMode.ENABLE,
          background: '#ffffff',
        })
        await renderTask.promise
        if (!canceled) setError('')
      } catch (previewError) {
        if (
          !canceled &&
          (previewError as { name?: string }).name !== 'RenderingCancelledException'
        ) {
          setError(previewError instanceof Error ? previewError.message : String(previewError))
        }
      }
    })()
    return () => {
      canceled = true
      renderTask?.cancel()
      void loadingTask?.destroy()
    }
  }, [filePath, pageIndex])

  if (!area) return <></>

  const point = (event: ReactPointerEvent<HTMLElement>): readonly [number, number] => {
    const bounds = event.currentTarget.closest('.pdf-redaction-preview')!.getBoundingClientRect()
    return [
      clamp((event.clientX - bounds.left) / bounds.width, 0, 1),
      clamp((event.clientY - bounds.top) / bounds.height, 0, 1),
    ]
  }

  const start = (
    event: ReactPointerEvent<HTMLElement>,
    type: Interaction['type'],
    corner?: Corner,
  ) => {
    if (disabled || event.button !== 0) return
    event.preventDefault()
    event.stopPropagation()
    event.currentTarget.setPointerCapture(event.pointerId)
    const [startX, startY] = point(event)
    interactionRef.current = { type, corner, startX, startY, area }
    if (type === 'draw') {
      onChange(drawRedactionArea(area.pageIndex, startX, startY, startX, startY))
    }
  }

  const move = (event: ReactPointerEvent<HTMLElement>) => {
    const interaction = interactionRef.current
    if (!interaction) return
    const [x, y] = point(event)
    if (interaction.type === 'move') {
      onChange(moveRedactionArea(interaction.area, x - interaction.startX, y - interaction.startY))
    } else if (interaction.type === 'resize' && interaction.corner) {
      onChange(resizeRedactionArea(interaction.area, interaction.corner, x, y))
    } else {
      onChange(drawRedactionArea(area.pageIndex, interaction.startX, interaction.startY, x, y))
    }
  }

  const finish = () => {
    interactionRef.current = null
  }

  return (
    <div
      className="pdf-redaction-preview"
      style={{ aspectRatio: `${size.width} / ${size.height}` }}
      role="group"
      aria-label={label}
      onPointerDown={(event) => start(event, 'draw')}
      onPointerMove={move}
      onPointerUp={finish}
      onPointerCancel={finish}
    >
      <canvas ref={canvasRef} aria-hidden="true" />
      {error && <span className="pdf-redaction-preview-error">{error}</span>}
      {areas.map(
        (candidate, index) =>
          candidate.pageIndex === area.pageIndex && (
            <div
              className={`pdf-redaction-preview-area${index === selectedIndex ? ' selected' : ''}`}
              style={{
                left: `${candidate.x * 100}%`,
                top: `${candidate.y * 100}%`,
                width: `${candidate.width * 100}%`,
                height: `${candidate.height * 100}%`,
                backgroundColor: color,
              }}
              aria-hidden={index !== selectedIndex}
              key={index}
              onPointerDown={index === selectedIndex ? (event) => start(event, 'move') : undefined}
            >
              {index === selectedIndex &&
                (['nw', 'ne', 'se', 'sw'] as const).map((corner) => (
                  <button
                    className={`pdf-redaction-preview-handle ${corner}`}
                    type="button"
                    aria-label={resizeLabel}
                    disabled={disabled}
                    key={corner}
                    onPointerDown={(event) => start(event, 'resize', corner)}
                  />
                ))}
            </div>
          ),
      )}
    </div>
  )
}
