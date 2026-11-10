/**
 * Object-level hit testing for the viewer's Select object mode (an Acrobat-lite
 * subset of Stirling-PDF #8325): clicking near a page object picks the image or
 * text run under the pointer so it can be deleted with Backspace/Delete.
 *
 * Geometry is PDF user space (y-up) throughout — the same space as
 * `listPageImages` rects and the search index's text runs — so a display click
 * is converted once with `viewToPdf` and every candidate compares in one space.
 */
import type { PageEntry } from './search'

export type ObjectRect = [number, number, number, number]

/** An existing content-stream image on the clicked page */
export interface SelectableImage {
  kind: 'image'
  /** Viewer page index the delete op is staged on */
  pageIndex: number
  /** Bounds as reported by listPageImages (PDF user space) */
  rect: ObjectRect
  /** Painted after the page's first text run (display only) */
  aboveText?: boolean
}

/** One text run from the search index (per pdf.js text item) */
export interface SelectableTextRun {
  kind: 'text'
  pageIndex: number
  /** Run's layout box in PDF user space; doubles as the erase edit's match key */
  rect: ObjectRect
  /** The run's original text */
  text: string
  /** Font height in PDF units; the erase edit's fontSize */
  fontSize: number
}

export type SelectableObject = SelectableImage | SelectableTextRun

/** How far from an object's edge a click still selects it, in display pixels */
export const OBJECT_HIT_TOLERANCE_PX = 8

/**
 * `OBJECT_HIT_TOLERANCE_PX` in PDF user space at the current zoom. One display
 * pixel is `1 / (scale × userUnit)` PDF units: `pageGeom` divides the stored
 * (userUnit-scaled) page size back down, and layout multiplies by `scale`.
 */
export function objectSelectTolerance(scale: number, userUnit = 1): number {
  if (!(scale > 0)) return Infinity
  return OBJECT_HIT_TOLERANCE_PX / (scale * (userUnit > 0 ? userUnit : 1))
}

/** Distance from a point to an axis-aligned rect; 0 when the point is inside */
export function pointRectDistance(
  x: number,
  y: number,
  rect: readonly [number, number, number, number],
): number {
  const dx = Math.max(rect[0] - x, 0, x - rect[2])
  const dy = Math.max(rect[1] - y, 0, y - rect[3])
  return Math.hypot(dx, dy)
}

export function pointInRect(
  x: number,
  y: number,
  rect: readonly [number, number, number, number],
): boolean {
  return x >= rect[0] && x <= rect[2] && y >= rect[1] && y <= rect[3]
}

/** Axis-aligned overlap of two PDF-space rects (shared edges do not count) */
export function rectsOverlap(
  a: readonly [number, number, number, number],
  b: readonly [number, number, number, number],
): boolean {
  return a[0] < b[2] && b[0] < a[2] && a[1] < b[3] && b[1] < a[3]
}

/**
 * Nearest object whose rect is within `tolerance` of the click point.
 *
 * Image-over-text on overlap: when the pointer is inside an image, that image
 * wins even if a text run's box also covers the point — a picture is painted
 * over the text it overlaps, so clicking it must not pick the caption beneath.
 * A text run right under the pointer still beats an image that is merely near
 * it (outside the pointer), so the preference stays predictable.
 */
export function hitTestObject(
  point: readonly [number, number],
  objects: readonly SelectableObject[],
  tolerance: number,
): SelectableObject | null {
  const [x, y] = point
  let best: SelectableObject | null = null
  let bestD = Infinity
  for (const object of objects) {
    const d = pointRectDistance(x, y, object.rect)
    if (d > tolerance) continue
    if (d < bestD - 1e-9) {
      best = object
      bestD = d
    } else if (Math.abs(d - bestD) <= 1e-9 && object.kind === 'image' && best?.kind === 'text') {
      // Same distance: the image is the more specific target
      best = object
      bestD = d
    }
  }
  if (best?.kind === 'text') {
    const image = objects.find(
      (object) => object.kind === 'image' && pointInRect(x, y, object.rect),
    )
    if (image) return image
  }
  return best
}

/**
 * Drop text runs a pending edit already owns or hides: a run overlapping a
 * pending text edit would stage a second, overlapping erase, and a run covered
 * by a pending opaque above-text image (e.g. the table bridge's rendered
 * workbook) would stage an invisible erase that only surfaces if that image is
 * later removed. All rects are PDF user space.
 */
export function selectableTextRuns(
  runs: readonly SelectableTextRun[],
  pendingTextEditRects: readonly ObjectRect[],
  coveredByImageRects: readonly ObjectRect[],
): SelectableTextRun[] {
  return runs.filter(
    (run) =>
      !pendingTextEditRects.some((rect) => rectsOverlap(run.rect, rect)) &&
      !coveredByImageRects.some((rect) => rectsOverlap(run.rect, rect)),
  )
}

/**
 * Selectable text runs of one page from its search-index entry. Rotated runs are
 * skipped — their axis-aligned box would cover glyphs that are not there — as
 * are blank runs and runs without usable geometry.
 */
export function textRunsForPage(
  pageIndex: number,
  entry: PageEntry | undefined | null,
): SelectableTextRun[] {
  if (!entry) return []
  const runs: SelectableTextRun[] = []
  for (const item of entry.items) {
    if (item.rot) continue
    if (!(item.w > 0) || !(item.h > 0)) continue
    const text = entry.text.slice(item.start, item.end)
    if (!text.trim()) continue
    runs.push({
      kind: 'text',
      pageIndex,
      rect: [item.x, item.y, item.x + item.w, item.y + item.h],
      text,
      fontSize: item.h,
    })
  }
  return runs
}

/**
 * Backspace/Delete must never delete a selected object while the user is typing:
 * an editable control that owns the keyboard event keeps the key. Mirrors the
 * viewer's global `inEditable` guard so the two cannot drift apart.
 */
export function isTextEntryTarget(target: EventTarget | null): boolean {
  const el = target as { tagName?: unknown; isContentEditable?: unknown } | null
  if (!el || typeof el.tagName !== 'string') return false
  return (
    el.tagName === 'INPUT' ||
    el.tagName === 'TEXTAREA' ||
    el.tagName === 'SELECT' ||
    el.isContentEditable === true
  )
}
