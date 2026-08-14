import { describe, expect, it, vi } from 'vitest'
import {
  cameraCaptureSize,
  cameraFailureKind,
  cameraPageFileName,
} from '../src/renderer/camera-capture'

describe('camera capture helpers', () => {
  it('keeps small frames and caps the long edge of large frames', () => {
    expect(cameraCaptureSize(1200, 900)).toEqual({ width: 1200, height: 900 })
    expect(cameraCaptureSize(4000, 3000)).toEqual({ width: 2560, height: 1920 })
    expect(cameraCaptureSize(3000, 4000)).toEqual({ width: 1920, height: 2560 })
  })

  it('rejects invalid frame dimensions', () => {
    expect(() => cameraCaptureSize(0, 100)).toThrow('invalid')
    expect(() => cameraCaptureSize(Number.NaN, 100)).toThrow('invalid')
  })

  it('creates stable local scan filenames', () => {
    expect(cameraPageFileName(3, new Date('2026-08-14T01:02:03.456Z'))).toBe(
      'scan-2026-08-14_01-02-03-456-03.jpg',
    )
  })

  it('maps browser camera failures to user-facing categories', () => {
    const original = navigator.mediaDevices
    Object.defineProperty(navigator, 'mediaDevices', {
      configurable: true,
      value: { getUserMedia: vi.fn() },
    })
    expect(cameraFailureKind(new DOMException('denied', 'NotAllowedError'))).toBe('denied')
    expect(cameraFailureKind(new DOMException('busy', 'NotReadableError'))).toBe('unavailable')
    Object.defineProperty(navigator, 'mediaDevices', { configurable: true, value: original })
  })
})
