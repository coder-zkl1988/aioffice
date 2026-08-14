import { describe, expect, it } from 'vitest'
import { stampDrawPlacement } from '../src/shared/ipc'

function rotatedCenter(placement: ReturnType<typeof stampDrawPlacement>): [number, number] {
  const radians = (placement.rotation * Math.PI) / 180
  return [
    placement.x + (placement.width * Math.cos(radians) - placement.height * Math.sin(radians)) / 2,
    placement.y + (placement.width * Math.sin(radians) + placement.height * Math.cos(radians)) / 2,
  ]
}

describe('stampDrawPlacement', () => {
  it('preserves the original lower-left placement without rotation', () => {
    expect(stampDrawPlacement([10, 20, 110, 70])).toEqual({
      x: 10,
      y: 20,
      width: 100,
      height: 50,
      rotation: 0,
    })
  })

  it.each([-90, -35, 35, 90, 180])('rotates %s degrees around the rect center', (rotation) => {
    const placement = stampDrawPlacement([10, 20, 110, 70], rotation)
    const center = rotatedCenter(placement)
    expect(center[0]).toBeCloseTo(60, 8)
    expect(center[1]).toBeCloseTo(45, 8)
  })
})
