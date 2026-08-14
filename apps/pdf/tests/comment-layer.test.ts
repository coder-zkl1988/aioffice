import { describe, expect, it } from 'vitest'
import { textCommentsFromAnnotations } from '../src/renderer/CommentLayer'

describe('textCommentsFromAnnotations', () => {
  it('maps standard PDF text annotations for the viewer layer', () => {
    expect(
      textCommentsFromAnnotations([
        {
          id: '7R',
          subtype: 'Text',
          rect: [72, 642, 94, 664],
          color: new Uint8ClampedArray([255, 242, 102]),
          contentsObj: { str: '请核对这个敏感编号。' },
          titleObj: { str: '杨科' },
        },
      ]),
    ).toEqual([
      {
        id: '7R',
        rect: [72, 642, 94, 664],
        color: 'rgb(255 242 102)',
        text: '请核对这个敏感编号。',
        author: '杨科',
      },
    ])
  })

  it('ignores non-text and invalid annotations', () => {
    expect(
      textCommentsFromAnnotations([
        { subtype: 'Link', rect: [0, 0, 10, 10] },
        { subtype: 'Text', rect: [0, Number.NaN, 10, 10] },
      ]),
    ).toEqual([])
  })
})
