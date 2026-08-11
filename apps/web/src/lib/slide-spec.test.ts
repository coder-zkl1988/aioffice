import { describe, expect, it } from 'vitest'
import {
  buildWebSlideSpecPrompts,
  createSlideSpecMarker,
  parseGeneratedSlideSpec,
  parseSlideSpecMarker,
} from './slide-spec'

describe('editable slide spec', () => {
  it('prompts the model for native editable elements and embeds the visual skill', () => {
    const prompts = buildWebSlideSpecPrompts({
      brief: 'Compare the two launch options.',
      title: 'Option A reaches market earlier',
      styleSkill: 'White background with a red accent.',
      deckContext: { page_index: 2, total_pages: 6 },
      images: [{ url: 'https://images.example.com/product.jpg', caption: 'Product' }],
      width: 1280,
      height: 720,
    })

    expect(prompts.system).toContain('GenOffice Native Visual Skill v1')
    expect(prompts.system).toContain('Return JSON only')
    expect(prompts.system).toContain('text stays text')
    expect(prompts.user).toContain('https://images.example.com/product.jpg')
    expect(prompts.user).toContain('Option A reaches market earlier')
  })

  it('normalizes a model response and round-trips it through a marker', () => {
    const request = {
      brief: 'Explain growth.',
      title: 'Growth accelerated',
      images: [{ url: 'https://images.example.com/growth.jpg' }],
    }
    const spec = parseGeneratedSlideSpec(
      `\`\`\`json
      {
        "version": 1,
        "title": "Growth accelerated",
        "layout": "split_visual",
        "background": "#ffffff",
        "elements": [
          {"kind":"shape","shape":"roundRect","x":64,"y":160,"w":500,"h":420,"fill":"#F2F5FA","z":1},
          {"kind":"image","url":"https://images.example.com/growth.jpg","x":680,"y":130,"w":520,"h":500,"fit":"cover","z":2},
          {"kind":"text","text":"Growth accelerated","x":64,"y":50,"w":900,"h":80,"fontSize":44,"color":"#171717","bold":true,"z":10},
          {"kind":"text","text":"Revenue increased while acquisition costs declined.","x":96,"y":220,"w":420,"h":150,"fontSize":24,"color":"#333333","z":10}
        ]
      }
      \`\`\``,
      request,
    )

    expect(spec.background).toBe('#FFFFFF')
    expect(spec.elements).toHaveLength(4)
    expect(spec.elements[2]).toMatchObject({ kind: 'text', fontSize: 44, bold: true })
    expect(parseSlideSpecMarker(createSlideSpecMarker(spec))).toEqual(spec)
  })

  it('rejects invented image URLs and slides without editable content', () => {
    expect(() =>
      parseGeneratedSlideSpec(
        JSON.stringify({
          version: 1,
          elements: [
            {
              kind: 'image',
              url: 'https://bad.example.com/invented.jpg',
              x: 0,
              y: 0,
              w: 1280,
              h: 720,
            },
          ],
        }),
        { brief: 'A slide', images: [] },
      ),
    ).toThrow('图片 URL 未获批准')
  })
})
