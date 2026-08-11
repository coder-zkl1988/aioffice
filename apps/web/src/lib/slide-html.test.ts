// @vitest-environment jsdom

import { describe, expect, it } from 'vitest'
import { buildWebSlideHtmlPrompts, sanitizeGeneratedSlideHtml } from './slide-html'

describe('buildWebSlideHtmlPrompts', () => {
  it('passes the page content, style, dimensions, and approved images to the model', () => {
    const prompts = buildWebSlideHtmlPrompts({
      brief: 'Explain the launch milestones without inventing dates.',
      title: 'Launch plan',
      styleSkill: 'White background, black type, red accent.',
      deckContext: { page_index: 2, total_pages: 8 },
      images: [{ url: 'https://images.example.com/launch.jpg', caption: 'Product photo' }],
      width: 1600,
      height: 900,
    })

    expect(prompts.system).toContain('1600x900px')
    expect(prompts.system).toContain('Return only the complete HTML document')
    expect(prompts.user).toContain('Launch plan')
    expect(prompts.user).toContain('White background, black type, red accent.')
    expect(prompts.user).toContain('https://images.example.com/launch.jpg')
    expect(prompts.user).toContain('"page_index": 2')
  })
})

describe('sanitizeGeneratedSlideHtml', () => {
  it('unwraps fenced HTML, removes active content, and fixes the canvas size', () => {
    const html = sanitizeGeneratedSlideHtml(
      `Here is the slide:\n\`\`\`html
      <html><head>
        <style>@import url('https://bad.example/style.css'); .hero { background:url('https://bad.example/tracker.png') }</style>
        <script>alert('no')</script>
      </head><body onload="alert('no')">
        <h1 style="background-image:url('https://bad.example/other.png')">Launch plan</h1>
        <img src="https://images.example.com/launch.jpg" onerror="alert('no')">
        <img src="https://bad.example/unapproved.jpg">
        <svg><image xlink:href="https://bad.example/vector.svg"></image></svg>
        <video poster="https://bad.example/poster.jpg"><source src="https://bad.example/video.mp4"></video>
        <iframe src="https://bad.example"></iframe>
      </body></html>
      \`\`\``,
      {
        width: 1600,
        height: 900,
        images: [{ url: 'https://images.example.com/launch.jpg' }],
      },
    )

    expect(html).toContain('<!doctype html>')
    expect(html).toContain('width:1600px!important')
    expect(html).toContain('height:900px!important')
    expect(html).toContain('src="https://images.example.com/launch.jpg"')
    expect(html).not.toContain('<script')
    expect(html).not.toContain('<iframe')
    expect(html).not.toContain('onload=')
    expect(html).not.toContain('onerror=')
    expect(html).not.toContain('https://bad.example')
  })

  it('accepts an HTML fragment and rejects an empty response', () => {
    expect(sanitizeGeneratedSlideHtml('<main><h1>Quarterly review</h1></main>')).toContain(
      'Quarterly review',
    )
    expect(() => sanitizeGeneratedSlideHtml('')).toThrow('模型未返回幻灯片 HTML')
  })
})
