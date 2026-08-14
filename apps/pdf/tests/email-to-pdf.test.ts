import { describe, expect, it } from 'vitest'
import { prepareEmailMarkup, sanitizeEmailHtml } from '../src/renderer/email-to-pdf'
import type { EmailDocumentLabels } from '../src/renderer/email-to-pdf'
import type { Attachment } from 'postal-mime'

const labels: EmailDocumentLabels = {
  from: 'From',
  to: 'To',
  cc: 'Cc',
  bcc: 'Bcc',
  date: 'Date',
  attachments: 'Attachments',
  emptyBody: 'No body',
  remoteImageRemoved: 'remote image removed',
  untitled: 'Untitled email',
}

const options = {
  includeAttachments: true,
  maxAttachmentSizeMb: 10,
  includeAllRecipients: true,
  labels,
  locale: 'en-US',
}

function emailBytes(value: string): Uint8Array {
  return new TextEncoder().encode(value.replace(/^\n/u, '').replace(/\n/g, '\r\n'))
}

describe('email document preparation', () => {
  it('sanitizes active content and only resolves CID images', () => {
    const inline: Attachment = {
      filename: 'logo.png',
      mimeType: 'image/png',
      disposition: 'inline',
      contentId: '<logo@example>',
      content: new Uint8Array([137, 80, 78, 71]),
    }
    const output = sanitizeEmailHtml(
      '<style>@import "https://tracker.test/x"</style><script>alert(1)</script>' +
        '<p style="color:red;position:fixed;background-image:url(https://tracker.test/x)">Hello</p>' +
        '<img src="cid:logo@example" alt="Logo"><img src="https://tracker.test/pixel" alt="Pixel">' +
        '<iframe src="https://tracker.test/frame"></iframe>',
      [inline],
      labels,
    )

    expect(output).toContain('style="color:red"')
    expect(output).toContain('src="data:image/png;base64,')
    expect(output).toContain('Pixel (remote image removed)')
    expect(output).not.toMatch(/script|iframe|position|background-image|tracker\.test/iu)
  })

  it('renders headers, body and embeds only regular attachments', async () => {
    const prepared = await prepareEmailMarkup(
      emailBytes(`
From: Alice <alice@example.com>
To: Bob <bob@example.com>
Cc: Carol <carol@example.com>
Bcc: Hidden <hidden@example.com>
Date: Thu, 14 Aug 2026 10:15:00 +0800
Subject: =?UTF-8?B?5rWL6K+V6YKu5Lu2?=
MIME-Version: 1.0
Content-Type: multipart/mixed; boundary="mixed"

--mixed
Content-Type: text/html; charset=utf-8

<h2>Hello</h2><img src="https://tracker.test/pixel.png">
--mixed
Content-Type: text/plain; name="../../report.txt"
Content-Disposition: attachment; filename="../../report.txt"
Content-Transfer-Encoding: base64

cmVwb3J0IGNvbnRlbnQ=
--mixed
Content-Type: image/png; name="inline.png"
Content-Disposition: inline; filename="inline.png"
Content-ID: <inline-logo>
Content-Transfer-Encoding: base64

iVBORw0KGgo=
--mixed--
`),
      options,
    )

    expect(prepared.html).toContain('测试邮件')
    expect(prepared.html).toContain('Alice &lt;alice@example.com&gt;')
    expect(prepared.html).toContain('Carol &lt;carol@example.com&gt;')
    expect(prepared.html).toContain('Hidden &lt;hidden@example.com&gt;')
    expect(prepared.html).toContain('remote image removed')
    expect(prepared.html).not.toContain('tracker.test')
    expect(prepared.attachments).toHaveLength(1)
    expect(prepared.attachments[0]?.name).toBe('report.txt')
    expect(new TextDecoder().decode(prepared.attachments[0]?.bytes)).toBe('report content')
  })

  it('omits CC, BCC and binary attachment data when disabled', async () => {
    const prepared = await prepareEmailMarkup(
      emailBytes(`
From: sender@example.com
To: receiver@example.com
Cc: cc@example.com
Bcc: bcc@example.com
Subject: Plain text
Content-Type: text/plain; charset=utf-8

Line one
Line two
`),
      { ...options, includeAttachments: false, includeAllRecipients: false },
    )

    expect(prepared.html).toContain('Line one')
    expect(prepared.html).not.toContain('cc@example.com')
    expect(prepared.html).not.toContain('bcc@example.com')
    expect(prepared.attachments).toEqual([])
  })

  it('rejects empty and oversized input before parsing', async () => {
    await expect(prepareEmailMarkup(new Uint8Array(), options)).rejects.toThrow('empty')
    await expect(prepareEmailMarkup(new Uint8Array(50 * 1024 * 1024 + 1), options)).rejects.toThrow(
      '50 MB',
    )
  })
})
