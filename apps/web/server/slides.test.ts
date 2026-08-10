import { readFile } from 'node:fs/promises'
import { fileURLToPath } from 'node:url'
import PptxGenJS from 'pptxgenjs'
import { describe, expect, it } from 'vitest'
import { SlidesWebService } from './slides'

const fixture = fileURLToPath(
  new URL(
    '../../../packages/pptx-engine/tests/fixtures/01_standard_business.pptx',
    import.meta.url,
  ),
)

describe('SlidesWebService', () => {
  it('opens, edits, saves, and reopens a real presentation', async () => {
    const service = new SlidesWebService()
    const bytes = await readFile(fixture)
    const opened = await service.open({
      name: 'business.pptx',
      fitWidthPx: 1280,
      pptxBase64: bytes.toString('base64'),
    })

    expect(opened.result.slides).toHaveLength(5)
    const added = (await service.call({
      sessionId: opened.sessionId,
      action: 'addElement',
      args: [
        {
          slideIndex: 0,
          kind: 'textbox',
          xPx: 80,
          yPx: 80,
          wPx: 420,
          hPx: 70,
          fitWidthPx: 1280,
          text: 'GENOFFICE_WEB_TEST',
        },
      ],
    })) as { sourceId: string }

    expect(added.sourceId).toBeTruthy()
    await expect(service.call({ sessionId: opened.sessionId, action: 'isDirty' })).resolves.toBe(
      true,
    )

    const saved = await service.save(opened.sessionId, 'business-web.pptx')
    await expect(service.call({ sessionId: opened.sessionId, action: 'isDirty' })).resolves.toBe(
      false,
    )

    const reopened = await service.open({
      name: 'business-web.pptx',
      fitWidthPx: 1280,
      pptxBase64: saved.pptxBase64,
    })
    expect(JSON.stringify(reopened.result.slides)).toContain('GENOFFICE_WEB_TEST')
  })

  it('persists advanced browser editing operations', async () => {
    const service = new SlidesWebService()
    const opened = await service.blank(1280)
    const call = (action: string, op?: unknown) =>
      service.call({ sessionId: opened.sessionId, action, args: op === undefined ? [] : [op] })

    const first = (await call('addElement', {
      slideIndex: 0,
      kind: 'textbox',
      xPx: 80,
      yPx: 60,
      wPx: 300,
      hPx: 50,
      fitWidthPx: 1280,
      text: 'WEB_ADVANCED',
    })) as { sourceId: string }
    const second = (await call('addElement', {
      slideIndex: 0,
      kind: 'rect',
      xPx: 420,
      yPx: 60,
      wPx: 180,
      hPx: 100,
      fitWidthPx: 1280,
      text: 'GROUPED',
      fillColor: '#E9F2FF',
    })) as { sourceId: string }
    const grouped = (await call('groupElements', {
      slideIndex: 0,
      sourceIds: [first.sourceId, second.sourceId],
    })) as { groupId: string }

    expect(grouped.groupId).toBeTruthy()
    await expect(
      call('ungroupElement', { slideIndex: 0, sourceId: grouped.groupId }),
    ).resolves.toBeTruthy()
    await expect(
      call('addTable', {
        slideIndex: 0,
        rows: 2,
        cols: 3,
        xPx: 80,
        yPx: 220,
        wPx: 520,
        hPx: 160,
        fitWidthPx: 1280,
      }),
    ).resolves.toBeTruthy()
    await expect(
      call('addChart', {
        slideIndex: 0,
        kind: 'bar',
        title: 'Web Chart',
        categories: ['A', 'B'],
        series: [{ name: 'S1', values: [4, 7] }],
        xPx: 650,
        yPx: 80,
        wPx: 500,
        hPx: 300,
        fitWidthPx: 1280,
      }),
    ).resolves.toBeTruthy()
    await expect(
      call('addSmartArt', {
        slideIndex: 0,
        layout: 'process',
        items: ['Plan', 'Build', 'Ship'],
        xPx: 100,
        yPx: 430,
        wPx: 900,
        hPx: 150,
        fitWidthPx: 1280,
      }),
    ).resolves.toBeTruthy()

    const saved = await service.save(opened.sessionId, 'advanced-web.pptx')
    const reopened = await service.open({
      name: 'advanced-web.pptx',
      fitWidthPx: 1280,
      pptxBase64: saved.pptxBase64,
    })
    expect(JSON.stringify(reopened.result.slides)).toContain('WEB_ADVANCED')
  })

  it('imports every generated HTML slide and supports grouped child editing', async () => {
    const service = new SlidesWebService()
    const opened = await service.blank(1280)
    const call = (action: string, op?: unknown) =>
      service.call({ sessionId: opened.sessionId, action, args: op === undefined ? [] : [op] })
    const presentation = new PptxGenJS()
    presentation.layout = 'LAYOUT_WIDE'
    presentation.addSlide().addText('PAGE_ONE', { x: 1, y: 1, w: 5, h: 1 })
    presentation.addSlide().addText('PAGE_TWO', { x: 1, y: 1, w: 5, h: 1 })
    const generated = await presentation.write({ outputType: 'arraybuffer' })
    const imported = (await call('importGenerated', {
      pptxBase64: Buffer.from(generated as ArrayBuffer).toString('base64'),
      fitWidthPx: 1280,
      mode: 'replace',
      deckName: 'generated.pptx',
    })) as { slides: unknown[] }

    expect(imported.slides).toHaveLength(2)
    expect(JSON.stringify(imported.slides)).toContain('PAGE_TWO')

    const first = (await call('addElement', {
      slideIndex: 0,
      kind: 'textbox',
      xPx: 80,
      yPx: 80,
      wPx: 260,
      hPx: 60,
      fitWidthPx: 1280,
      text: 'BEFORE_EDIT',
    })) as { sourceId: string }
    const second = (await call('addElement', {
      slideIndex: 0,
      kind: 'rect',
      xPx: 380,
      yPx: 80,
      wPx: 160,
      hPx: 80,
      fitWidthPx: 1280,
      text: 'SECOND',
    })) as { sourceId: string }
    const grouped = (await call('groupElements', {
      slideIndex: 0,
      sourceIds: [first.sourceId, second.sourceId],
    })) as {
      groupId: string
      slide: { nodes: Array<{ sourceId: string; children?: Array<{ sourceId: string }> }> }
    }
    const groupNode = grouped.slide.nodes.find((node) => node.sourceId === grouped.groupId)
    const childSourceId = groupNode?.children?.find((child) =>
      JSON.stringify(child).includes('BEFORE_EDIT'),
    )?.sourceId
    expect(childSourceId).toBeTruthy()
    const edited = await call('editText', {
      slideIndex: 0,
      sourceId: childSourceId,
      groupId: grouped.groupId,
      paragraphs: [{ runs: [{ text: 'GROUP_CHILD_EDITED' }] }],
    })
    expect(JSON.stringify(edited)).toContain('GROUP_CHILD_EDITED')
  })

  it('collapses history batches, restores AI snapshots, and repastes in place', async () => {
    const service = new SlidesWebService()
    const opened = await service.blank(1280)
    const call = (action: string, op?: unknown) =>
      service.call({ sessionId: opened.sessionId, action, args: op === undefined ? [] : [op] })

    await call('beginHistoryBatch')
    for (const [index, text] of ['BATCH_ONE', 'BATCH_TWO'].entries()) {
      await call('addElement', {
        slideIndex: 0,
        kind: 'textbox',
        xPx: 80,
        yPx: 80 + index * 80,
        wPx: 260,
        hPx: 60,
        fitWidthPx: 1280,
        text,
      })
    }
    const snapshotId = (await call('endHistoryBatch')) as number
    expect(snapshotId).toBeTypeOf('number')
    expect(JSON.stringify(await call('undo'))).not.toContain('BATCH_ONE')
    expect(JSON.stringify(await call('redo'))).toContain('BATCH_TWO')
    expect(JSON.stringify(await call('aiSnapshotRestore', snapshotId))).not.toContain('BATCH_TWO')

    await call('copySlide', 0)
    const pasted = (await call('pasteSlide', {
      afterIndex: 0,
      fitWidthPx: 1280,
      mode: 'theme',
    })) as { slides: unknown[] }
    expect(pasted.slides).toHaveLength(2)
    const repasted = (await call('repasteSlide', { fitWidthPx: 1280, mode: 'source' })) as {
      slides: unknown[]
    }
    expect(repasted.slides).toHaveLength(2)
  })

  it('does not mark a failed edit as dirty', async () => {
    const service = new SlidesWebService()
    const opened = await service.blank(1280)
    await service.call({
      sessionId: opened.sessionId,
      action: 'deleteElement',
      args: [{ slideIndex: 0, sourceId: 'missing' }],
    })
    await expect(service.call({ sessionId: opened.sessionId, action: 'isDirty' })).resolves.toBe(
      false,
    )
  })
})
