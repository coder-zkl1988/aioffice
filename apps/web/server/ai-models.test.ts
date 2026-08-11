import { describe, expect, it } from 'vitest'
import { normalizeProviderModels } from './ai-models'

describe('normalizeProviderModels', () => {
  it('normalizes OpenAI-compatible model lists', () => {
    expect(
      normalizeProviderModels({
        data: [{ id: 'gpt-5' }, { id: 'gpt-4.1' }, { id: 'gpt-5' }, { id: '' }],
      }),
    ).toEqual(['gpt-4.1', 'gpt-5'])
  })

  it('accepts common local gateway model shapes', () => {
    expect(
      normalizeProviderModels({
        models: [{ name: 'qwen3:32b' }, { model: 'deepseek-r1:70b' }, 'glm-4.5'],
      }),
    ).toEqual(['deepseek-r1:70b', 'glm-4.5', 'qwen3:32b'])
  })

  it('ignores invalid and oversized identifiers', () => {
    expect(normalizeProviderModels({ data: [{ id: 123 }, { id: 'x'.repeat(257) }, null] })).toEqual(
      [],
    )
  })
})
