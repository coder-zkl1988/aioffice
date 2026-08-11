import { defaultAiSettings } from '@genoffice/ai-provider'
import { describe, expect, it } from 'vitest'
import { hasConfiguredWebAi } from './ai'

describe('hasConfiguredWebAi', () => {
  it('enables Web AI capabilities only for a complete custom provider configuration', () => {
    const settings = defaultAiSettings()
    expect(hasConfiguredWebAi(settings)).toBe(false)

    settings.provider = 'custom'
    settings.providers.custom = {
      apiKey: 'test-key',
      model: 'test-model',
      baseUrl: 'https://ai.example.com/v1',
    }
    expect(hasConfiguredWebAi(settings)).toBe(true)

    settings.providers.custom.model = ''
    expect(hasConfiguredWebAi(settings)).toBe(false)
  })
})
