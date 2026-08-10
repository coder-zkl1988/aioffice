import { describe, expect, it } from 'vitest'
import { validateProviderBaseUrl } from './security'

describe('validateProviderBaseUrl', () => {
  it('accepts a public HTTPS API URL', async () => {
    await expect(validateProviderBaseUrl('https://api.openai.com/v1')).resolves.toBe(
      'https://api.openai.com/v1',
    )
  })

  it.each([
    'http://api.openai.com/v1',
    'https://localhost/v1',
    'https://127.0.0.1/v1',
    'https://192.168.1.10/v1',
    'https://169.254.169.254/latest',
    'https://[::1]/v1',
  ])('rejects unsafe URL %s', async (url) => {
    await expect(validateProviderBaseUrl(url)).rejects.toThrow()
  })
})
