import { lookup } from 'node:dns/promises'
import { BlockList, isIP } from 'node:net'

const blockedAddresses = new BlockList()

for (const [network, prefix] of [
  ['0.0.0.0', 8],
  ['10.0.0.0', 8],
  ['100.64.0.0', 10],
  ['127.0.0.0', 8],
  ['169.254.0.0', 16],
  ['172.16.0.0', 12],
  ['192.0.0.0', 24],
  ['192.0.2.0', 24],
  ['192.168.0.0', 16],
  ['198.18.0.0', 15],
  ['198.51.100.0', 24],
  ['203.0.113.0', 24],
  ['224.0.0.0', 4],
  ['240.0.0.0', 4],
] as const) {
  blockedAddresses.addSubnet(network, prefix, 'ipv4')
}

for (const [network, prefix] of [
  ['::', 128],
  ['::1', 128],
  ['64:ff9b::', 96],
  ['100::', 64],
  ['2001:db8::', 32],
  ['fc00::', 7],
  ['fe80::', 10],
  ['ff00::', 8],
] as const) {
  blockedAddresses.addSubnet(network, prefix, 'ipv6')
}

function isBlockedAddress(address: string): boolean {
  const family = isIP(address)
  if (family === 4) return blockedAddresses.check(address, 'ipv4')
  if (family === 6) {
    const mapped = /^::ffff:(\d+\.\d+\.\d+\.\d+)$/i.exec(address)?.[1]
    return mapped ? blockedAddresses.check(mapped, 'ipv4') : blockedAddresses.check(address, 'ipv6')
  }
  return true
}

async function validatePublicUrl(
  value: unknown,
  options: { allowQuery: boolean; label: string },
): Promise<URL> {
  if (typeof value !== 'string' || value.length > 4096) throw new Error(`${options.label} 无效`)

  let url: URL
  try {
    url = new URL(value)
  } catch {
    throw new Error(`${options.label} 格式不正确`)
  }

  const allowHttp = process.env.AI_ALLOW_HTTP === 'true'
  const allowPrivateNetwork = process.env.AI_ALLOW_PRIVATE_NETWORK === 'true'
  if (url.protocol !== 'https:' && !(allowHttp && url.protocol === 'http:')) {
    throw new Error(`${options.label} 必须使用 HTTPS`)
  }
  if (url.username || url.password) throw new Error(`${options.label} 不能包含用户名或密码`)
  if ((!options.allowQuery && url.search) || url.hash) {
    throw new Error(`${options.label} 包含不允许的查询参数或锚点`)
  }

  const hostname = url.hostname.replace(/^\[|\]$/g, '').toLowerCase()
  if (!hostname || hostname === 'localhost' || hostname.endsWith('.localhost')) {
    throw new Error(`${options.label} 不能指向本机或内网`)
  }

  if (allowPrivateNetwork) return url

  const literalFamily = isIP(hostname)
  if (literalFamily) {
    if (isBlockedAddress(hostname)) throw new Error(`${options.label} 不能指向本机或内网`)
  } else {
    const addresses = await lookup(hostname, { all: true, verbatim: true })
    if (addresses.length === 0 || addresses.some(({ address }) => isBlockedAddress(address))) {
      throw new Error(`${options.label} 不能解析到本机或内网`)
    }
  }

  return url
}

export async function validateProviderBaseUrl(value: unknown): Promise<string> {
  const url = await validatePublicUrl(value, { allowQuery: false, label: 'Base URL' })
  return url.toString().replace(/\/$/, '')
}

export async function validatePublicResourceUrl(value: unknown): Promise<string> {
  return (await validatePublicUrl(value, { allowQuery: true, label: '资源 URL' })).toString()
}
