import { env } from '../config/env.js'

const UNOSERVER_URL = env.unoserverUrl

interface ConvertOptions {
  from?: string
  to: string
  filter?: string
  timeout?: number
}

export async function convert(fileBuffer: Buffer, options: ConvertOptions): Promise<Buffer> {
  const { to, filter, timeout = env.convertTimeout } = options

  const controller = new AbortController()
  const timer = setTimeout(() => controller.abort(), timeout)

  try {
    const formData = new FormData()
    formData.append('file', new Blob([new Uint8Array(fileBuffer)]), 'input')
    formData.append('convert-to', to)
    if (filter) {
      formData.append('filter', filter)
    }

    const response = await fetch(`${UNOSERVER_URL}/request`, {
      method: 'POST',
      body: formData,
      signal: controller.signal,
    })

    if (!response.ok) {
      const text = await response.text().catch(() => '')
      throw new Error(`unoserver conversion failed (${response.status}): ${text}`)
    }

    const arrayBuffer = await response.arrayBuffer()
    return Buffer.from(arrayBuffer)
  } finally {
    clearTimeout(timer)
  }
}

export async function checkHealth(): Promise<boolean> {
  try {
    const controller = new AbortController()
    const timer = setTimeout(() => controller.abort(), 5000)

    const response = await fetch(`${UNOSERVER_URL}/request`, {
      method: 'POST',
      signal: controller.signal,
    })
    clearTimeout(timer)
    return response.status !== 0
  } catch (err: any) {
    if (err?.name === 'AbortError') return false
    if (err?.cause?.code === 'ECONNREFUSED') return false
    return false
  }
}
