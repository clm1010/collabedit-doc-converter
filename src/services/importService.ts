import JSZip from 'jszip'
import { convert } from './unoClient.js'
import { extractMetadata } from '../utils/metadataExtractor.js'
import { postProcessLoHtml } from '../utils/htmlPostProcess.js'
import type { ImportResult } from '../types/docMetadata.js'

const MIME_BY_EXT: Record<string, string> = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  tiff: 'image/tiff',
  tif: 'image/tiff',
  svg: 'image/svg+xml',
  emf: 'image/x-emf',
  wmf: 'image/x-wmf',
}

function guessMimeFromBuffer(buf: Buffer): string {
  if (buf[0] === 0x89 && buf[1] === 0x50) return 'image/png'
  if (buf[0] === 0xff && buf[1] === 0xd8) return 'image/jpeg'
  if (buf[0] === 0x47 && buf[1] === 0x49) return 'image/gif'
  if (buf[0] === 0x42 && buf[1] === 0x4d) return 'image/bmp'
  return ''
}

interface ExtractedImage {
  base64: string
  mime: string
  fileName: string
}

function naturalSort(a: string, b: string): number {
  const re = /(\d+)/g
  const aParts = a.split(re)
  const bParts = b.split(re)
  for (let i = 0; i < Math.max(aParts.length, bParts.length); i++) {
    const ap = aParts[i] ?? ''
    const bp = bParts[i] ?? ''
    const an = Number(ap)
    const bn = Number(bp)
    if (!isNaN(an) && !isNaN(bn)) {
      if (an !== bn) return an - bn
    } else {
      if (ap !== bp) return ap.localeCompare(bp)
    }
  }
  return 0
}

async function extractImagesFromDocx(fileBuffer: Buffer): Promise<ExtractedImage[]> {
  const images: ExtractedImage[] = []
  try {
    const zip = await JSZip.loadAsync(fileBuffer)
    const mediaFiles = Object.keys(zip.files)
      .filter((f) => f.startsWith('word/media/') && !zip.files[f].dir)
      .sort((a, b) => naturalSort(a, b))

    for (const filePath of mediaFiles) {
      const file = zip.file(filePath)
      if (!file) continue
      const uint8 = await file.async('uint8array')
      const base64 = Buffer.from(uint8).toString('base64')
      const ext = filePath.split('.').pop()?.toLowerCase() ?? ''
      const bufMime = guessMimeFromBuffer(Buffer.from(uint8.buffer, uint8.byteOffset, uint8.byteLength))
      const mime = bufMime || MIME_BY_EXT[ext] || 'image/png'
      const fileName = filePath.split('/').pop()!
      images.push({ base64, mime, fileName })
    }

    console.log(`[import] Extracted ${images.length} images: ${images.map((i) => i.fileName).join(', ')}`)
  } catch (err) {
    console.warn('[import] Failed to extract images from DOCX:', err)
  }
  return images
}

function embedImagesInHtml(html: string, images: ExtractedImage[]): string {
  if (images.length === 0) return html

  const byName = new Map<string, ExtractedImage>()
  const byBase = new Map<string, ExtractedImage>()
  for (const img of images) {
    byName.set(img.fileName.toLowerCase(), img)
    byBase.set(img.fileName.replace(/\.[^.]+$/, '').toLowerCase(), img)
  }

  let orderIdx = 0
  let matched = 0

  const result = html.replace(
    /<img([^>]*)src=(["'])([^"']+)\2([^>]*)>/gi,
    (match, before: string, quote: string, src: string, after: string) => {
      if (src.startsWith('data:')) return match

      const fileName = (src.split('/').pop() ?? '').toLowerCase()
      const fileBase = fileName.replace(/\.[^.]+$/, '')

      let img = byName.get(fileName) || byBase.get(fileBase)

      if (!img) {
        for (const [base, data] of byBase) {
          if (fileBase.includes(base) || base.includes(fileBase)) {
            img = data
            break
          }
        }
      }

      if (!img && orderIdx < images.length) {
        img = images[orderIdx]
        orderIdx++
      }

      if (img) {
        matched++
        return `<img${before}src=${quote}data:${img.mime};base64,${img.base64}${quote}${after}>`
      }

      console.warn(`[import] Image not matched, removed: "${src}"`)
      return ''
    }
  )

  console.log(`[import] Embedded ${matched}/${images.length} images`)
  return result
}

export async function importDocx(fileBuffer: Buffer): Promise<ImportResult> {
  console.log(`[import] Starting import, buffer size: ${fileBuffer.length}`)

  const [metadata, images] = await Promise.all([
    extractMetadata(fileBuffer),
    extractImagesFromDocx(fileBuffer),
  ])

  let html: string
  try {
    const xhtmlBuffer = await convert(fileBuffer, { to: 'html', filter: 'XHTML Writer File' })
    html = xhtmlBuffer.toString('utf-8')
    const hasEmbeddedImages = html.includes('data:image/')
    if (!hasEmbeddedImages && images.length > 0) {
      throw new Error('XHTML filter did not embed images, falling back')
    }
    console.log('[import] XHTML filter succeeded, images auto-embedded')
  } catch (err) {
    console.warn('[import] XHTML filter failed, falling back to HTML + manual embed:', err)
    const htmlBuffer = await convert(fileBuffer, { to: 'html' })
    html = htmlBuffer.toString('utf-8')
    html = embedImagesInHtml(html, images)
  }

  console.log(`[import] Raw HTML length: ${html.length}`)

  html = postProcessLoHtml(html)

  console.log(`[import] Final HTML length: ${html.length}`)
  return { html, metadata }
}
