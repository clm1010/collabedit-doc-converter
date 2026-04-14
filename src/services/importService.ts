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

async function extractImagesFromDocx(
  fileBuffer: Buffer
): Promise<Map<string, { base64: string; mime: string }>> {
  const imageMap = new Map<string, { base64: string; mime: string }>()
  try {
    const zip = await JSZip.loadAsync(fileBuffer)

    const mediaFiles = Object.keys(zip.files).filter(
      (f) => f.startsWith('word/media/') && !zip.files[f].dir
    )

    await Promise.all(
      mediaFiles.map(async (filePath) => {
        const file = zip.file(filePath)
        if (!file) return
        const base64 = await file.async('base64')
        const ext = filePath.split('.').pop()?.toLowerCase() ?? ''
        const mime = MIME_BY_EXT[ext] || 'image/png'
        const fileName = filePath.split('/').pop()!
        imageMap.set(fileName, { base64, mime })
      })
    )

    console.log(`[import] Extracted ${imageMap.size} images from DOCX`)
  } catch (err) {
    console.warn('[import] Failed to extract images from DOCX:', err)
  }
  return imageMap
}

function embedImagesInHtml(
  html: string,
  imageMap: Map<string, { base64: string; mime: string }>
): string {
  if (imageMap.size === 0) return html

  return html.replace(
    /<img([^>]*)src=(["'])([^"']+)\2([^>]*)>/gi,
    (match, before: string, quote: string, src: string, after: string) => {
      if (src.startsWith('data:')) return match

      const fileName = src.split('/').pop() ?? ''
      const img = imageMap.get(fileName)
      if (img) {
        return `<img${before}src=${quote}data:${img.mime};base64,${img.base64}${quote}${after}>`
      }

      for (const [name, data] of imageMap) {
        const srcBase = fileName.replace(/\.[^.]+$/, '').toLowerCase()
        const nameBase = name.replace(/\.[^.]+$/, '').toLowerCase()
        if (srcBase === nameBase) {
          return `<img${before}src=${quote}data:${data.mime};base64,${data.base64}${quote}${after}>`
        }
      }

      console.warn(`[import] Image not found in DOCX: ${src}`)
      return match
    }
  )
}

export async function importDocx(fileBuffer: Buffer): Promise<ImportResult> {
  const [metadata, htmlBuffer, imageMap] = await Promise.all([
    extractMetadata(fileBuffer),
    convert(fileBuffer, { to: 'html' }),
    extractImagesFromDocx(fileBuffer),
  ])

  const rawHtml = htmlBuffer.toString('utf-8')
  let html = postProcessLoHtml(rawHtml)
  html = embedImagesInHtml(html, imageMap)

  return { html, metadata }
}
