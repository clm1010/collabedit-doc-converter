import type { DocxArchive } from './zipExtractor.js'
import type { ExtractedImage, RelationshipMap } from '../types/ooxml.js'

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
  webp: 'image/webp',
}

function guessMimeFromBuffer(buf: Uint8Array): string {
  if (buf.length < 4) return ''
  if (buf[0] === 0x89 && buf[1] === 0x50) return 'image/png'
  if (buf[0] === 0xff && buf[1] === 0xd8) return 'image/jpeg'
  if (buf[0] === 0x47 && buf[1] === 0x49) return 'image/gif'
  if (buf[0] === 0x42 && buf[1] === 0x4d) return 'image/bmp'
  if (buf[0] === 0x52 && buf[1] === 0x49 && buf[2] === 0x46 && buf[3] === 0x46) return 'image/webp'
  return ''
}

export function extractImages(archive: DocxArchive): Map<string, ExtractedImage> {
  const images = new Map<string, ExtractedImage>()
  const mediaFiles = archive.listFiles('word/media/')

  for (const filePath of mediaFiles) {
    const data = archive.getBuffer(filePath)
    if (!data || data.length === 0) continue

    const fileName = filePath.split('/').pop()!
    const ext = fileName.split('.').pop()?.toLowerCase() ?? ''
    const bufMime = guessMimeFromBuffer(data)
    const mime = bufMime || MIME_BY_EXT[ext] || 'application/octet-stream'
    const base64 = Buffer.from(data).toString('base64')

    const relPath = filePath.replace(/^word\//, '')
    images.set(relPath, { base64, mime, fileName, relPath })
  }

  return images
}

/** 通过 rId 获取图片的 data URL */
export function getImageDataUrl(
  rId: string,
  rels: RelationshipMap,
  images: Map<string, ExtractedImage>,
): string | null {
  const rel = rels[rId]
  if (!rel) return null

  const target = rel.target
  const img = images.get(target)
  if (!img) return null

  return `data:${img.mime};base64,${img.base64}`
}
