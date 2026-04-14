import JSZip from 'jszip'
import { convert } from './unoClient.js'
import { enhanceHtmlForExport } from '../utils/htmlExportEnhance.js'
import { validateAndProcessImages } from '../utils/imageProcessor.js'
import type { DocMetadata } from '../types/docMetadata.js'

export async function exportToDocx(
  html: string,
  metadata?: Partial<DocMetadata>
): Promise<{ buffer: Buffer; warnings: string[] }> {
  const { html: validatedHtml, warnings } = validateAndProcessImages(html)
  const enhancedHtml = enhanceHtmlForExport(validatedHtml, metadata)

  const htmlBuffer = Buffer.from(enhancedHtml, 'utf-8')
  let docxBuffer = await convert(htmlBuffer, { to: 'docx' })

  if (metadata?.headers || metadata?.footers) {
    docxBuffer = await injectHeaderFooter(docxBuffer, metadata)
  }

  return { buffer: docxBuffer, warnings }
}

export async function exportToPdf(
  html: string,
  metadata?: Partial<DocMetadata>
): Promise<{ buffer: Buffer; warnings: string[] }> {
  const { html: validatedHtml, warnings } = validateAndProcessImages(html)
  const enhancedHtml = enhanceHtmlForExport(validatedHtml, metadata)

  const htmlBuffer = Buffer.from(enhancedHtml, 'utf-8')
  const pdfBuffer = await convert(htmlBuffer, { to: 'pdf' })

  return { buffer: pdfBuffer, warnings }
}

async function injectHeaderFooter(
  docxBuffer: Buffer,
  metadata: Partial<DocMetadata>
): Promise<Buffer> {
  try {
    const zip = await JSZip.loadAsync(docxBuffer)

    if (metadata.headers?.default) {
      zip.file('word/header1.xml', wrapHeaderXml(metadata.headers.default))
      await ensureRelationship(
        zip,
        'header1.xml',
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header'
      )
    }

    if (metadata.footers?.default) {
      zip.file('word/footer1.xml', wrapFooterXml(metadata.footers.default))
      await ensureRelationship(
        zip,
        'footer1.xml',
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer'
      )
    }

    return await zip.generateAsync({ type: 'nodebuffer' })
  } catch (err) {
    console.warn('Header/footer injection failed, returning original:', err)
    return docxBuffer
  }
}

function wrapHeaderXml(content: string): string {
  if (content.includes('<?xml')) return content
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
${content}
</w:hdr>`
}

function wrapFooterXml(content: string): string {
  if (content.includes('<?xml')) return content
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
${content}
</w:ftr>`
}

async function ensureRelationship(zip: JSZip, filename: string, type: string) {
  const relsPath = 'word/_rels/document.xml.rels'
  const relsFile = zip.file(relsPath)
  if (!relsFile) return

  let relsXml = await relsFile.async('text')
  if (relsXml.includes(filename)) return

  const idMatches = relsXml.match(/Id="rId(\d+)"/g)
  const maxId = idMatches
    ? Math.max(...idMatches.map((m) => Number(m.match(/\d+/)?.[0] ?? 0)))
    : 0
  const newId = `rId${maxId + 1}`

  relsXml = relsXml.replace(
    '</Relationships>',
    `  <Relationship Id="${newId}" Type="${type}" Target="${filename}"/>\n</Relationships>`
  )
  zip.file(relsPath, relsXml)
}
