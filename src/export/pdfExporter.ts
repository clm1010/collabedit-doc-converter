import type { TiptapDoc } from '../types/tiptapJson.js'
import type { DocMetadata } from '../types/docMetadata.js'
import { jsonToDocx } from './jsonToDocx.js'
import { convert } from '../services/unoClient.js'

/**
 * Tiptap JSON → PDF via DOCX intermediate + LibreOffice
 */
export async function jsonToPdf(
  doc: TiptapDoc,
  metadata?: Partial<DocMetadata>,
): Promise<{ buffer: Buffer; warnings: string[] }> {
  const docxResult = await jsonToDocx(doc, metadata)

  const pdfBuffer = await convert(docxResult.buffer, { to: 'pdf' })

  return {
    buffer: pdfBuffer,
    warnings: [
      ...docxResult.warnings,
    ],
  }
}
