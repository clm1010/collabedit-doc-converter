import type { TiptapDoc } from '../types/tiptapJson.js'
import type { DocMetadata } from '../types/docMetadata.js'
import { jsonToDocx } from '../export/jsonToDocx.js'
import { jsonToPdf } from '../export/pdfExporter.js'

export async function exportToDocx(
  content: TiptapDoc,
  metadata?: Partial<DocMetadata>,
): Promise<{ buffer: Buffer; warnings: string[] }> {
  return jsonToDocx(content, metadata)
}

export async function exportToPdf(
  content: TiptapDoc,
  metadata?: Partial<DocMetadata>,
): Promise<{ buffer: Buffer; warnings: string[] }> {
  return jsonToPdf(content, metadata)
}
