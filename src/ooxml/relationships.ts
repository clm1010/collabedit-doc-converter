import type { RelationshipMap } from '../types/ooxml.js'
import { parseXml, ensureArray } from './xmlParser.js'
import type { DocxArchive } from './zipExtractor.js'

export function parseRelationships(archive: DocxArchive, relsPath: string): RelationshipMap {
  const xml = archive.getText(relsPath)
  if (!xml) return {}

  const parsed = parseXml(xml)
  const rels = parsed['Relationships'] as Record<string, unknown> | undefined
  if (!rels) return {}

  const items = ensureArray(rels['Relationship'] as Record<string, unknown>[])
  const map: RelationshipMap = {}

  for (const item of items) {
    const id = item['@_Id'] as string
    const target = item['@_Target'] as string
    const type = item['@_Type'] as string
    if (id && target) {
      map[id] = { target, type: type ?? '' }
    }
  }

  return map
}

export function parseDocumentRelationships(archive: DocxArchive): RelationshipMap {
  return parseRelationships(archive, 'word/_rels/document.xml.rels')
}

export function resolveRelTarget(rels: RelationshipMap, rId: string): string | undefined {
  return rels[rId]?.target
}

const REL_TYPE = {
  hyperlink: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
  image: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
  header: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header',
  footer: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer',
  numbering: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering',
  styles: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
  theme: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
} as const

export { REL_TYPE }
